"""
Lecture d'un Bulletin de la cote de la Bourse de Casablanca (PDF).

Extrait trois tableaux :
  - cotation    : une ligne par valeur (cours, plus haut/bas, extrêmes annuels...)
  - physionomie : une ligne par valeur et par niveau de prix (marché central uniquement)
  - indices     : MASI et indices sectoriels
et vérifie que la somme de la physionomie = total officiel "ACTIONS".
"""
import re
from datetime import datetime

import pandas as pd
import pdfplumber

NUM_RE = re.compile(r"^-?\d{1,3}(?:[ \u00a0]\d{3})*(?:,\d+)?$")
ISIN_RE = re.compile(r"^MA\d{10}$")
DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{2}$")


def to_num(txt):
    """'1 589 226,50' -> 1589226.5 ; '-' ou vide -> None"""
    if txt is None:
        return None
    t = str(txt).strip().replace("\u00a0", "").replace(" ", "").replace("%", "").replace(",", ".")
    if t in ("", "-"):
        return None
    try:
        return float(t)
    except ValueError:
        return None


def is_num(txt):
    return bool(NUM_RE.match(str(txt).strip()))


def _lines(page):
    """Regroupe les mots d'une page par ligne (même hauteur)."""
    words = sorted(page.extract_words(x_tolerance=1.5, keep_blank_chars=True), key=lambda w: w["top"])
    groups, current, top = [], [], None
    for w in words:
        if top is None or w["top"] - top > 2.5:  # nouvelle ligne si écart vertical > 2,5 pt
            if current:
                groups.append(current)
            current, top = [], w["top"]
        current.append(w)
    if current:
        groups.append(current)
    out = []
    for ws in groups:
        ws = sorted(ws, key=lambda w: w["x0"])
        cells = []
        for w in ws:
            # une "cellule" peut contenir du remplissage : on coupe sur 2 espaces ou plus
            cells += [c.strip() for c in re.split(r"\s{2,}", w["text"]) if c.strip()]
        if cells:
            out.append(cells)
    return out


def parse_bulletin(pdf_path_or_file):
    """Retourne un dict : date, cotation (DataFrame), physionomie (DataFrame),
    indices (DataFrame), totaux (dict), controle (dict)."""
    cotation, physio, indices = [], [], []
    totaux = {}
    seance = None
    segment = None       # 'CENTRAL' ou 'TRANSFERTS'
    in_physio = False

    with pdfplumber.open(pdf_path_or_file) as pdf:
        for page in pdf.pages:
            head = (page.extract_text() or "")[:400]
            page_actions = "MARCHE PRINCIPAL - ACTIONS" in head.replace("  ", " ")
            for cells in _lines(page):
                first = cells[0]

                # Date de séance
                if seance is None:
                    m = re.search(r"Séance du (\d{2}/\d{2}/\d{4})", " ".join(cells))
                    if m:
                        seance = datetime.strptime(m.group(1), "%d/%m/%Y").date()

                # Repères de sections
                if first.startswith("PHYSIONOMIE DES TRANSACTIONS"):
                    in_physio = True
                    continue
                if first.startswith("MARCHE CENTRAL"):
                    segment = "CENTRAL"
                if first.startswith("TRANSFERTS"):
                    segment = "TRANSFERTS"
                if first.startswith("TOTAL GENERAL") or first.startswith("DROITS"):
                    in_physio = False
                    continue

                # ---- Physionomie : Nom | Prix | Quantité | Volume | Contrats
                if in_physio and segment == "CENTRAL":
                    if len(cells) == 4 and first in ("ACTIONS", "MARCHE CENTRAL") and all(is_num(c) for c in cells[1:]):
                        totaux.setdefault(first, (to_num(cells[1]), to_num(cells[2]), to_num(cells[3])))
                        continue
                    if len(cells) == 5 and all(is_num(c) for c in cells[1:]):
                        physio.append({
                            "Valeur": first,
                            "Prix": to_num(cells[1]),
                            "Quantite": to_num(cells[2]),
                            "Volume_MAD": to_num(cells[3]),
                            "Contrats": int(to_num(cells[4])),
                        })
                    continue

                # ---- Cotation : ligne qui contient un code ISIN
                isin_idx = next((i for i, c in enumerate(cells) if ISIN_RE.match(c)), None)
                if page_actions and isin_idx is not None and isin_idx >= 3 and len(cells) >= isin_idx + 6:
                    before, after = cells[:isin_idx], cells[isin_idx + 1:]
                    sector = before[2] if len(before) > 2 else None
                    ref = to_num(before[-3]) if len(before) >= 3 and DATE_RE.match(before[-2]) else None
                    name = after[0]
                    vals = after[1:]
                    statut_idx = next((i for i, c in enumerate(vals) if c in ("T", "NT", "S")), None)
                    if statut_idx is None:
                        continue
                    jour = vals[:statut_idx] + [None] * (4 - len(vals[:statut_idx]))
                    reste = vals[statut_idx + 1:]
                    reste = reste + [None] * (4 - len(reste))
                    cotation.append({
                        "ISIN": cells[isin_idx],
                        "Valeur": name,
                        "Secteur": sector,
                        "Statut": vals[statut_idx],
                        "Cours_ref": ref,
                        "Ouverture": to_num(jour[0]),
                        "Cloture": to_num(jour[1]),
                        "Plus_haut": to_num(jour[2]),
                        "Plus_bas": to_num(jour[3]),
                        "Demande": to_num(reste[0]),
                        "Offre": to_num(reste[1]),
                        "Plus_haut_annuel": to_num(reste[2]),
                        "Plus_bas_annuel": to_num(reste[3]),
                    })
                    continue

                # ---- Indices : Nom | Valeur | Var veille % | Var 31/12 %
                if first.startswith("MASI") and len(cells) == 4 and cells[2].endswith("%"):
                    indices.append({
                        "Indice": first.replace("®", "").strip(),
                        "Valeur": to_num(cells[1]),
                        "Var_jour_pct": to_num(cells[2]),
                        "Var_annee_pct": to_num(cells[3]),
                    })

    df_cot = pd.DataFrame(cotation).drop_duplicates(subset=["ISIN"])
    df_phy = pd.DataFrame(physio)
    df_idx = pd.DataFrame(indices).drop_duplicates(subset=["Indice"])
    for df in (df_cot, df_phy, df_idx):
        df.insert(0, "Date", pd.Timestamp(seance))

    # Variation du jour par rapport au cours de référence
    if not df_cot.empty:
        df_cot["Var_pct"] = (df_cot["Cloture"] / df_cot["Cours_ref"] - 1) * 100

    # Contrôle : somme physionomie vs total officiel
    officiel = totaux.get("ACTIONS", (None, None, None))
    somme = df_phy["Volume_MAD"].sum() if not df_phy.empty else 0
    controle = {
        "volume_officiel": officiel[1],
        "volume_lu": somme,
        "ok": officiel[1] is not None and abs(somme - officiel[1]) < 1,
    }

    return {
        "date": seance,
        "cotation": df_cot,
        "physionomie": df_phy,
        "indices": df_idx,
        "totaux": {"quantite": officiel[0], "volume": officiel[1], "contrats": officiel[2]},
        "controle": controle,
    }


if __name__ == "__main__":
    import sys

    for path in sys.argv[1:]:
        r = parse_bulletin(path)
        c = r["controle"]
        print(f"{r['date']} | {len(r['cotation'])} valeurs | {len(r['physionomie'])} lignes physionomie | "
              f"volume lu {c['volume_lu']:,.2f} / officiel {c['volume_officiel']:,.2f} | "
              f"{'OK' if c['ok'] else 'ECART'}")
