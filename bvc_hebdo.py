"""
Revue hebdomadaire MASI 20 : prix, contributions et volumes, exportée en Word (A4 portrait).

Poids : si masi20.csv contient une colonne Poids_pct remplie (poids officiels), ils sont utilisés.
Sinon, poids approchés = nombre de titres x cours de référence du 1er jour de la semaine.
La page affiche l'écart entre la variation reconstituée et la variation publiée du MASI 20,
ce qui permet de vérifier que les poids sont cohérents.
"""
import io

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
from docx import Document  # noqa: E402
from docx.enum.section import WD_ORIENT  # noqa: E402
from docx.enum.table import WD_TABLE_ALIGNMENT  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm, Pt, RGBColor  # noqa: E402

from bvc_excel import _signal  # noqa: E402

NAVY, GOLD, GREEN, RED, GREY = "0B1F3A", "C8A24A", "12A150", "D9483B", "6B7A90"
MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
        "septembre", "octobre", "novembre", "décembre"]


# ====================================================================== composition
def charger_composition(path="masi20.csv"):
    try:
        c = pd.read_csv(path)
    except FileNotFoundError:
        return pd.DataFrame(columns=["Valeur", "Poids_pct"])
    if "Poids_pct" not in c.columns:
        c["Poids_pct"] = np.nan
    c["Valeur"] = c["Valeur"].astype(str).str.strip()
    return c


def semaines(dates):
    """{(année, semaine ISO): [dates]} dans l'ordre chronologique."""
    out = {}
    for d in sorted(pd.to_datetime(pd.Series(dates)).unique()):
        k = tuple(pd.Timestamp(d).isocalendar()[:2])
        out.setdefault(k, []).append(pd.Timestamp(d))
    return out


def libelle_semaine(jours):
    a, b = jours[0], jours[-1]
    return f"Semaine du {a.day} au {b.day} {MOIS[b.month - 1]} {b.year}" if a.month == b.month else \
        f"Semaine du {a.day} {MOIS[a.month - 1]} au {b.day} {MOIS[b.month - 1]} {b.year}"


# ====================================================================== calculs
def _fr(v, d=1):
    return f"{v:,.{d}f}".replace(",", " ").replace(".", ",")


def _pct(v, d=2):
    return ("+" if v > 0 else "") + _fr(v, d) + " %"


def donnees_hebdo(h, vj, liq, compo, jours, jours_prec=None):
    jours = [pd.Timestamp(d) for d in jours]
    noms = list(compo["Valeur"])
    cot = h["Cotation"].copy()
    # Nombre de titres : on complète les trous (anciens historiques) avec la dernière valeur connue
    if "Nb_titres" in cot.columns and cot["Nb_titres"].notna().any():
        connu = cot.dropna(subset=["Nb_titres"]).sort_values("Date").groupby("Valeur")["Nb_titres"].last()
        cot["Nb_titres"] = cot["Nb_titres"].fillna(cot["Valeur"].map(connu))
    else:
        cot["Nb_titres"] = np.nan
    d = vj[vj["Date"].isin(jours) & vj["Valeur"].isin(noms)].copy()
    manquants = sorted(set(noms) - set(d["Valeur"]))

    # ---- poids au début de la semaine
    c0 = cot[(cot["Date"] == jours[0]) & cot["Valeur"].isin(noms)].set_index("Valeur")
    officiels = compo.set_index("Valeur")["Poids_pct"].dropna()
    if len(officiels) == len(noms):
        poids = officiels / officiels.sum()
        source_poids = "poids officiels (masi20.csv)"
    else:
        cap = c0["Nb_titres"] * c0["Cours_ref"]
        poids = cap / cap.sum()
        source_poids = "poids approchés (nombre de titres × cours, sans flottant ni plafonnement)"

    # ---- poids de chaque séance (recalculés chaque jour avec le cours de référence)
    poids_jour = {}
    for dt in jours:
        cj = cot[(cot["Date"] == dt) & cot["Valeur"].isin(noms)].set_index("Valeur")
        if len(officiels) == len(noms):
            poids_jour[dt] = poids
        else:
            capj = cj["Nb_titres"] * cj["Cours_ref"]
            poids_jour[dt] = capj / capj.sum()

    # ---- MASI 20 publié
    idx = h["Indices"]
    m20 = idx[(idx["Indice"] == "MASI 20") & idx["Date"].isin(jours)].sort_values("Date")
    masi = idx[(idx["Indice"] == "MASI") & idx["Date"].isin(jours)].sort_values("Date")
    m20_debut = m20["Valeur"].iloc[0] / (1 + m20["Var_jour_pct"].iloc[0] / 100)
    m20_fin = m20["Valeur"].iloc[-1]
    m20_var = (m20_fin / m20_debut - 1) * 100
    masi_var = (masi["Valeur"].iloc[-1] / (masi["Valeur"].iloc[0] / (1 + masi["Var_jour_pct"].iloc[0] / 100)) - 1) * 100

    # ---- par valeur
    rows = []
    vol_prec = {}
    if jours_prec:
        vp = vj[vj["Date"].isin(jours_prec) & vj["Valeur"].isin(noms)]
        vol_prec = vp.groupby("Valeur")["Volume_MAD"].sum().to_dict()
    blocs_semaine = []
    for v in noms:
        g = d[d["Valeur"] == v].sort_values("Date")
        if g.empty:
            continue
        ref0 = g["Cours_ref"].iloc[0]
        close = g["Cloture"].iloc[-1]
        r = (close / ref0 - 1) * 100 if ref0 and close else 0.0
        # contribution chaînée : somme des (poids du jour x variation du jour)
        contrib = float(sum(poids_jour[dt].get(v, 0) * vr for dt, vr in zip(g["Date"], g["Var_pct"].fillna(0))))
        vol = g["Volume_MAD"].sum()
        vv = g["Volume_MAD"].values
        k = len(vv) - len(vv) // 2
        if v in vol_prec and vol_prec[v] > 0:
            tendance = (vol / len(jours)) / (vol_prec[v] / len(jours_prec))
        else:
            tendance = vv[k:].mean() / vv[:k].mean() if len(vv) > 1 and vv[:k].mean() > 0 else np.nan
        w = float(poids.get(v, 0))
        lq = liq[liq["Valeur"] == v]
        rows.append({
            "Valeur": v, "Secteur": g["Secteur_nom"].iloc[-1], "Poids_pct": w * 100,
            "Cloture": close, "Var_semaine_pct": r, "Contribution_pct": contrib,
            "Contribution_pts": contrib / 100 * m20_debut,
            "Volume_MAD": vol, "Tendance_volume": tendance,
            "Ticket_moyen": vol / g["Contrats"].sum() if g["Contrats"].sum() else np.nan,
            "Signal": _signal(r, tendance, vol),
            "Score_liquidite": lq["Score"].iloc[0] if len(lq) else np.nan,
            "Plus_haut_annuel": g["Plus_haut_annuel"].iloc[-1],
        })
    t = pd.DataFrame(rows)
    vol_m20 = t["Volume_MAD"].sum()
    t["Part_volume_pct"] = t["Volume_MAD"] / vol_m20 * 100
    t["Volume_vs_poids"] = t["Part_volume_pct"] / t["Poids_pct"]
    t["Confirme"] = t["Tendance_volume"].fillna(1) >= 1.2
    t = t.sort_values("Contribution_pct", ascending=False).reset_index(drop=True)

    # ---- séance par séance
    marche = h["Marche"].set_index("Date")
    jour = []
    for dt in jours:
        x = d[d["Date"] == dt]
        jour.append({
            "Date": dt,
            "MASI20": m20.set_index("Date")["Valeur"].get(dt, np.nan),
            "MASI20_var": m20.set_index("Date")["Var_jour_pct"].get(dt, np.nan),
            "Volume_MASI20": x["Volume_MAD"].sum(),
            "Volume_marche": marche["Volume_MAD"].get(dt, np.nan),
            "Hausses": int((x["Var_pct"] > 0).sum()), "Baisses": int((x["Var_pct"] < 0).sum()),
            "Reconstitue": float((x.set_index("Valeur")["Var_pct"] * poids_jour[dt]).sum()),
        })
    jour = pd.DataFrame(jour)
    jour["Ecart"] = jour["Reconstitue"] - jour["MASI20_var"]

    # ---- grosses transactions sur les valeurs de l'indice
    phy = h["Physionomie"]
    b = phy[phy["Date"].isin(jours) & phy["Valeur"].isin(noms)].copy()
    b["par_contrat"] = b["Volume_MAD"] / b["Contrats"]
    b = b[(b["Volume_MAD"] >= 5e6) & (b["par_contrat"] >= 1e6)].sort_values("Volume_MAD", ascending=False)
    blocs_semaine = b
    vol_blocs = b["Volume_MAD"].sum()

    # ---- secteurs
    sect = t.groupby("Secteur").agg(Poids_pct=("Poids_pct", "sum"), Contribution_pct=("Contribution_pct", "sum"),
                                    Volume_MAD=("Volume_MAD", "sum")).reset_index()
    sect["Part_volume_pct"] = sect["Volume_MAD"] / vol_m20 * 100
    sect = sect.sort_values("Contribution_pct")

    # ---- qualité du mouvement
    pos, neg = t[t["Contribution_pct"] > 0], t[t["Contribution_pct"] < 0]
    q_pos = pos.loc[pos["Confirme"], "Contribution_pct"].sum() / pos["Contribution_pct"].sum() * 100 if len(pos) else np.nan
    q_neg = neg.loc[neg["Confirme"], "Contribution_pct"].sum() / neg["Contribution_pct"].sum() * 100 if len(neg) else np.nan

    vol_marche = jour["Volume_marche"].sum()
    resume = {
        "libelle": libelle_semaine(jours), "jours": jours, "m20_debut": m20_debut, "m20_fin": m20_fin,
        "m20_var": m20_var, "masi_var": masi_var, "vol_m20": vol_m20, "vol_marche": vol_marche,
        "part_m20": vol_m20 / vol_marche * 100 if vol_marche else np.nan,
        "vol_m20_prec": sum(vol_prec.values()) if vol_prec else None,
        "vol_blocs": vol_blocs, "hausses": int((t["Var_semaine_pct"] > 0).sum()),
        "baisses": int((t["Var_semaine_pct"] < 0).sum()), "reconstitue": t["Contribution_pct"].sum(),
        "q_pos": q_pos, "q_neg": q_neg, "source_poids": source_poids, "manquants": manquants,
    }
    resume["messages"] = _messages(resume, t, sect)
    resume["surveiller"] = _surveiller(t, blocs_semaine)
    return {"resume": resume, "valeurs": t, "jours": jour, "secteurs": sect, "blocs": blocs_semaine}


def _noms(df, n=3):
    return ", ".join(f"{r.Valeur} ({_pct(r.Var_semaine_pct, 1)})" for r in df.head(n).itertuples())


def _messages(r, t, sect):
    sens = "gagné" if r["m20_var"] >= 0 else "perdu"
    msgs = [f"Le MASI 20 a {sens} {_fr(abs(r['m20_var']), 2)} % sur la semaine "
            f"({_fr(r['m20_debut'], 2)} → {_fr(r['m20_fin'], 2)}), contre {_pct(r['masi_var'])} pour le MASI. "
            f"{r['hausses']} valeurs en hausse, {r['baisses']} en baisse."]
    up, dn = t[t["Contribution_pct"] > 0], t[t["Contribution_pct"] < 0].sort_values("Contribution_pct")
    if len(up):
        msgs.append(f"Ont tiré l'indice vers le haut : {_noms(up)}. "
                    f"Vers le bas : {_noms(dn) if len(dn) else 'aucune valeur'}.")
    v = f"Les 20 valeurs ont traité {_fr(r['vol_m20'] / 1e6, 1)} M MAD, soit {_fr(r['part_m20'], 0)} % du marché"
    if r["vol_m20_prec"]:
        v += f" ({_pct((r['vol_m20'] / r['vol_m20_prec'] - 1) * 100, 0)} par rapport à la semaine précédente)"
    v += f". Les grosses transactions en représentent {_fr(r['vol_blocs'] / r['vol_m20'] * 100, 0)} %."
    msgs.append(v)
    q = []
    if not np.isnan(r["q_pos"]):
        q.append(f"{_fr(r['q_pos'], 0)} % de la contribution positive vient de valeurs dont le volume augmente")
    if not np.isnan(r["q_neg"]):
        q.append(f"{_fr(r['q_neg'], 0)} % de la contribution négative aussi")
    if q:
        msgs.append("Qualité du mouvement : " + " ; ".join(q) + ".")
    if len(sect):
        best, worst = sect.iloc[-1], sect.iloc[0]
        msgs.append(f"Secteurs : {best.Secteur} contribue le plus ({_pct(best.Contribution_pct)}), "
                    f"{worst.Secteur} le moins ({_pct(worst.Contribution_pct)}).")
    return msgs


def _surveiller(t, blocs):
    out = []
    for r in t.itertuples():
        if r.Signal == "Hausse sans volume" and r.Contribution_pct > 0:
            out.append(f"{r.Valeur} : hausse de {_pct(r.Var_semaine_pct, 1)} sans volume, elle reste fragile.")
        if r.Signal == "Baisse sous volume":
            out.append(f"{r.Valeur} : baisse de {_pct(r.Var_semaine_pct, 1)} avec un volume en hausse, la pression vendeuse est réelle.")
        if r.Volume_vs_poids >= 2:
            out.append(f"{r.Valeur} : {_fr(r.Part_volume_pct, 0)} % du volume de l'indice pour {_fr(r.Poids_pct, 1)} % du poids, "
                       "l'argent s'y concentre.")
        if r.Plus_haut_annuel and r.Cloture >= 0.98 * r.Plus_haut_annuel:
            out.append(f"{r.Valeur} : clôture à moins de 2 % de son plus haut annuel.")
    for r in blocs.head(3).itertuples():
        out.append(f"{r.Valeur} : bloc de {_fr(r.Volume_MAD / 1e6, 1)} M MAD à {_fr(r.Prix, 2)} le {r.Date:%d/%m}.")
    seen, res = set(), []
    for x in out:
        if x not in seen:
            seen.add(x)
            res.append(x)
    return res[:8]


# ====================================================================== graphiques
def _style_ax(ax):
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.spines["left"].set_color("#C9D2DE")
    ax.spines["bottom"].set_color("#C9D2DE")
    ax.tick_params(colors="#4A5A70", labelsize=7)
    ax.grid(axis="y", color="#E9EDF3", lw=.6)
    ax.set_axisbelow(True)


def _png(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def graphe_indice_volume(jours):
    fig, ax = plt.subplots(figsize=(7.2, 2.4))
    x = np.arange(len(jours))
    lab = [d.strftime("%a %d/%m") for d in jours["Date"]]
    ax.bar(x, jours["Volume_marche"] / 1e6, color="#DCE3EC", width=.6, label="Marché total")
    ax.bar(x, jours["Volume_MASI20"] / 1e6, color="#" + NAVY, width=.6, label="Valeurs du MASI 20")
    ax.set_ylabel("Volume (M MAD)", fontsize=7, color="#4A5A70")
    ax.set_xticks(x, lab)
    _style_ax(ax)
    ax2 = ax.twinx()
    ax2.plot(x, jours["MASI20"], color="#" + GOLD, lw=2.2, marker="o", ms=4, label="MASI 20")
    for i, (v, c) in enumerate(zip(jours["MASI20"], jours["MASI20_var"])):
        ax2.annotate(_pct(c), (i, v), textcoords="offset points", xytext=(0, 8), ha="center", fontsize=6.5,
                     color="#" + (GREEN if c >= 0 else RED),
                     bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="none", alpha=.9))
    ax2.spines["top"].set_visible(False)
    ax2.tick_params(colors="#" + GOLD, labelsize=7)
    lo, hi = jours["MASI20"].min(), jours["MASI20"].max()
    ax2.set_ylim(lo - (hi - lo) * .6 - 1, hi + (hi - lo) * .6 + 1)
    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, fontsize=6.5, frameon=False, loc="upper left", ncol=3)
    return _png(fig)


def graphe_contributions(t):
    s = t.sort_values("Contribution_pct")
    fig, ax = plt.subplots(figsize=(3.6, 4.6))
    cols = ["#" + (GREEN if c > 0 else RED) for c in s["Contribution_pct"]]
    alphas = [1 if conf else .45 for conf in s["Confirme"]]
    bars = ax.barh(s["Valeur"], s["Contribution_pct"] * 100, color=cols)
    for b, a in zip(bars, alphas):
        b.set_alpha(a)
    ax.axvline(0, color="#4A5A70", lw=.8)
    ax.set_xlabel("Contribution (points de base)", fontsize=7, color="#4A5A70")
    _style_ax(ax)
    ax.grid(axis="x", color="#E9EDF3", lw=.6)
    ax.grid(axis="y", visible=False)
    ax.tick_params(axis="y", labelsize=6.5)
    ax.set_title("Couleur pleine = volume en hausse", fontsize=6.5, color="#4A5A70", loc="left")
    return _png(fig)


def graphe_volume_poids(t):
    fig, ax = plt.subplots(figsize=(3.6, 4.6))
    m = max(t["Poids_pct"].max(), t["Part_volume_pct"].max()) * 1.1
    ax.plot([0, m], [0, m], color="#C9D2DE", lw=1, ls="--")
    cols = ["#" + (GREEN if c > 0 else RED) for c in t["Var_semaine_pct"]]
    ax.scatter(t["Poids_pct"], t["Part_volume_pct"], s=28, c=cols, edgecolors="white", linewidths=.6, zorder=3)
    for r in t.itertuples():
        if r.Volume_vs_poids >= 1.6 or r.Poids_pct >= 8 or r.Volume_vs_poids <= .35:
            ax.annotate(r.Valeur.title()[:16], (r.Poids_pct, r.Part_volume_pct), fontsize=5.5,
                        xytext=(3, 3), textcoords="offset points", color="#13294B")
    ax.set_xlim(0, m)
    ax.set_ylim(0, m)
    ax.set_xlabel("Poids dans l'indice (%)", fontsize=7, color="#4A5A70")
    ax.set_ylabel("Part du volume de l'indice (%)", fontsize=7, color="#4A5A70")
    ax.text(m * .03, m * .93, "Plus de volume\nque de poids", fontsize=6, color="#" + GREY, va="top")
    ax.text(m * .97, m * .05, "Moins de volume\nque de poids", fontsize=6, color="#" + GREY, ha="right")
    _style_ax(ax)
    return _png(fig)


# ====================================================================== Word
def _shade(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _cell_text(cell, text, size=8, bold=False, color=None, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    if align:
        p.alignment = align
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def _largeurs(table, cms):
    """Fixe la largeur des colonnes (Word lit la grille ET chaque cellule)."""
    grid = table._tbl.tblGrid
    for j, gc in enumerate(grid.findall(qn("w:gridCol"))):
        if j < len(cms):
            gc.set(qn("w:w"), str(int(Cm(cms[j]).twips)))
    for row in table.rows:
        for j, w in enumerate(cms):
            row.cells[j].width = Cm(w)


def _titre(doc, txt):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(txt.upper())
    r.bold = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor.from_string(NAVY)
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    for k, v in (("w:val", "single"), ("w:sz", "8"), ("w:space", "1"), ("w:color", GOLD)):
        b.set(qn(k), v)
    bdr.append(b)
    pPr.append(bdr)


def rapport_word(data):
    r, t, jours = data["resume"], data["valeurs"], data["jours"]
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.6)
    sec.top_margin, sec.bottom_margin = Cm(1.3), Cm(1.3)
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.font.size = Pt(9)

    # ---- bandeau
    tb = doc.add_table(rows=1, cols=2)
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    left, right = tb.rows[0].cells
    for c in (left, right):
        _shade(c, NAVY)
    left.width, right.width = Cm(12.3), Cm(5.5)
    _cell_text(left, "MASI 20 · REVUE HEBDOMADAIRE", 15, True, "FFFFFF")
    p = left.add_paragraph()
    rr = p.add_run(r["libelle"] + " · prix, contributions et volumes")
    rr.font.size, rr.font.color.rgb = Pt(9), RGBColor.from_string("B9C4D6")
    _cell_text(right, _fr(r["m20_fin"], 2), 17, True, GOLD, WD_ALIGN_PARAGRAPH.RIGHT)
    p = right.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    rr = p.add_run(f"{_pct(r['m20_var'])} sur la semaine")
    rr.bold = True
    rr.font.size = Pt(9)
    rr.font.color.rgb = RGBColor.from_string("3DDC84" if r["m20_var"] >= 0 else "FF6B5E")

    # ---- chiffres clés
    _titre(doc, "Chiffres clés")
    k = doc.add_table(rows=2, cols=5)
    k.alignment = WD_TABLE_ALIGNMENT.CENTER
    vals = [
        ("MASI 20", _pct(r["m20_var"]), "MASI " + _pct(r["masi_var"])),
        ("Hausses / baisses", f"{r['hausses']} / {r['baisses']}", "sur 20 valeurs"),
        ("Volume MASI 20", _fr(r["vol_m20"] / 1e6, 0) + " M MAD", f"{_fr(r['part_m20'], 0)} % du marché"),
        ("Grosses transactions", _fr(r["vol_blocs"] / r["vol_m20"] * 100, 0) + " %", "du volume de l'indice"),
        ("Contrôle des poids", _pct(r["reconstitue"]), f"reconstitué vs {_pct(r['m20_var'])} publié"),
    ]
    for j, (lab, val, sub) in enumerate(vals):
        c0, c1 = k.rows[0].cells[j], k.rows[1].cells[j]
        _shade(c0, "F4F6FA")
        _shade(c1, "F4F6FA")
        _cell_text(c0, lab.upper(), 6.5, True, GREY)
        _cell_text(c1, val, 12, True, NAVY)
        p = c1.add_paragraph()
        x = p.add_run(sub)
        x.font.size, x.font.color.rgb = Pt(7), RGBColor.from_string(GREY)

    # ---- messages
    _titre(doc, "L'essentiel de la semaine")
    for m in r["messages"]:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(1)
        p.add_run(m).font.size = Pt(8.5)

    # ---- graphe indice / volume
    _titre(doc, "Indice et volume, séance par séance")
    doc.add_picture(graphe_indice_volume(jours), width=Cm(17.6))

    # ---- contributions + volume vs poids
    _titre(doc, "Qui a fait bouger l'indice, et où va l'argent")
    g = doc.add_table(rows=1, cols=2)
    a, b = g.rows[0].cells
    a.paragraphs[0].add_run().add_picture(graphe_contributions(t), width=Cm(8.6))
    b.paragraphs[0].add_run().add_picture(graphe_volume_poids(t), width=Cm(8.6))

    # ---- tableau
    doc.add_page_break()
    _titre(doc, "Détail des 20 valeurs (triées par contribution)")
    cols = ["Valeur", "Poids", "Var. sem.", "Contrib. (pb)", "Volume (M MAD)", "Vol./poids", "Tend. vol.", "Liquidité", "Signal"]
    tab = doc.add_table(rows=1, cols=len(cols))
    tab.alignment = WD_TABLE_ALIGNMENT.CENTER
    tab.style = "Table Grid"
    tab.autofit = False
    for j, c in enumerate(cols):
        _shade(tab.rows[0].cells[j], NAVY)
        _cell_text(tab.rows[0].cells[j], c, 7, True, "FFFFFF")
    for i, x in enumerate(t.itertuples()):
        row = tab.add_row().cells
        vals = [x.Valeur, _fr(x.Poids_pct, 1) + " %", _pct(x.Var_semaine_pct, 1), _fr(x.Contribution_pct * 100, 1),
                _fr(x.Volume_MAD / 1e6, 1), _fr(x.Volume_vs_poids, 1) + "×",
                "—" if np.isnan(x.Tendance_volume) else _fr(x.Tendance_volume, 2) + "×",
                "—" if np.isnan(x.Score_liquidite) else _fr(x.Score_liquidite, 0), x.Signal]
        for j, v in enumerate(vals):
            color = None
            if j in (2, 3):
                color = GREEN if x.Var_semaine_pct > 0 else RED if x.Var_semaine_pct < 0 else None
            _cell_text(row[j], v, 7, j == 0, color, WD_ALIGN_PARAGRAPH.LEFT if j in (0, 8) else WD_ALIGN_PARAGRAPH.RIGHT)
            if i % 2:
                _shade(row[j], "F4F6FA")
    _largeurs(tab, [3.9, 1.3, 1.5, 1.6, 1.8, 1.4, 1.4, 1.4, 3.5])

    # ---- secteurs
    _titre(doc, "Par secteur")
    s = data["secteurs"].sort_values("Contribution_pct", ascending=False)
    st_ = doc.add_table(rows=1, cols=4)
    st_.style = "Table Grid"
    st_.autofit = False
    for j, c in enumerate(["Secteur", "Poids", "Contribution (pb)", "Part du volume"]):
        _shade(st_.rows[0].cells[j], NAVY)
        _cell_text(st_.rows[0].cells[j], c, 7, True, "FFFFFF")
    for x in s.itertuples():
        row = st_.add_row().cells
        for j, v in enumerate([x.Secteur, _fr(x.Poids_pct, 1) + " %", _fr(x.Contribution_pct * 100, 1),
                               _fr(x.Part_volume_pct, 1) + " %"]):
            _cell_text(row[j], v, 7, j == 0,
                       (GREEN if x.Contribution_pct > 0 else RED) if j == 2 else None,
                       WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.RIGHT)

    _largeurs(st_, [6.0, 3.6, 4.2, 4.0])

    # ---- à surveiller
    _titre(doc, "À surveiller la semaine prochaine")
    for m in r["surveiller"] or ["Rien de particulier."]:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(1)
        p.add_run(m).font.size = Pt(8.5)

    # ---- méthode
    _titre(doc, "Méthode")
    notes = [
        f"Poids : {r['source_poids']}, recalculés chaque séance. Contribution = somme sur la semaine de (poids du jour × variation du jour), "
        "en points de base (1 pb = 0,01 %). Le contrôle compare la somme des contributions à la variation publiée : "
        "un écart faible confirme que les poids sont cohérents.",
        "Vol./poids = part du volume de l'indice ÷ poids. Au-dessus de 1, la valeur attire plus d'argent que son poids. "
        "Tend. vol. = volume moyen de la semaine ÷ semaine précédente (à défaut, 2ᵉ moitié ÷ 1ʳᵉ moitié de la semaine).",
        "Grosses transactions : niveaux de prix d'au moins 5 M MAD et 1 M MAD par contrat. "
        "Liquidité : score de 0 à 100 (volume hors blocs, fréquence, écart achat/vente, impact prix).",
        "Source : bulletins de la cote. Outil indépendant, non affilié à la Bourse de Casablanca. "
        "Ce document ne constitue pas un conseil en investissement.",
    ]
    if r["manquants"]:
        notes.insert(0, "Valeurs de la composition absentes des bulletins : " + ", ".join(r["manquants"]) + ".")
    for n in notes:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        x = p.add_run(n)
        x.font.size, x.font.color.rgb = Pt(7), RGBColor.from_string(GREY)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
