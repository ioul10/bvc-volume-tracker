"""
Indice de liquidité du MASI 20 (base 100 = période de référence).

Étape 1 - trois mesures par action et par semaine (moyennes par séance, pour comparer des semaines incomplètes) :
    Argent        = volume par séance, hors blocs probables (MAD)
    Intervenants  = nombre de contrats par séance
    Résistance    = 1 / moyenne des ratios journaliers d'Amihud (|variation %| / volume en M MAD)
                    = M MAD qu'il faut échanger pour faire bouger le cours de 1 %
Étape 2 - chaque mesure est divisée par sa moyenne sur la période de référence, rapport borné entre 0,25 et 4.
Étape 3 - note de l'action : L = racine cubique (rapport argent x rapport intervenants x rapport résistance).
Étape 4 - indice = 100 x somme(poids plafonné à 15 % x L).
Étape 5 - météo : >= 110 plus liquide, 90 à 110 habituel, < 90 moins liquide.

Référence : les premières semaines de l'historique (12 semaines = définitive ; de 4 à 11 = provisoire).
"""
import numpy as np
import pandas as pd

from bvc_liquidite import volume_fond_jour

SEMAINES_REF = 12          # référence définitive
SEMAINES_REF_MIN = 4       # référence provisoire
BORNE_BAS, BORNE_HAUT = 0.25, 4.0
PLAFOND_POIDS = 0.15
SEUIL_HAUT, SEUIL_BAS = 110, 90


def _semaine(d):
    iso = pd.Timestamp(d).isocalendar()
    return (int(iso[0]), int(iso[1]))


def mesures_hebdo(h, vj, noms):
    """Une ligne par action et par semaine : Argent, Intervenants, Résistance (par séance)."""
    d = vj[vj["Valeur"].isin(noms)][["Date", "Valeur", "Volume_MAD", "Contrats", "Var_pct"]].copy()
    f = volume_fond_jour(h)
    d = d.merge(f, on=["Date", "Valeur"], how="left")
    d["Volume_fond"] = d["Volume_fond"].fillna(0)
    d["Semaine"] = d["Date"].map(_semaine)
    # Amihud du jour, seulement les jours échangés
    d["Amihud"] = np.where(d["Volume_MAD"] > 0, d["Var_pct"].abs() / (d["Volume_MAD"] / 1e6), np.nan)
    seances = d.groupby("Semaine")["Date"].nunique().rename("Seances")
    g = d.groupby(["Semaine", "Valeur"]).agg(
        Argent_total=("Volume_fond", "sum"), Contrats_total=("Contrats", "sum"),
        Amihud_moyen=("Amihud", "mean"), Debut=("Date", "min"), Fin=("Date", "max")).reset_index()
    g = g.merge(seances, on="Semaine")
    g["Argent"] = g["Argent_total"] / g["Seances"]
    g["Intervenants"] = g["Contrats_total"] / g["Seances"]
    g["Resistance"] = np.where(g["Amihud_moyen"] > 0, 1e6 / g["Amihud_moyen"], np.nan)  # MAD pour 1 %
    return g


def poids_plafonnes(w, plafond=PLAFOND_POIDS):
    w = w.fillna(0).clip(lower=0)
    if w.sum() <= 0:
        return w
    w = w / w.sum()
    for _ in range(50):
        over = w > plafond + 1e-12
        if not over.any():
            break
        exces = (w[over] - plafond).sum()
        w[over] = plafond
        libres = ~over & (w < plafond)
        if w[libres].sum() <= 0:
            break
        w[libres] += exces * w[libres] / w[libres].sum()
    return w


def poids_semaine(h, noms, debut):
    """Poids approchés (nombre de titres x cours de référence) au 1er jour de la semaine, plafonnés."""
    cot = h["Cotation"].copy()
    if "Nb_titres" not in cot.columns:
        cot["Nb_titres"] = np.nan
    connu = cot.dropna(subset=["Nb_titres"]).sort_values("Date").groupby("Valeur")["Nb_titres"].last()
    c = cot[(cot["Date"] == debut) & cot["Valeur"].isin(noms)].copy()
    c["Nb_titres"] = c["Nb_titres"].fillna(c["Valeur"].map(connu))
    cap = (c["Nb_titres"] * c["Cours_ref"]).fillna(0)
    cap.index = c["Valeur"]
    return poids_plafonnes(cap.reindex(noms))


def poids_officiels(compo):
    off = compo.dropna(subset=["Poids_pct"]).set_index("Valeur")["Poids_pct"]
    if len(off) == len(compo) and len(off):
        return poids_plafonnes(off)
    return None


def calculer(h, vj, compo):
    """Retourne un dict : etat, semaines (série de l'indice), detail (par action et semaine), reference."""
    noms = [v for v in compo["Valeur"] if v in set(vj["Valeur"])]
    m = mesures_hebdo(h, vj, noms)
    sem = sorted(m["Semaine"].unique())
    n_sem = len(sem)
    n_seances = int(vj["Date"].nunique())
    etat = {"semaines": n_sem, "seances": n_seances,
            "statut": "definitive" if n_sem >= SEMAINES_REF else "provisoire" if n_sem >= SEMAINES_REF_MIN else "insuffisant"}
    if etat["statut"] == "insuffisant":
        return {"etat": etat, "mesures": m, "semaines": None, "detail": None, "reference": None, "noms": noms}

    ref_sem = sem[:min(SEMAINES_REF, n_sem)]
    etat["ref_debut"] = m[m["Semaine"] == ref_sem[0]]["Debut"].min()
    etat["ref_fin"] = m[m["Semaine"] == ref_sem[-1]]["Fin"].max()
    ref = m[m["Semaine"].isin(ref_sem)].groupby("Valeur").agg(
        Ref_argent=("Argent", "mean"), Ref_intervenants=("Intervenants", "mean"),
        Ref_resistance=("Resistance", "mean")).reset_index()

    d = m.merge(ref, on="Valeur", how="left")

    def rapport(a, b):
        r = np.where((b > 0) & np.isfinite(b), a / b, np.nan)
        return np.clip(r, BORNE_BAS, BORNE_HAUT)

    d["R_argent"] = rapport(d["Argent"], d["Ref_argent"])
    d["R_intervenants"] = rapport(d["Intervenants"], d["Ref_intervenants"])
    d["R_resistance"] = rapport(d["Resistance"], d["Ref_resistance"])
    # Une semaine sans aucun échange : liquidité au plancher. Une résistance non mesurable (prix immobile) : neutre.
    vide = d["Argent"] <= 0
    d.loc[vide, ["R_argent", "R_intervenants", "R_resistance"]] = BORNE_BAS
    d["R_resistance"] = d["R_resistance"].fillna(1.0)
    d[["R_argent", "R_intervenants"]] = d[["R_argent", "R_intervenants"]].fillna(BORNE_BAS)
    d["L"] = np.cbrt(d["R_argent"] * d["R_intervenants"] * d["R_resistance"])

    w_off = poids_officiels(compo)
    lignes = []
    for s in sem:
        x = d[d["Semaine"] == s].set_index("Valeur").reindex(noms)
        w = w_off if w_off is not None else poids_semaine(h, noms, x["Debut"].min())
        x["Poids"] = w.reindex(noms).fillna(0).values
        x["Contribution"] = 100 * x["Poids"] * x["L"].fillna(BORNE_BAS)
        d.loc[d["Semaine"] == s, "Poids"] = d.loc[d["Semaine"] == s, "Valeur"].map(x["Poids"])
        d.loc[d["Semaine"] == s, "Contribution"] = d.loc[d["Semaine"] == s, "Valeur"].map(x["Contribution"])
        lignes.append({"Semaine": s, "Debut": x["Debut"].min(), "Fin": x["Fin"].max(),
                       "Indice": x["Contribution"].sum(),
                       "Profondeur_MAD": x["Resistance"].sum(skipna=True),
                       "Argent": x["Argent"].sum(), "Intervenants": x["Intervenants"].sum(),
                       "Dans_reference": s in ref_sem})
    serie = pd.DataFrame(lignes)
    serie["Meteo"] = serie["Indice"].map(meteo)
    return {"etat": etat, "mesures": m, "semaines": serie, "detail": d, "reference": ref, "noms": noms,
            "source_poids": "officiels" if w_off is not None else "approchés"}


def meteo(v):
    if v is None or np.isnan(v):
        return ("⚪", "Indisponible", "neutre")
    if v >= SEUIL_HAUT:
        return ("🟢", "Plus liquide que la référence", "ok")
    if v < SEUIL_BAS:
        return ("🔴", "Moins liquide que la référence", "bad")
    return ("🟡", "Liquidité habituelle", "warn")
