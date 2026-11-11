"""
Liquidité du MASI 20 - règle unique : le ratio de liquidité (ratio Amivest).

    Ratio = argent échangé (MAD, hors blocs probables) / somme des |variations quotidiennes en %|
          = MAD qu'il faut échanger pour faire bouger le prix de 1 %

Pour le MASI 20 (chaque semaine) :
    Ratio MASI 20 = argent échangé sur ses 20 valeurs / somme des |variations quotidiennes du MASI 20|
    Indice        = 100 x ratio de la semaine / ratio de la période de référence
                    (même formule appliquée à toute la période de référence)
Pour chaque valeur :
    Ratio de la valeur        = son argent échangé / somme de ses |variations quotidiennes|
    Coût via cette valeur     = ratio de la valeur / son poids dans le MASI 20
                              = MAD pour faire bouger le MASI 20 de 1 % à travers cette seule valeur

Référence : premières semaines de l'historique (12 = définitive, 4 à 11 = provisoire).
"""
import numpy as np
import pandas as pd

from bvc_liquidite import volume_fond_jour

SEMAINES_REF = 12
SEMAINES_REF_MIN = 4
SEUIL_HAUT, SEUIL_BAS = 110, 90


def _semaine(d):
    iso = pd.Timestamp(d).isocalendar()
    return (int(iso[0]), int(iso[1]))


def _ratio(argent, mouvement):
    return np.where(mouvement > 0, argent / np.where(mouvement > 0, mouvement, 1), np.nan)


def poids_approches(h, noms, date):
    """Nombre de titres x cours de référence, à une date (approximation sans flottant ni plafonnement)."""
    cot = h["Cotation"].copy()
    if "Nb_titres" not in cot.columns:
        cot["Nb_titres"] = np.nan
    connu = cot.dropna(subset=["Nb_titres"]).sort_values("Date").groupby("Valeur")["Nb_titres"].last()
    c = cot[(cot["Date"] == pd.Timestamp(date)) & cot["Valeur"].isin(noms)].copy()
    c["Nb_titres"] = c["Nb_titres"].fillna(c["Valeur"].map(connu))
    cap = (c["Nb_titres"] * c["Cours_ref"]).fillna(0)
    cap.index = c["Valeur"]
    cap = cap.reindex(noms).fillna(0)
    return cap / cap.sum() if cap.sum() > 0 else cap


def calculer(h, vj, compo):
    noms = [v for v in compo["Valeur"] if v in set(vj["Valeur"])]
    off = compo.dropna(subset=["Poids_pct"]).set_index("Valeur")["Poids_pct"]
    poids_off = (off / off.sum()) if len(off) == len(compo) and len(off) else None

    # ---- données par valeur et par séance
    d = vj[vj["Valeur"].isin(noms)][["Date", "Valeur", "Var_pct", "Volume_MAD"]].copy()
    d = d.merge(volume_fond_jour(h), on=["Date", "Valeur"], how="left")
    d["Argent"] = d["Volume_fond"].fillna(0)
    d["Mouvement"] = d["Var_pct"].abs().fillna(0)
    d["Semaine"] = d["Date"].map(_semaine)

    # ---- MASI 20 par séance
    idx = h["Indices"]
    m20 = idx[idx["Indice"] == "MASI 20"][["Date", "Valeur", "Var_jour_pct"]].rename(
        columns={"Valeur": "MASI20", "Var_jour_pct": "Var_MASI20"})
    jour = d.groupby("Date")["Argent"].sum().rename("Argent").reset_index().merge(m20, on="Date", how="left")
    jour["Mouvement"] = jour["Var_MASI20"].abs()
    jour["Semaine"] = jour["Date"].map(_semaine)

    # ---- MASI 20 par semaine
    sem = jour.groupby("Semaine").agg(Debut=("Date", "min"), Fin=("Date", "max"), Seances=("Date", "nunique"),
                                      Argent=("Argent", "sum"), Mouvement=("Mouvement", "sum"),
                                      MASI20_fin=("MASI20", "last")).reset_index()
    sem["Ratio"] = _ratio(sem["Argent"], sem["Mouvement"])
    n_sem = len(sem)
    etat = {"semaines": n_sem, "seances": int(jour["Date"].nunique()),
            "statut": "definitive" if n_sem >= SEMAINES_REF else "provisoire" if n_sem >= SEMAINES_REF_MIN else "insuffisant"}

    # ---- par valeur et par semaine
    val = d.groupby(["Semaine", "Valeur"]).agg(Argent=("Argent", "sum"), Mouvement=("Mouvement", "sum"),
                                               Debut=("Date", "min")).reset_index()
    val["Ratio"] = _ratio(val["Argent"], val["Mouvement"])
    poids = []
    for s, g in val.groupby("Semaine"):
        w = poids_off if poids_off is not None else poids_approches(h, noms, g["Debut"].min())
        poids.append(pd.DataFrame({"Semaine": [s] * len(noms), "Valeur": noms, "Poids": w.reindex(noms).fillna(0).values}))
    val = val.merge(pd.concat(poids), on=["Semaine", "Valeur"], how="left")
    val["Cout_via_valeur"] = np.where(val["Poids"] > 0, val["Ratio"] / val["Poids"].replace(0, np.nan), np.nan)

    ref = None
    sem["Indice"] = np.nan
    sem["Dans_reference"] = False
    if etat["statut"] != "insuffisant":
        ref_sem = list(sem["Semaine"])[:min(SEMAINES_REF, n_sem)]
        r = sem[sem["Semaine"].isin(ref_sem)]
        ratio_ref = r["Argent"].sum() / r["Mouvement"].sum() if r["Mouvement"].sum() > 0 else np.nan
        sem["Indice"] = 100 * sem["Ratio"] / ratio_ref
        sem["Dans_reference"] = sem["Semaine"].isin(ref_sem)
        etat.update({"ref_debut": r["Debut"].min(), "ref_fin": r["Fin"].max(), "ratio_ref": ratio_ref})
        vr = val[val["Semaine"].isin(ref_sem)].groupby("Valeur")[["Argent", "Mouvement"]].sum()
        ref = pd.Series(_ratio(vr["Argent"], vr["Mouvement"]), index=vr.index)
        val["Ratio_ref"] = val["Valeur"].map(ref)
        val["Vs_ref_pct"] = (val["Ratio"] / val["Ratio_ref"] - 1) * 100
    else:
        val["Ratio_ref"] = np.nan
        val["Vs_ref_pct"] = np.nan

    # Points faibles : coût via la valeur comparé au ratio du MASI 20 de la même semaine
    val = val.merge(sem[["Semaine", "Ratio"]].rename(columns={"Ratio": "Ratio_MASI20"}), on="Semaine", how="left")
    rel = val["Cout_via_valeur"] / val["Ratio_MASI20"]
    val["Statut"] = np.select([val["Ratio"].isna(), rel < 0.25, rel < 0.5],
                              ["⚪ Non mesurable", "🔴 Point faible", "🟡 À surveiller"], "🟢 Solide")
    sem["Meteo"] = sem["Indice"].map(meteo)
    return {"etat": etat, "semaines": sem, "valeurs": val, "jours": jour, "noms": noms,
            "source_poids": "officiels" if poids_off is not None else "approchés (nombre de titres × cours)"}


def meteo(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ("⚪", "Indisponible", "neutre")
    if v >= SEUIL_HAUT:
        return ("🟢", "Plus liquide que la référence", "ok")
    if v < SEUIL_BAS:
        return ("🔴", "Moins liquide que la référence", "bad")
    return ("🟡", "Liquidité habituelle", "warn")
