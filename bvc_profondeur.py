"""
Profondeur du MASI 20 : combien d'argent faut-il échanger pour faire bouger l'indice de 1 % ?

Étape 1 - pour chaque action, sur la semaine :
    Profondeur = argent échangé (MAD) / somme des mouvements du cours (|variation %| de chaque séance)
Étape 2 - pour le MASI 20 :
    Profondeur MASI 20 = somme des profondeurs des 20 actions
Étape 3 - indice de liquidité :
    Indice = 100 x profondeur de la semaine / moyenne des profondeurs des semaines précédentes (12 au plus)
"""
import numpy as np
import pandas as pd

from bvc_liquidite import volume_fond_jour

SEMAINES_NORMALE = 12


def profondeur_actions(h, vj, noms, jours, exclure_blocs=True):
    """Une ligne par action du MASI 20 pour les séances `jours`."""
    jours = [pd.Timestamp(d) for d in jours]
    d = vj[vj["Valeur"].isin(noms) & vj["Date"].isin(jours)][["Date", "Valeur", "Volume_MAD", "Var_pct"]].copy()
    if exclure_blocs:
        f = volume_fond_jour(h, jours)
        d = d.merge(f, on=["Date", "Valeur"], how="left")
        d["Argent"] = d["Volume_fond"].fillna(0)
    else:
        d["Argent"] = d["Volume_MAD"]
    d["Mouvement"] = d["Var_pct"].abs().fillna(0)
    g = d.groupby("Valeur").agg(Argent_MAD=("Argent", "sum"), Volume_total_MAD=("Volume_MAD", "sum"),
                                Mouvement_pct=("Mouvement", "sum"),
                                Seances_echangees=("Volume_MAD", lambda s: int((s > 0).sum()))).reset_index()
    g = pd.DataFrame({"Valeur": noms}).merge(g, on="Valeur", how="left").fillna(
        {"Argent_MAD": 0, "Volume_total_MAD": 0, "Mouvement_pct": 0, "Seances_echangees": 0})
    g["Profondeur_MAD"] = np.where(g["Mouvement_pct"] > 0, g["Argent_MAD"] / g["Mouvement_pct"].replace(0, np.nan), np.nan)
    tot = g["Profondeur_MAD"].sum(skipna=True)
    g["Part_pct"] = g["Profondeur_MAD"] / tot * 100 if tot else np.nan
    g["Remarque"] = np.where(g["Seances_echangees"] == 0, "Aucun échange",
                             np.where(g["Mouvement_pct"] == 0, "Prix immobile : profondeur non mesurable", ""))
    return g.sort_values("Profondeur_MAD", ascending=False, na_position="last").reset_index(drop=True)


def serie_hebdo(h, vj, noms, semaines, exclure_blocs=True):
    """Profondeur du MASI 20 semaine par semaine, et indice de liquidité base 100."""
    rows = []
    for k, jours in semaines.items():
        g = profondeur_actions(h, vj, noms, jours, exclure_blocs)
        rows.append({"Cle": k, "Debut": jours[0], "Fin": jours[-1], "Seances": len(jours),
                     "Profondeur_MAD": g["Profondeur_MAD"].sum(skipna=True),
                     "Actions_mesurees": int(g["Profondeur_MAD"].notna().sum()),
                     "Part_max_pct": g["Part_pct"].max(), "Action_max": g.iloc[0]["Valeur"]})
    s = pd.DataFrame(rows)
    normale, indice, n_ref = [], [], []
    for i in range(len(s)):
        prec = s["Profondeur_MAD"].iloc[max(0, i - SEMAINES_NORMALE):i]
        if len(prec):
            normale.append(prec.mean())
            indice.append(100 * s["Profondeur_MAD"].iloc[i] / prec.mean())
        else:
            normale.append(np.nan)
            indice.append(np.nan)
        n_ref.append(len(prec))
    s["Normale_MAD"], s["Indice"], s["Semaines_reference"] = normale, indice, n_ref
    return s
