"""
Vue « Indices » : MASI, MASI 20, MASI ESG, et leur liquidité.

Liquidité d'un indice
  - Score (0-100)      = moyenne des scores de liquidité des composants, pondérée par leur poids dans l'indice.
  - Coût pour bouger 1 % = somme, sur les composants, des MAD qu'il faut échanger pour faire varier
                         chacun d'eux de 1 % (inverse de la mesure d'impact). C'est l'argent nécessaire
                         pour que tout l'indice bouge de 1 %.
Poids : nombre de titres x cours de clôture (approximation sans flottant ni plafonnement),
        sauf si le fichier de composition donne des poids officiels (colonne Poids_pct).
"""
import numpy as np
import pandas as pd

from bvc_liquidite import volume_fond_jour

INDICES = {
    "MASI": {"fichier": None, "texte": "Toutes les actions cotées"},
    "MASI 20": {"fichier": "masi20.csv", "texte": "Les 20 valeurs les plus liquides"},
    "MASI ESG": {"fichier": "masi_esg.csv", "texte": "Valeurs notées sur l'environnement, le social et la gouvernance"},
}


def composition(h, nom):
    """DataFrame Valeur, Poids_pct (peut être vide). MASI = toutes les actions cotées."""
    fichier = INDICES[nom]["fichier"]
    if fichier is None:
        noms = sorted(h["Cotation"]["Valeur"].unique())
        return pd.DataFrame({"Valeur": noms, "Poids_pct": np.nan})
    try:
        c = pd.read_csv(fichier)
    except (FileNotFoundError, pd.errors.EmptyDataError):
        return pd.DataFrame(columns=["Valeur", "Poids_pct"])
    if "Poids_pct" not in c.columns:
        c["Poids_pct"] = np.nan
    c["Valeur"] = c["Valeur"].astype(str).str.strip()
    return c.dropna(subset=["Valeur"])


def serie(h, nom):
    idx = h["Indices"]
    return idx[idx["Indice"] == nom].sort_values("Date")[["Date", "Valeur", "Var_jour_pct", "Var_annee_pct"]]


def variation_semaine(s, date):
    """Variation depuis la clôture de la semaine précédente jusqu'à `date`."""
    date = pd.Timestamp(date)
    iso = date.isocalendar()[:2]
    sem = s[(s["Date"] <= date) & s["Date"].map(lambda d: pd.Timestamp(d).isocalendar()[:2] == iso)]
    if sem.empty:
        return np.nan, 0
    debut = sem["Valeur"].iloc[0] / (1 + sem["Var_jour_pct"].iloc[0] / 100)
    return (sem["Valeur"].iloc[-1] / debut - 1) * 100, len(sem)


def poids(h, compo, date):
    """Poids des composants (somme = 1) à une date."""
    off = compo.dropna(subset=["Poids_pct"])
    if len(off) and len(off) == len(compo):
        w = off.set_index("Valeur")["Poids_pct"]
        return w / w.sum(), "officiels"
    cot = h["Cotation"].copy()
    if "Nb_titres" not in cot.columns:
        cot["Nb_titres"] = np.nan
    connu = cot.dropna(subset=["Nb_titres"]).sort_values("Date").groupby("Valeur")["Nb_titres"].last()
    c = cot[(cot["Date"] == pd.Timestamp(date)) & cot["Valeur"].isin(compo["Valeur"])].copy()
    c["Nb_titres"] = c["Nb_titres"].fillna(c["Valeur"].map(connu))
    cours = c["Cloture"].fillna(c["Cours_ref"])
    cap = (c["Nb_titres"] * cours).fillna(0)
    cap.index = c["Valeur"]
    if cap.sum() <= 0:
        return pd.Series(dtype=float), "indisponibles"
    return cap / cap.sum(), "approchés"


def liquidite_indice(h, liq, compo, date):
    """Score, coût pour bouger 1 %, et détail par composant."""
    w, source = poids(h, compo, date)
    if w.empty:
        return None
    d = liq.set_index("Valeur").reindex(w.index)
    det = pd.DataFrame({
        "Valeur": w.index,
        "Poids_pct": w.values * 100,
        "Score": d["Score"].values,
        "Classe": d["Classe"].values,
        "MAD_pour_1pct": d["MAD_pour_1pct_impact"].values,
        "Pts_volume": d["Pts_volume"].values, "Pts_frequence": d["Pts_frequence"].values,
        "Pts_ecart": d["Pts_ecart"].values, "Pts_impact": d["Pts_impact"].values,
    })
    det["Score"] = det["Score"].fillna(0)  # composant jamais échangé = liquidité nulle
    det["Apport_score"] = det["Poids_pct"] / 100 * det["Score"]
    score = det["Apport_score"].sum()
    cout = det["MAD_pour_1pct"].sum(skipna=True)
    ingr = {k: float((det["Poids_pct"] / 100 * det[f"Pts_{k}"].fillna(0)).sum())
            for k in ("volume", "frequence", "ecart", "impact")}
    return {
        "score": score, "cout_1pct": cout, "source_poids": source, "n": len(det),
        "n_sans_impact": int(det["MAD_pour_1pct"].isna().sum()),
        "ingredients": ingr,
        "detail": det.sort_values("Apport_score", ascending=False).reset_index(drop=True),
    }


def volume_indice(h, vj, compo, dates):
    """Volume total et hors grosses transactions des composants, par séance."""
    noms = set(compo["Valeur"])
    v = vj[vj["Valeur"].isin(noms) & vj["Date"].isin(dates)].groupby("Date")["Volume_MAD"].sum()
    f = volume_fond_jour(h, dates)
    f = f[f["Valeur"].isin(noms)].groupby("Date")["Volume_fond"].sum()
    out = pd.DataFrame({"Volume": v, "Volume_fond": f}).fillna(0).reset_index().rename(columns={"index": "Date"})
    return out.sort_values("Date")


def tendance_liquidite(vol, date):
    """Volume hors blocs de la semaine de `date` vs semaine précédente (à défaut, 2e moitié vs 1re)."""
    date = pd.Timestamp(date)
    vol = vol[vol["Date"] <= date].copy()
    vol["sem"] = vol["Date"].map(lambda d: pd.Timestamp(d).isocalendar()[:2])
    cur = date.isocalendar()[:2]
    semaines = list(dict.fromkeys(vol["sem"]))
    actuelle = vol[vol["sem"] == cur]["Volume_fond"]
    if len(semaines) >= 2 and semaines[-1] == cur:
        prec = vol[vol["sem"] == semaines[-2]]["Volume_fond"]
        if prec.mean() > 0:
            return (actuelle.mean() / prec.mean() - 1) * 100, "vs semaine précédente"
    v = actuelle.values
    if len(v) >= 2:
        k = len(v) - len(v) // 2
        if v[:k].mean() > 0:
            return (v[k:].mean() / v[:k].mean() - 1) * 100, "fin de semaine vs début"
    return np.nan, ""


def lecture(var_prix, var_liq, seuil_prix=0.25, seuil_liq=10):
    """Grille prix x liquidité -> (titre, phrase, ton)."""
    if np.isnan(var_prix) or np.isnan(var_liq):
        return "Lecture indisponible", "Pas assez de séances pour comparer.", "neutre"
    up, dn = var_prix >= seuil_prix, var_prix <= -seuil_prix
    lu, ld = var_liq >= seuil_liq, var_liq <= -seuil_liq
    if up and lu:
        return "Hausse solide", "L'indice monte et l'argent échangé augmente : les acheteurs sont nombreux.", "ok"
    if up and ld:
        return "Hausse fragile", "L'indice monte mais l'argent échangé diminue : peu d'acheteurs suffisent.", "warn"
    if dn and lu:
        return "Vraie vente", "L'indice baisse avec plus d'argent échangé : des vendeurs sortent en nombre.", "bad"
    if dn and ld:
        return "Baisse par manque d'acheteurs", "L'indice baisse dans un marché qui s'assèche.", "warn"
    if up:
        return "Hausse, activité stable", "L'indice monte avec un volume habituel.", "ok"
    if dn:
        return "Baisse, activité stable", "L'indice baisse avec un volume habituel.", "bad"
    if lu:
        return "Rotation", "L'indice bouge peu mais l'argent échangé augmente : il change de mains ou de valeurs.", "neutre"
    return "Marché calme", "Peu de mouvement sur le prix comme sur le volume.", "neutre"
