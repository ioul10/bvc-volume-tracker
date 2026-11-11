"""
Liquidité par valeur : le ratio de liquidité (ratio Amivest).

    Ratio = argent échangé hors blocs probables (MAD) / somme des |variations quotidiennes du cours en %|
          = MAD qu'il faut échanger pour faire bouger le cours de 1 %
Plus le ratio est élevé, plus l'action est liquide.
Blocs probables : niveaux de prix d'au moins 5 M MAD avec au moins 1 M MAD par contrat.
"""
import io

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from bvc_excel import BLOC_PAR_CONTRAT_MIN, BLOC_VOLUME_MIN

def liquidite(h, vj, nb_seances=20):
    """Classement de toutes les actions par ratio de liquidité sur les N dernières séances.

    Ratio = argent échangé hors blocs (MAD) / somme des |variations quotidiennes en %|
          = MAD qu'il faut échanger pour faire bouger le cours de 1 %.
    Retourne (classement, marché par séance).
    """
    dates = sorted(vj["Date"].unique())[-nb_seances:]
    n = len(dates)
    d = vj[vj["Date"].isin(dates)].copy()
    fond = volume_fond_jour(h, dates)
    phy = h["Physionomie"]
    phy = phy[phy["Date"].isin(dates)]
    bloc = (phy["Volume_MAD"] >= BLOC_VOLUME_MIN) & (phy["Volume_MAD"] / phy["Contrats"] >= BLOC_PAR_CONTRAT_MIN)
    blocs = phy[bloc].groupby(["Date", "Valeur"])["Volume_MAD"].sum().rename("Volume_blocs").reset_index()
    d = d.merge(fond, on=["Date", "Valeur"], how="left").merge(blocs, on=["Date", "Valeur"], how="left")
    d[["Volume_fond", "Volume_blocs"]] = d[["Volume_fond", "Volume_blocs"]].fillna(0)
    d["Mouvement"] = d["Var_pct"].abs().fillna(0)

    g = d.groupby("Valeur").agg(
        Secteur=("Secteur_nom", "last"),
        Argent_MAD=("Volume_fond", "sum"),
        Mouvement_pct=("Mouvement", "sum"),
        Seances_echangees=("Volume_MAD", lambda s: int((s > 0).sum())),
        Contrats=("Contrats", "sum"),
        Cloture=("Cloture", "last"),
    ).reset_index()
    g["Ratio_MAD"] = np.where(g["Mouvement_pct"] > 0, g["Argent_MAD"] / g["Mouvement_pct"].replace(0, np.nan), np.nan)
    g["Argent_par_seance"] = g["Argent_MAD"] / n
    g["Seances"] = n
    # Classée seulement si échangée au moins une séance sur deux (sinon un seul échange sans mouvement
    # de prix donnerait un ratio artificiellement élevé)
    min_seances = int(np.ceil(n / 2))
    g["Mesure"] = np.select(
        [g["Argent_MAD"] <= 0, g["Seances_echangees"] < min_seances, g["Mouvement_pct"] <= 0],
        ["Aucun échange", f"Échangée moins de {min_seances} séances sur {n}", "Prix immobile : non mesurable"], "")
    g.loc[g["Mesure"] != "", "Ratio_MAD"] = np.nan
    g = g.sort_values(["Ratio_MAD", "Argent_MAD"], ascending=False, na_position="last").reset_index(drop=True)
    g.insert(0, "Rang", np.where(g["Ratio_MAD"].notna(), np.arange(1, len(g) + 1), np.nan))

    marche = d.groupby("Date").agg(
        Volume_total=("Volume_MAD", "sum"),
        Volume_fond=("Volume_fond", "sum"),
        Volume_blocs=("Volume_blocs", "sum"),
        Valeurs_traitees=("Volume_MAD", lambda s: int((s > 0).sum())),
        Valeurs_cotees=("Valeur", "count"),
    ).reset_index()
    marche["Largeur_pct"] = marche["Valeurs_traitees"] / marche["Valeurs_cotees"] * 100
    cols = ["Rang", "Valeur", "Secteur", "Ratio_MAD", "Argent_MAD", "Argent_par_seance", "Mouvement_pct",
            "Seances_echangees", "Seances", "Contrats", "Cloture", "Mesure"]
    return g[cols], marche


def volume_fond_jour(h, dates=None):
    """Volume hors grosses transactions, par valeur et par séance."""
    phy = h["Physionomie"]
    if dates is not None:
        phy = phy[phy["Date"].isin(dates)]
    bloc = (phy["Volume_MAD"] >= BLOC_VOLUME_MIN) & (phy["Volume_MAD"] / phy["Contrats"] >= BLOC_PAR_CONTRAT_MIN)
    return phy[~bloc].groupby(["Date", "Valeur"])["Volume_MAD"].sum().rename("Volume_fond").reset_index()


def ajouter_feuille(xlsx_bytes, df, nom="Liquidite"):
    """Ajoute le tableau de liquidité au classeur Excel produit par bvc_excel."""
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    if nom in wb.sheetnames:
        del wb[nom]
    ws = wb.create_sheet(nom, 1)
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append([None if (isinstance(v, float) and np.isnan(v)) else v for v in row])
    fill = PatternFill("solid", fgColor="1F3A5F")
    formats = {"Ratio": "#,##0", "Argent": "#,##0", "_pct": '0.00" %"', "Contrats": "#,##0",
               "Cloture": "#,##0.00", "Rang": "0"}
    for j, col in enumerate(df.columns, start=1):
        c = ws.cell(row=1, column=j)
        c.font, c.fill = Font(bold=True, color="FFFFFF"), fill
        ws.column_dimensions[get_column_letter(j)].width = 26 if col in ("Valeur", "Secteur") else 16
        fmt = next((f for k, f in formats.items() if k in col), None)
        if fmt:
            for i in range(2, ws.max_row + 1):
                ws.cell(row=i, column=j).number_format = fmt
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
