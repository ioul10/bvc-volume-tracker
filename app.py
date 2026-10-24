import io
from datetime import date

import altair as alt
import pandas as pd
import streamlit as st

from bvc_excel import (build_workbook, bulletins_to_frames, merge_history, read_existing)
from bvc_parser import parse_bulletin

st.set_page_config(page_title="BVC Volume Tracker", page_icon="📊", layout="wide")


@st.cache_data(show_spinner=False)
def lire_pdf(nom, contenu):
    return parse_bulletin(io.BytesIO(contenu))


def mm(v):
    return f"{v / 1e6:,.1f} M".replace(",", " ").replace(".", ",")


def m(v):
    return mm(v) + " MAD"


# ------------------------------------------------------------------ barre latérale
st.sidebar.title("📥 Données")
pdfs = st.sidebar.file_uploader("Bulletins de la cote (PDF)", type="pdf", accept_multiple_files=True)
ancien = st.sidebar.file_uploader("Historique Excel (optionnel)", type="xlsx",
                                  help="Le fichier téléchargé la fois précédente : les nouveaux jours s'y ajoutent.")
nb = st.sidebar.slider("Période d'analyse (séances)", 2, 20, 5)
st.sidebar.caption("Les seuils des signaux sont réglés dans bvc_excel.py.")

st.title("📊 BVC Volume Tracker")
st.caption("Bourse de Casablanca · analyse des volumes à partir des bulletins de la cote")

if not pdfs and not ancien:
    st.info("Dépose un ou plusieurs bulletins PDF (BCFR_AAAAMMJJ.pdf) dans la barre de gauche. "
            "Tu peux aussi ajouter l'Excel produit la fois précédente pour continuer l'historique.")
    st.stop()

# ------------------------------------------------------------------ lecture
resultats, controles = [], []
with st.spinner("Lecture des bulletins…"):
    for f in pdfs or []:
        try:
            r = lire_pdf(f.name, f.getvalue())
            resultats.append(r)
            c = r["controle"]
            controles.append({"Fichier": f.name, "Séance": r["date"], "Valeurs": len(r["cotation"]),
                              "Lignes physionomie": len(r["physionomie"]),
                              "Volume lu": c["volume_lu"], "Volume officiel": c["volume_officiel"],
                              "Contrôle": "✅ OK" if c["ok"] else "⚠️ Écart"})
        except Exception as e:  # noqa: BLE001
            controles.append({"Fichier": f.name, "Contrôle": f"❌ Illisible : {e}"})

historique = read_existing(ancien) if ancien else {}
h = merge_history(historique, bulletins_to_frames(resultats) if resultats else {})

if h.get("Marche") is None or h["Marche"].empty:
    st.error("Aucune séance exploitable.")
    st.stop()

if controles:
    with st.expander(f"Contrôle de lecture ({len(controles)} fichier(s))",
                     expanded=any("OK" not in str(c.get("Contrôle")) for c in controles)):
        st.dataframe(pd.DataFrame(controles), hide_index=True, use_container_width=True,
                     column_config={"Volume lu": st.column_config.NumberColumn(format="%.2f"),
                                    "Volume officiel": st.column_config.NumberColumn(format="%.2f")})

nb = min(nb, h["Marche"]["Date"].nunique())
xlsx, t = build_workbook(h, nb)
vj, sig, blocs = t["volumes_jour"], t["signaux"], t["blocs"]

dates = sorted(h["Marche"]["Date"].unique())
periode = dates[-nb:]
mk = h["Marche"].sort_values("Date")
mk_p = mk[mk["Date"].isin(periode)]

st.download_button("⬇️ Télécharger l'Excel", data=xlsx, type="primary",
                   file_name=f"bvc_volumes_{pd.Timestamp(dates[-1]).strftime('%Y%m%d')}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.caption(f"Historique : {len(dates)} séance(s), du {pd.Timestamp(dates[0]):%d/%m/%Y} au {pd.Timestamp(dates[-1]):%d/%m/%Y}. "
           f"Analyse sur les {nb} dernières.")

# ------------------------------------------------------------------ chiffres clés
masi_debut = mk_p["MASI"].iloc[0] / (1 + mk_p["MASI_var_pct"].iloc[0] / 100)
masi_var = (mk_p["MASI"].iloc[-1] / masi_debut - 1) * 100
moy = mk_p["Volume_MAD"].mean()
top5 = sig["Part_periode_pct"].head(5).sum() if not sig.empty else 0
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Volume de la période (MAD)", mm(mk_p["Volume_MAD"].sum()))
c2.metric("Moyenne par séance (MAD)", mm(moy))
c3.metric("Dernière séance (MAD)", mm(mk_p["Volume_MAD"].iloc[-1]), f"{(mk_p['Volume_MAD'].iloc[-1] / moy - 1) * 100:+.0f} % vs moyenne")
c4.metric("Part des 5 premières valeurs", f"{top5:.0f} %")
c5.metric("MASI", f"{mk_p['MASI'].iloc[-1]:,.2f}".replace(",", " "), f"{masi_var:+.2f} % sur la période")

tab1, tab2, tab3, tab4, tab5 = st.tabs(["Marché", "Signaux", "Zoom valeur", "Grosses transactions", "Données"])

# ------------------------------------------------------------------ Marché
with tab1:
    a, b = st.columns(2)
    with a:
        st.subheader("Volume du marché par séance")
        d = mk_p.assign(Volume_M=mk_p["Volume_MAD"] / 1e6, Seance=mk_p["Date"].dt.strftime("%d/%m"))
        bars = alt.Chart(d).mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
            x=alt.X("Seance:N", title=None, sort=None), y=alt.Y("Volume_M:Q", title="M MAD"),
            tooltip=[alt.Tooltip("Seance", title="Séance"), alt.Tooltip("Volume_M", title="Volume (M MAD)", format=",.1f"),
                     alt.Tooltip("Contrats", format=","), alt.Tooltip("MASI_var_pct", title="MASI (%)", format="+.2f")])
        st.altair_chart(bars + bars.mark_text(dy=-8).encode(text=alt.Text("Volume_M:Q", format=",.1f")),
                        use_container_width=True)
    with b:
        st.subheader("Où est allé l'argent")
        top = sig.head(12)
        day_cols = [c for c in sig.columns if "/" in str(c)]
        long = top.melt(id_vars="Valeur", value_vars=day_cols, var_name="Séance", value_name="Volume")
        long["Volume_M"] = long["Volume"] / 1e6
        st.altair_chart(alt.Chart(long).mark_bar().encode(
            y=alt.Y("Valeur:N", sort=list(top["Valeur"]), title=None),
            x=alt.X("sum(Volume_M):Q", title="M MAD"),
            color=alt.Color("Séance:N", sort=day_cols),
            order=alt.Order("Séance:N"),
            tooltip=["Valeur", "Séance", alt.Tooltip("Volume_M", title="M MAD", format=",.1f")]),
            use_container_width=True)
        st.caption("Une couleur dominante = pic ponctuel (souvent un bloc), pas un intérêt régulier.")

# ------------------------------------------------------------------ Signaux
with tab2:
    st.subheader(f"Signaux prix / volume sur {nb} séances")
    filtre = st.multiselect("Filtrer par signal", sorted(sig["Signal"].unique()))
    s = sig[sig["Signal"].isin(filtre)] if filtre else sig
    st.dataframe(
        s[["Valeur", "Secteur", "Volume_periode_MAD", "Part_periode_pct", "Var_periode_pct", "Tendance_volume",
           "Pic_vs_mediane", "Ratio_dernier_jour_vs_moy20", "Ticket_moyen", "Flux", "Signal"]],
        hide_index=True, use_container_width=True, height=520,
        column_config={
            "Volume_periode_MAD": st.column_config.NumberColumn("Volume (MAD)", format="%,.0f"),
            "Part_periode_pct": st.column_config.NumberColumn("Part", format="%.1f %%"),
            "Var_periode_pct": st.column_config.NumberColumn("Var. période", format="%+.1f %%"),
            "Tendance_volume": st.column_config.NumberColumn("Tendance vol.", format="%.2f×"),
            "Pic_vs_mediane": st.column_config.NumberColumn("Pic", format="%.1f×"),
            "Ratio_dernier_jour_vs_moy20": st.column_config.NumberColumn("Dernier jour / moy. 20", format="%.1f×"),
            "Ticket_moyen": st.column_config.NumberColumn("Ticket moyen", format="%,.0f"),
        })
    with st.expander("Comment lire les signaux"):
        st.markdown(
            "- **Hausse confirmée** : le cours monte (≥ +2 %) et le volume augmente (≥ +20 %).\n"
            "- **Hausse sans volume** : le cours monte mais l'activité baisse, la hausse est fragile.\n"
            "- **Baisse sous volume** : le cours baisse et le volume grossit, pression vendeuse réelle.\n"
            "- **Absorption** : gros volume (≥ 20 M MAD) et prix quasi immobile (< 1 %).\n"
            "- **Tendance vol.** : moyenne de la 2ᵉ moitié de la période ÷ moyenne de la 1ʳᵉ moitié.\n"
            "- **Pic** : plus forte séance ÷ séance médiane. **Dernier jour / moy. 20** : volume du dernier jour "
            "÷ moyenne des 20 séances précédentes (fiable à partir de ~20 séances d'historique).\n"
            "- **Flux** : ticket moyen ≥ 250 000 MAD = institutionnel, ≤ 60 000 MAD = particuliers.")

# ------------------------------------------------------------------ Zoom valeur
with tab3:
    choix = st.selectbox("Valeur", sig["Valeur"].tolist())
    r = sig[sig["Valeur"] == choix].iloc[0]
    st.markdown(f"**{choix}** · {m(r['Volume_periode_MAD'])} · {r['Var_periode_pct']:+.1f} % · "
                f"ticket moyen {r['Ticket_moyen'] / 1e3:,.0f} k MAD · **{r['Signal']}**")
    a, b = st.columns(2)
    with a:
        st.subheader("Prix et volume")
        d = vj[(vj["Valeur"] == choix) & (vj["Date"].isin(periode))].copy()
        d["Seance"] = d["Date"].dt.strftime("%d/%m")
        d["Volume_M"] = d["Volume_MAD"] / 1e6
        base = alt.Chart(d).encode(x=alt.X("Seance:N", title=None, sort=None))
        vol = base.mark_bar(opacity=0.6, cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
            y=alt.Y("Volume_M:Q", title="Volume (M MAD)"))
        pmin, pmax = d[["Cloture", "VWAP"]].min().min(), d[["Cloture", "VWAP"]].max().max()
        sc = alt.Scale(domain=[pmin * 0.99, pmax * 1.01])
        close = base.mark_line(point=True, color="#E4572E", strokeWidth=2).encode(
            y=alt.Y("Cloture:Q", title="Cours", scale=sc),
            tooltip=[alt.Tooltip("Seance", title="Séance"), alt.Tooltip("Cloture", title="Clôture", format=",.2f"),
                     alt.Tooltip("VWAP", format=",.2f"), alt.Tooltip("Volume_M", title="Volume (M MAD)", format=",.1f"),
                     alt.Tooltip("Contrats", format=",")])
        vwap = base.mark_tick(color="#555", thickness=2, size=30).encode(y=alt.Y("VWAP:Q", scale=sc))
        st.altair_chart(alt.layer(vol, alt.layer(close, vwap)).resolve_scale(y="independent"),
                        use_container_width=True)
        st.caption("Barres : volume · ligne : clôture · trait gris : VWAP du jour. "
                   "Clôture au-dessus du VWAP = les acheteurs ont fini plus forts.")
    with b:
        st.subheader("Profil de volume par prix")
        p = h["Physionomie"]
        p = p[(p["Valeur"] == choix) & (p["Date"].isin(periode))]
        nbins = int(min(18, p["Prix"].nunique()))
        if nbins > 1:
            p = p.assign(Tranche=pd.cut(p["Prix"], bins=nbins))
            prof = p.groupby("Tranche", observed=True)["Volume_MAD"].sum().reset_index()
            prof["Prix"] = prof["Tranche"].apply(lambda x: f"{x.left:,.2f} – {x.right:,.2f}")
        else:
            prof = p.groupby("Prix", as_index=False)["Volume_MAD"].sum()
            prof["Prix"] = prof["Prix"].map(lambda x: f"{x:,.2f}")
        prof["Volume_M"] = prof["Volume_MAD"] / 1e6
        st.altair_chart(alt.Chart(prof).mark_bar().encode(
            y=alt.Y("Prix:N", sort="descending", title=None), x=alt.X("Volume_M:Q", title="M MAD"),
            tooltip=["Prix", alt.Tooltip("Volume_M", title="M MAD", format=",.2f")]), use_container_width=True)
        st.caption("Les niveaux les plus chargés servent souvent de support ou de résistance.")

# ------------------------------------------------------------------ Blocs
with tab4:
    st.subheader("Grosses transactions (≥ 5 M MAD et ≥ 1 M MAD par contrat)")
    bp = blocs[blocs["Date"].isin(periode)]
    st.dataframe(bp, hide_index=True, use_container_width=True,
                 column_config={"Date": st.column_config.DateColumn(format="DD/MM/YYYY"),
                                "Volume_MAD": st.column_config.NumberColumn(format="%,.0f"),
                                "MAD_par_contrat": st.column_config.NumberColumn(format="%,.0f"),
                                "Part_du_jour_pct": st.column_config.NumberColumn(format="%.0f %%"),
                                "Prix_vs_cloture_pct": st.column_config.NumberColumn(format="%+.2f %%")})
    st.caption("Un bloc payé au-dessus de la clôture montre un acheteur décidé ; en dessous, un vendeur pressé.")

# ------------------------------------------------------------------ Données
with tab5:
    st.subheader("Volumes par valeur et par séance")
    st.dataframe(vj[vj["Date"].isin(periode)], hide_index=True, use_container_width=True, height=520)
