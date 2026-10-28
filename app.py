import html
import io

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

from bvc_excel import build_workbook, bulletins_to_frames, merge_history, read_existing
from bvc_hebdo import charger_composition, donnees_hebdo, libelle_semaine, rapport_word, semaines
from bvc_indices import (INDICES, composition, lecture, liquidite_indice, serie, tendance_liquidite,
                         variation_semaine, volume_indice)
from bvc_liquidite import ajouter_feuille, liquidite
from bvc_parser import parse_bulletin

st.set_page_config(page_title="BVC Volume Tracker", page_icon="📈", layout="wide")

# ====================================================================== style
NAVY, INK, GOLD, GREEN, RED, MUTED = "#0B1F3A", "#13294B", "#C8A24A", "#12A150", "#D9483B", "#6B7A90"

# Motif géométrique (étoile à 8 branches) en SVG pour le bandeau
ZELLIGE = ("data:image/svg+xml;utf8,"
           "<svg xmlns='http://www.w3.org/2000/svg' width='56' height='56' viewBox='0 0 56 56'>"
           "<g fill='none' stroke='%23C8A24A' stroke-opacity='0.22' stroke-width='1'>"
           "<rect x='16' y='16' width='24' height='24'/>"
           "<rect x='16' y='16' width='24' height='24' transform='rotate(45 28 28)'/>"
           "<circle cx='28' cy='28' r='5'/></g></svg>")

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');
html, body, [class*="css"], .stMarkdown, .stDataFrame, button, input {{ font-family: 'Barlow', sans-serif; }}
.block-container {{ padding-top: 3.2rem; max-width: 1500px; }}
[data-testid="stSidebar"] [data-testid="stFileChip"] {{ background:#1B3A66 !important; border:1px solid #2A4A78; }}
[data-testid="stSidebar"] {{ background: {NAVY}; }}
[data-testid="stSidebar"] * {{ color: #E8ECF3 !important; }}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {{ background: #13294B; border: 1px dashed {GOLD}; }}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {{ background:{GOLD}; color:{NAVY} !important; border:none; }}

.bvc-hero {{
  background: linear-gradient(120deg, {NAVY} 0%, #12305A 60%, #0E2748 100%), url("{ZELLIGE}");
  background-blend-mode: normal;
  border-radius: 14px; padding: 22px 28px 18px; color: #fff; position: relative; overflow: hidden;
  border-bottom: 3px solid {GOLD};
}}
.bvc-hero::after {{ content:""; position:absolute; inset:0; background:url("{ZELLIGE}"); opacity:.9; pointer-events:none; }}
.bvc-hero h1 {{ font-weight:700; font-size:2.1rem; margin:0; letter-spacing:.5px; color:#fff; position:relative; z-index:1; }}
.bvc-hero h1 span {{ color:{GOLD}; }}
.bvc-hero p {{ margin:.25rem 0 0; color:#B9C4D6; position:relative; z-index:1; font-size:1rem; }}
.bvc-session {{ position:absolute; right:28px; top:20px; z-index:1; text-align:right; }}
.bvc-session .lbl {{ font-size:.75rem; letter-spacing:2px; text-transform:uppercase; color:{GOLD}; }}
.bvc-session .val {{ font-family:'JetBrains Mono', monospace; font-size:1.35rem; color:#fff; }}
.bvc-session .dot {{ display:inline-block; width:8px; height:8px; border-radius:50%; background:{GREEN};
                     margin-right:6px; box-shadow:0 0 0 0 rgba(18,161,80,.6); animation:pulse 2s infinite; }}
@keyframes pulse {{ 0% {{box-shadow:0 0 0 0 rgba(18,161,80,.6);}} 70% {{box-shadow:0 0 0 10px rgba(18,161,80,0);}} 100% {{box-shadow:0 0 0 0 rgba(18,161,80,0);}} }}

.bvc-ticker {{ background:#07162B; border-radius:10px; overflow:hidden; white-space:nowrap; margin:10px 0 18px;
               border:1px solid #1C3558; }}
.bvc-ticker .track {{ display:inline-block; padding:9px 0; animation: scroll 70s linear infinite; }}
.bvc-ticker:hover .track {{ animation-play-state: paused; }}
.bvc-ticker .it {{ font-family:'JetBrains Mono', monospace; font-size:.9rem; color:#DCE3EE; margin:0 22px; }}
.bvc-ticker .nm {{ color:#fff; font-weight:600; margin-right:8px; }}
.bvc-ticker .up {{ color:#3DDC84; }} .bvc-ticker .dn {{ color:#FF6B5E; }} .bvc-ticker .eq {{ color:#9AA8BC; }}
@keyframes scroll {{ from {{ transform: translateX(0); }} to {{ transform: translateX(-50%); }} }}
@media (prefers-reduced-motion: reduce) {{ .bvc-ticker .track {{ animation:none; }} }}

.kpis {{ display:grid; grid-template-columns:repeat(auto-fit, minmax(190px,1fr)); gap:12px; margin-bottom:8px; }}
.kpi {{ background:#fff; border:1px solid #E3E8F0; border-radius:12px; padding:14px 16px; border-top:3px solid {INK}; }}
.kpi.gold {{ border-top-color:{GOLD}; }}
.kpi .l {{ font-size:.78rem; text-transform:uppercase; letter-spacing:1px; color:{MUTED}; }}
.kpi .v {{ font-family:'JetBrains Mono', monospace; font-size:1.55rem; font-weight:600; color:{INK}; margin-top:4px; }}
.kpi .s {{ font-size:.85rem; margin-top:2px; color:{MUTED}; }}
.kpi .v.txt {{ font-family:'Barlow', sans-serif; font-size:1.25rem; }}
.kpi .s.up {{ color:{GREEN}; }} .kpi .s.dn {{ color:{RED}; }}

.heat {{ display:flex; flex-wrap:wrap; gap:4px; }}
.tile {{ border-radius:6px; padding:8px 10px; color:#fff; min-width:110px; flex:1 1 110px; display:flex;
         flex-direction:column; justify-content:space-between; min-height:70px; }}
.tile.big {{ flex:2 1 220px; min-height:110px; }}
.tile .n {{ font-weight:700; font-size:.85rem; line-height:1.1; }}
.tile.big .n {{ font-size:1.05rem; }}
.tile .p {{ font-family:'JetBrains Mono', monospace; font-size:.95rem; }}
.tile.big .p {{ font-size:1.3rem; }}
.tile .vol {{ font-size:.72rem; opacity:.85; }}

.stTabs [data-baseweb="tab-list"] {{ gap:4px; border-bottom:2px solid #E3E8F0; }}
.stTabs [data-baseweb="tab"] {{ background:#fff; border-radius:8px 8px 0 0; padding:8px 16px; font-weight:600; }}
.stTabs [aria-selected="true"] {{ background:{NAVY}; color:#fff !important; }}
.stTabs [aria-selected="true"] p {{ color:#fff !important; }}
.stDownloadButton button {{ background:{GOLD}; color:{NAVY}; border:none; font-weight:700; }}
.stDownloadButton button:hover {{ background:#B38E37; color:#fff; }}
h3 {{ color:{INK}; font-weight:700 !important; }}
.idx {{ background:#fff; border:1px solid #E3E8F0; border-radius:14px; padding:16px 18px 10px; border-top:4px solid {NAVY}; }}
.idx.m20 {{ border-top-color:{GOLD}; }}
.idx .nm {{ font-weight:700; font-size:1.15rem; color:{INK}; letter-spacing:.5px; }}
.idx .ds {{ font-size:.8rem; color:{MUTED}; margin-bottom:6px; }}
.idx .lv {{ font-family:'JetBrains Mono', monospace; font-size:2rem; font-weight:600; color:{INK}; }}
.idx .vr {{ display:flex; gap:14px; flex-wrap:wrap; font-size:.85rem; color:{MUTED}; margin-top:2px; }}
.idx .vr b {{ font-family:'JetBrains Mono', monospace; font-weight:600; }}
.up-t {{ color:{GREEN}; }} .dn-t {{ color:{RED}; }}
.liq {{ background:#F7F9FC; border:1px solid #E3E8F0; border-radius:12px; padding:12px 16px; margin-top:6px; }}
.liq .t {{ font-size:.75rem; text-transform:uppercase; letter-spacing:1px; color:{MUTED}; }}
.liq .sc {{ font-family:'JetBrains Mono', monospace; font-size:1.6rem; font-weight:600; color:{INK}; }}
.liq .cl {{ font-size:.85rem; font-weight:600; margin-left:6px; }}
.gauge {{ height:8px; border-radius:6px; background:#E3E8F0; margin:6px 0 10px; overflow:hidden; }}
.gauge i {{ display:block; height:100%; border-radius:6px; }}
.liq .row {{ display:flex; justify-content:space-between; gap:8px; font-size:.85rem; color:{INK}; padding:3px 0;
             border-top:1px dashed #E3E8F0; }}
.liq .row span:last-child {{ font-family:'JetBrains Mono', monospace; font-weight:600; text-align:right; }}
.verdict {{ border-radius:12px; padding:14px 18px; border-left:6px solid; background:#fff; border:1px solid #E3E8F0; }}
.verdict.ok {{ border-left:6px solid {GREEN}; }} .verdict.bad {{ border-left:6px solid {RED}; }}
.verdict.warn {{ border-left:6px solid {GOLD}; }} .verdict.neutre {{ border-left:6px solid {MUTED}; }}
.verdict .k {{ font-size:.75rem; text-transform:uppercase; letter-spacing:1px; color:{MUTED}; }}
.verdict .h {{ font-size:1.25rem; font-weight:700; color:{INK}; margin:2px 0; }}
.verdict .p {{ font-size:.92rem; color:{INK}; }}
.foot {{ color:{MUTED}; font-size:.8rem; text-align:center; margin-top:30px; }}
</style>
""", unsafe_allow_html=True)



# ====================================================================== utilitaires
def fr(v, d=1):
    return f"{v:,.{d}f}".replace(",", " ").replace(".", ",")


def mm(v):
    return fr(v / 1e6) + " M"


def pct(v, d=2):
    return ("+" if v > 0 else "") + fr(v, d) + " %"


def couleur(var):
    """Rouge -> gris -> vert selon la variation (saturée à ±4 %)."""
    if var is None or np.isnan(var):
        return "#5A6B82"
    x = max(-4, min(4, var)) / 4
    neutre = np.array([90, 107, 130])
    cible = np.array([18, 161, 80]) if x > 0 else np.array([217, 72, 59])
    c = neutre + (cible - neutre) * abs(x) ** 0.7
    return "#%02x%02x%02x" % tuple(int(v) for v in c)


def kpi(label, value, sub="", cls="", gold=False, txt=False):
    return (f'<div class="kpi{" gold" if gold else ""}"><div class="l">{label}</div>'
            f'<div class="v{" txt" if txt else ""}">{value}</div><div class="s {cls}">{sub}</div></div>')


@st.cache_data(show_spinner=False)
def lire_pdf(nom, contenu):
    return parse_bulletin(io.BytesIO(contenu))


# ====================================================================== barre latérale
st.sidebar.markdown("## 📥 Données")
pdfs = st.sidebar.file_uploader("Bulletins de la cote (PDF)", type="pdf", accept_multiple_files=True)
ancien = st.sidebar.file_uploader("Historique Excel (optionnel)", type="xlsx",
                                  help="Le fichier téléchargé la fois précédente : les nouveaux jours s'y ajoutent.")
st.sidebar.markdown("## ⚙️ Réglages")
nb = st.sidebar.slider("Période des signaux (séances)", 2, 20, 5)
nb_liq = st.sidebar.slider("Fenêtre de liquidité (séances)", 5, 60, 20)
st.sidebar.caption("Seuils des signaux : bvc_excel.py · poids du score de liquidité : bvc_liquidite.py")


def hero(seance=None):
    right = ""
    if seance is not None:
        right = (f'<div class="bvc-session"><div class="lbl">Dernière séance</div>'
                 f'<div class="val"><span class="dot"></span>{pd.Timestamp(seance):%d/%m/%Y}</div></div>')
    st.markdown(f"""<div class="bvc-hero">{right}
      <h1>BVC <span>Volume</span> Tracker</h1>
      <p>Volumes, signaux et liquidité · Bourse de Casablanca · d'après les bulletins de la cote</p>
    </div>""", unsafe_allow_html=True)


if not pdfs and not ancien:
    hero()
    st.markdown("<br>", unsafe_allow_html=True)
    st.info("👈 Dépose un ou plusieurs bulletins PDF (BCFR_AAAAMMJJ.pdf) dans la barre de gauche. "
            "Ajoute l'Excel de la fois précédente pour continuer ton historique.")
    st.stop()

# ====================================================================== lecture
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

h = merge_history(read_existing(ancien) if ancien else {}, bulletins_to_frames(resultats) if resultats else {})
if h.get("Marche") is None or h["Marche"].empty:
    hero()
    st.error("Aucune séance exploitable.")
    st.stop()

dates = sorted(h["Marche"]["Date"].unique())
nb = min(nb, len(dates))
xlsx, t = build_workbook(h, nb)
vj, sig, blocs = t["volumes_jour"], t["signaux"], t["blocs"]
liq, liq_m = liquidite(h, vj, nb_liq)
xlsx = ajouter_feuille(xlsx, liq)

periode = dates[-nb:]
mk = h["Marche"].sort_values("Date")
mk_p = mk[mk["Date"].isin(periode)]
last = vj[vj["Date"] == dates[-1]]

# ====================================================================== en-tête + bandeau
hero(dates[-1])
tick = last[last["Volume_MAD"] > 0].sort_values("Volume_MAD", ascending=False)
items = "".join(
    f'<span class="it"><span class="nm">{html.escape(r.Valeur)}</span>{fr(r.Cloture, 2)} '
    f'<span class="{"up" if r.Var_pct > 0 else "dn" if r.Var_pct < 0 else "eq"}">'
    f'{"▲" if r.Var_pct > 0 else "▼" if r.Var_pct < 0 else "■"} {pct(r.Var_pct)}</span></span>'
    for r in tick.itertuples())
tete = ""
for _nom in INDICES:
    _s = serie(h, _nom)
    if len(_s):
        _r = _s.iloc[-1]
        tete += (f'<span class="it"><span class="nm" style="color:{GOLD}">{_nom}</span>{fr(_r.Valeur, 2)} '
                 f'<span class="{"up" if _r.Var_jour_pct >= 0 else "dn"}">{pct(_r.Var_jour_pct)}</span></span>')
items = tete + items
st.markdown(f'<div class="bvc-ticker"><div class="track">{items}{items}</div></div>', unsafe_allow_html=True)

if controles:
    ok = all("OK" in str(c.get("Contrôle")) for c in controles)
    with st.expander(f"{'✅' if ok else '⚠️'} Contrôle de lecture · {len(controles)} fichier(s)", expanded=not ok):
        st.dataframe(pd.DataFrame(controles), hide_index=True, use_container_width=True,
                     column_config={"Volume lu": st.column_config.NumberColumn(format="%.2f"),
                                    "Volume officiel": st.column_config.NumberColumn(format="%.2f")})

# ====================================================================== téléchargement
lm = liq_m.iloc[-1]
st.sidebar.markdown("## 📤 Exporter")
st.sidebar.download_button("⬇️ Excel complet", data=xlsx, use_container_width=True,
                           file_name=f"bvc_volumes_{pd.Timestamp(dates[-1]):%Y%m%d}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.sidebar.caption(f"{len(dates)} séance(s) en historique, du {pd.Timestamp(dates[0]):%d/%m} au "
                   f"{pd.Timestamp(dates[-1]):%d/%m/%Y}. Réutilise ce fichier la prochaine fois pour continuer.")

tabs = st.tabs(["🏛️ Indices", "📰 Hebdo MASI 20", "🚦 Signaux", "💧 Liquidité", "🔎 Zoom valeur",
                "🧱 Grosses transactions", "🗺️ Carte du marché", "📊 Marché"])

# ====================================================================== Indices (accueil)
LIQ_COUL = {"Très liquide": GREEN, "Liquide": "#3E7CB1", "Peu liquide": GOLD, "Illiquide": RED}


def classe_score(x):
    return "Très liquide" if x >= 75 else "Liquide" if x >= 50 else "Peu liquide" if x >= 25 else "Illiquide"


def popup_indice(nom, L):
    """Contenu de la fenêtre d'explication de la liquidité d'un indice."""
    st.markdown(f"#### Liquidité du {nom} : {fr(L['score'], 0)} / 100")
    st.markdown(
        "**Comment on obtient ce chiffre**\n\n"
        "1. Chaque valeur reçoit un score de liquidité sur 100, fait de 4 ingrédients.\n"
        f"2. On fait la moyenne de ces scores, **pondérée par le poids de chaque valeur dans le {nom}**.\n\n"
        "Une grosse valeur peu liquide fait donc baisser la note de l'indice.")
    ing = pd.DataFrame({
        "Ingrédient": ["Argent échangé (hors grosses transactions)", "Régularité (jours avec échanges)",
                       "Écart entre acheteurs et vendeurs", "Coût pour faire bouger le cours"],
        "Points": [L["ingredients"]["volume"], L["ingredients"]["frequence"], L["ingredients"]["ecart"],
                   L["ingredients"]["impact"]],
        "Maximum": [30, 20, 20, 30]})
    ing["Lbl"] = ing.apply(lambda r: f"{fr(r.Points, 0)} / {r.Maximum}", axis=1)
    base = alt.Chart(ing).encode(y=alt.Y("Ingrédient:N", sort=None, title=None, axis=alt.Axis(labelLimit=260)))
    st.altair_chart(alt.layer(
        base.mark_bar(color="#E3E8F0", cornerRadius=4).encode(x=alt.X("Maximum:Q", title=None, axis=None)),
        base.mark_bar(color=GOLD, cornerRadius=4).encode(x="Points:Q"),
        base.mark_text(align="left", dx=4, color=INK).encode(x="Maximum:Q", text="Lbl:N"),
    ).properties(height=170), use_container_width=True)
    st.markdown("**Les valeurs qui pèsent le plus dans la note**")
    det = L["detail"].head(6)
    st.dataframe(det[["Valeur", "Poids_pct", "Score", "Apport_score"]], hide_index=True, use_container_width=True,
                 column_config={"Poids_pct": st.column_config.NumberColumn("Poids", format="%.1f %%"),
                                "Score": st.column_config.NumberColumn("Score valeur", format="%.0f"),
                                "Apport_score": st.column_config.NumberColumn("Apport à la note", format="%.1f pts")})
    st.markdown(
        f"**Coût pour bouger l'indice de 1 %** : {fr(L['cout_1pct'] / 1e6, 0)} M MAD. "
        "Pour chaque valeur, on calcule combien de MAD il faut échanger pour que son cours bouge de 1 % "
        "(variation ÷ argent échangé, en médiane sur la période), puis on additionne. "
        "Plus ce montant est élevé, plus l'indice est profond.")
    notes = [f"Poids {L['source_poids']} : nombre de titres × cours, sans flottant ni plafonnement."
             if L["source_poids"] == "approchés" else "Poids officiels issus du fichier de composition."]
    if L["n_sans_impact"]:
        notes.append(f"{L['n_sans_impact']} valeur(s) sans échange sur la période comptent pour 0.")
    st.caption(" ".join(notes))


with tabs[0]:
    date = dates[-1]
    st.markdown(f"### Les indices · séance du {pd.Timestamp(date):%d/%m/%Y}")
    cols = st.columns(3)
    verdicts = []
    for col, nom in zip(cols, INDICES):
        with col:
            s_ = serie(h, nom)
            if s_.empty:
                st.warning(f"{nom} introuvable dans les bulletins.")
                continue
            r_ = s_.iloc[-1]
            vs, nj = variation_semaine(s_, date)
            cl = lambda v: "up-t" if v >= 0 else "dn-t"  # noqa: E731
            st.markdown(
                f'<div class="idx{" m20" if nom == "MASI 20" else ""}"><div class="nm">{nom}</div>'
                f'<div class="ds">{INDICES[nom]["texte"]}</div>'
                f'<div class="lv">{fr(r_.Valeur, 2)}</div>'
                f'<div class="vr"><span>Jour <b class="{cl(r_.Var_jour_pct)}">{pct(r_.Var_jour_pct)}</b></span>'
                f'<span>Semaine <b class="{cl(vs)}">{pct(vs)}</b></span>'
                f'<span>Depuis le 1er janvier <b class="{cl(r_.Var_annee_pct)}">{pct(r_.Var_annee_pct)}</b></span></div>'
                f'</div>', unsafe_allow_html=True)
            if len(s_) > 1:
                sp = s_.assign(Seance=s_["Date"].dt.strftime("%d/%m"))
                st.altair_chart(alt.Chart(sp).mark_line(
                    color=GOLD if nom == "MASI 20" else INK, strokeWidth=2.5,
                    point=alt.OverlayMarkDef(size=35, color=GOLD if nom == "MASI 20" else INK)).encode(
                    x=alt.X("Seance:N", sort=None, title=None, axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("Valeur:Q", scale=alt.Scale(zero=False), title=None, axis=alt.Axis(tickCount=3)),
                    tooltip=[alt.Tooltip("Seance", title="Séance"), alt.Tooltip("Valeur", format=",.2f"),
                             alt.Tooltip("Var_jour_pct", title="Var. %", format="+.2f")]).properties(height=110),
                    use_container_width=True)

            compo_ = composition(h, nom)
            compo_ = compo_[compo_["Valeur"].isin(set(vj["Valeur"]))]
            if compo_.empty:
                st.markdown(
                    f'<div class="liq"><div class="t">Liquidité de l\'indice</div>'
                    f'<div class="row"><span>Composition non renseignée</span><span>—</span></div></div>',
                    unsafe_allow_html=True)
                st.caption(f"Ajoute la liste de ses valeurs dans {INDICES[nom]['fichier']} pour calculer sa liquidité.")
                continue
            L = liquidite_indice(h, liq, compo_, date)
            vol = volume_indice(h, vj, compo_, dates)
            tl, tl_txt = tendance_liquidite(vol, date)
            vday = vol[vol["Date"] == date]
            cls_ = classe_score(L["score"])
            st.markdown(
                f'<div class="liq"><div class="t">Liquidité de l\'indice · {nb_liq if nb_liq <= len(dates) else len(dates)} séances</div>'
                f'<div><span class="sc">{fr(L["score"], 0)}</span><span style="color:{MUTED}"> / 100</span>'
                f'<span class="cl" style="color:{LIQ_COUL[cls_]}">{cls_}</span></div>'
                f'<div class="gauge"><i style="width:{L["score"]:.0f}%;background:{LIQ_COUL[cls_]}"></i></div>'
                f'<div class="row"><span>Coût pour bouger de 1 %</span><span>{fr(L["cout_1pct"] / 1e6, 0)} M</span></div>'
                f'<div class="row"><span>Argent échangé ce jour*</span><span>{mm(vday["Volume_fond"].sum())}</span></div>'
                f'<div class="row"><span>Tendance de l\'argent échangé*</span>'
                f'<span class="{cl(tl) if not np.isnan(tl) else ""}">{pct(tl, 0) if not np.isnan(tl) else "—"}</span></div>'
                f'</div>', unsafe_allow_html=True)
            st.caption(f"Montants en MAD. *Hors grosses transactions. Tendance : {tl_txt}.")
            with st.popover("ⓘ Comment est calculée cette liquidité ?", use_container_width=True):
                popup_indice(nom, L)
            titre, phrase, ton = lecture(vs, tl)
            verdicts.append((nom, titre, phrase, ton, vs, tl))

    if verdicts:
        st.markdown("### Ce que disent le prix et la liquidité cette semaine")
        vc = st.columns(len(verdicts) + 1)
        for c_, (nom, titre, phrase, ton, vs, tl) in zip(vc, verdicts):
            c_.markdown(f'<div class="verdict {ton}"><div class="k">{nom} · {pct(vs)} · argent échangé {pct(tl, 0)}</div>'
                        f'<div class="h">{titre}</div><div class="p">{phrase}</div></div>', unsafe_allow_html=True)
        with vc[-1].popover("ⓘ Comment lire cette grille", use_container_width=True):
            st.markdown(
                "On croise **le sens de l'indice** sur la semaine et **l'évolution de l'argent échangé** "
                "sur ses valeurs (hors grosses transactions).")
            st.markdown(
                "| | Argent échangé en hausse | Argent échangé en baisse |\n|---|---|---|\n"
                "| **Indice en hausse** | 🟢 Hausse solide | 🟡 Hausse fragile |\n"
                "| **Indice en baisse** | 🔴 Vraie vente | 🟡 Baisse par manque d'acheteurs |")
            st.caption("Seuils : ±0,25 % pour l'indice, ±10 % pour l'argent échangé. Entre les deux : activité stable, "
                       "rotation ou marché calme. La comparaison se fait avec la semaine précédente quand elle est "
                       "disponible, sinon entre la fin et le début de la semaine.")

    st.markdown("### L'indice et l'argent qui le fait bouger")
    choix_idx = st.radio("Indice", [n for n in INDICES if not composition(h, n).empty], horizontal=True,
                         label_visibility="collapsed")
    compo_c = composition(h, choix_idx)
    vol = volume_indice(h, vj, compo_c, dates)
    vol["Grosses transactions"] = (vol["Volume"] - vol["Volume_fond"]) / 1e6
    vol["Échanges courants"] = vol["Volume_fond"] / 1e6
    s_c = serie(h, choix_idx)
    g = vol.merge(s_c, on="Date").assign(Seance=lambda x: x["Date"].dt.strftime("%d/%m"))
    long = g.melt(id_vars=["Seance"], value_vars=["Échanges courants", "Grosses transactions"], var_name="Type", value_name="M")
    bars = alt.Chart(long).mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(
        x=alt.X("Seance:N", sort=None, title=None, axis=alt.Axis(labelAngle=0)),
        y=alt.Y("sum(M):Q", title="Argent échangé sur ses valeurs (M MAD)"),
        color=alt.Color("Type:N", scale=alt.Scale(domain=["Échanges courants", "Grosses transactions"], range=[INK, GOLD]),
                        legend=alt.Legend(orient="top", title=None)),
        tooltip=["Seance", "Type", alt.Tooltip("M", title="M MAD", format=",.1f")])
    coul_l = RED if g["Var_jour_pct"].sum() < 0 else GREEN
    line = alt.Chart(g).mark_line(color=coul_l, strokeWidth=3,
                                  point=alt.OverlayMarkDef(size=60, color=coul_l)).encode(
        x=alt.X("Seance:N", sort=None), y=alt.Y("Valeur:Q", scale=alt.Scale(zero=False), title=choix_idx),
        tooltip=["Seance", alt.Tooltip("Valeur", title=choix_idx, format=",.2f"),
                 alt.Tooltip("Var_jour_pct", title="Var. %", format="+.2f")])
    st.altair_chart(alt.layer(bars, line).resolve_scale(y="independent").properties(height=320), use_container_width=True)
    st.caption("Barres : argent échangé chaque séance sur les valeurs de l'indice. En bleu les échanges courants, "
               "en or les grosses transactions entre institutionnels. Ligne : niveau de l'indice.")

# ====================================================================== Hebdo MASI 20
with tabs[1]:
    compo_def = charger_composition()
    sem = semaines(dates)
    cles = list(sem)
    a, b = st.columns([2, 3])
    with a:
        k = st.selectbox("Semaine", cles[::-1], format_func=lambda x: libelle_semaine(sem[x]))
    with b:
        presentes = set(vj["Valeur"])
        membres = st.multiselect(
            "Composition du MASI 20", sorted(presentes),
            default=[v for v in compo_def["Valeur"] if v in presentes],
            help="Par défaut : composition issue de la révision d'octobre 2025 (fichier masi20.csv). "
                 "Vérifie-la après chaque révision annuelle.")
    if len(membres) < 2:
        st.warning("Choisis au moins deux valeurs.")
    else:
        compo = pd.DataFrame({"Valeur": membres}).merge(compo_def, on="Valeur", how="left")
        i = cles.index(k)
        try:
            data = donnees_hebdo(h, vj, liq, compo, sem[k], sem[cles[i - 1]] if i > 0 else None)
        except (IndexError, KeyError) as e:
            st.error(f"Impossible de construire la revue : l'indice MASI 20 est introuvable dans les bulletins ({e}).")
            data = None
        if data:
            R, T = data["resume"], data["valeurs"]
            st.markdown('<div class="kpis">' +
                        kpi("MASI 20", fr(R["m20_fin"], 2), pct(R["m20_var"]) + " sur la semaine",
                            "up" if R["m20_var"] >= 0 else "dn", gold=True) +
                        kpi("MASI", pct(R["masi_var"]), "même semaine", "up" if R["masi_var"] >= 0 else "dn") +
                        kpi("Hausses / baisses", f"{R['hausses']} / {R['baisses']}", f"sur {len(T)} valeurs") +
                        kpi("Volume MASI 20", mm(R["vol_m20"]), f"{fr(R['part_m20'], 0)} % du marché") +
                        kpi("Contrôle des poids", pct(R["reconstitue"]), f"reconstitué vs {pct(R['m20_var'])} publié") +
                        '</div>', unsafe_allow_html=True)
            c1, c2 = st.columns([3, 1])
            with c1:
                st.markdown("#### L'essentiel de la semaine")
                st.markdown("\n".join(f"- {m}" for m in R["messages"]))
            with c2:
                st.download_button("📄 Revue Word (A4)", data=rapport_word(data), use_container_width=True,
                                   file_name=f"MASI20_hebdo_{R['jours'][-1]:%Y%m%d}.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                st.caption(R["source_poids"].capitalize() + ".")
            a, b = st.columns(2)
            with a:
                st.markdown("#### Qui a fait bouger l'indice")
                cc = T.assign(pb=T["Contribution_pct"] * 100,
                              Volume=np.where(T["Confirme"], "Volume en hausse", "Volume stable ou en baisse"))
                st.altair_chart(alt.Chart(cc).mark_bar(cornerRadiusTopRight=3, cornerRadiusBottomRight=3).encode(
                    y=alt.Y("Valeur:N", sort="-x", title=None), x=alt.X("pb:Q", title="Contribution (points de base)"),
                    color=alt.condition(alt.datum.pb > 0, alt.value(GREEN), alt.value(RED)),
                    opacity=alt.Opacity("Volume:N", scale=alt.Scale(domain=["Volume en hausse", "Volume stable ou en baisse"],
                                                                    range=[1, .4]), legend=alt.Legend(orient="bottom", title=None)),
                    tooltip=["Valeur", alt.Tooltip("Poids_pct", title="Poids %", format=".1f"),
                             alt.Tooltip("Var_semaine_pct", title="Var. %", format="+.2f"),
                             alt.Tooltip("pb", title="Contribution (pb)", format="+.1f"), "Signal"]
                ).properties(height=480), use_container_width=True)
            with b:
                st.markdown("#### Où va l'argent : volume et poids")
                mx = float(max(T["Poids_pct"].max(), T["Part_volume_pct"].max()) * 1.1)
                diag = alt.Chart(pd.DataFrame({"x": [0, mx], "y": [0, mx]})).mark_line(strokeDash=[4, 4], color="#C9D2DE").encode(
                    x="x:Q", y="y:Q")
                pts = alt.Chart(T.assign(Sens=np.where(T["Var_semaine_pct"] >= 0, "Hausse", "Baisse"))).mark_circle(
                    size=110, stroke="#fff", strokeWidth=1).encode(
                    x=alt.X("Poids_pct:Q", title="Poids dans l'indice (%)", scale=alt.Scale(domain=[0, mx])),
                    y=alt.Y("Part_volume_pct:Q", title="Part du volume de l'indice (%)", scale=alt.Scale(domain=[0, mx])),
                    color=alt.Color("Sens:N", scale=alt.Scale(domain=["Hausse", "Baisse"], range=[GREEN, RED]),
                                    legend=alt.Legend(orient="bottom", title=None)),
                    tooltip=["Valeur", alt.Tooltip("Poids_pct", title="Poids %", format=".1f"),
                             alt.Tooltip("Part_volume_pct", title="Part du volume %", format=".1f"),
                             alt.Tooltip("Volume_vs_poids", title="Volume / poids", format=".1f")])
                st.altair_chart((diag + pts).properties(height=480), use_container_width=True)
                st.caption("Au-dessus de la diagonale, la valeur attire plus d'argent que son poids dans l'indice.")
            st.markdown("#### À surveiller")
            st.markdown("\n".join(f"- {m}" for m in R["surveiller"]) or "Rien de particulier.")
            with st.expander("Détail des valeurs"):
                st.dataframe(T[["Valeur", "Secteur", "Poids_pct", "Var_semaine_pct", "Contribution_pct", "Volume_MAD",
                                "Part_volume_pct", "Volume_vs_poids", "Tendance_volume", "Score_liquidite", "Signal"]],
                             hide_index=True, use_container_width=True,
                             column_config={"Poids_pct": st.column_config.NumberColumn("Poids", format="%.1f %%"),
                                            "Var_semaine_pct": st.column_config.NumberColumn("Var.", format="%+.2f %%"),
                                            "Contribution_pct": st.column_config.NumberColumn("Contribution", format="%+.3f %%"),
                                            "Volume_MAD": st.column_config.NumberColumn("Volume", format="%,.0f"),
                                            "Part_volume_pct": st.column_config.NumberColumn("Part du volume", format="%.1f %%"),
                                            "Volume_vs_poids": st.column_config.NumberColumn("Volume / poids", format="%.1f×"),
                                            "Tendance_volume": st.column_config.NumberColumn("Tend. vol.", format="%.2f×"),
                                            "Score_liquidite": st.column_config.NumberColumn("Liquidité", format="%.0f")})

# ====================================================================== Carte du marché
with tabs[6]:
    a, b = st.columns([3, 1])
    with a:
        seance = st.select_slider("Séance", options=dates, value=dates[-1],
                                  format_func=lambda d: pd.Timestamp(d).strftime("%a %d/%m"))
    with b:
        mode = st.radio("Taille des tuiles", ["Volume", "Uniforme"], horizontal=True)
    j = vj[(vj["Date"] == seance) & (vj["Volume_MAD"] > 0)].sort_values("Volume_MAD", ascending=False)
    tot = j["Volume_MAD"].sum()
    tiles = []
    for i, r in enumerate(j.itertuples()):
        big = mode == "Volume" and i < 8
        grow = max(1, round(r.Volume_MAD / tot * 40)) if mode == "Volume" else 1
        tiles.append(
            f'<div class="tile{" big" if big else ""}" style="background:{couleur(r.Var_pct)};flex-grow:{grow}" '
            f'title="{html.escape(r.Valeur)} · clôture {fr(r.Cloture, 2)} · {r.Contrats:.0f} contrats">'
            f'<div class="n">{html.escape(r.Valeur)}</div><div class="p">{pct(r.Var_pct)}</div>'
            f'<div class="vol">{mm(r.Volume_MAD)} MAD · {fr(r.Volume_MAD / tot * 100, 1)} %</div></div>')
    st.markdown(f'<div class="heat">{"".join(tiles)}</div>', unsafe_allow_html=True)
    nt = int((vj["Date"] == seance).sum() - len(j))
    st.caption(f"{len(j)} valeurs traitées, {nt} non traitées. Couleur = variation du jour (saturée à ±4 %). "
               "Les plus grandes tuiles concentrent le volume.")

# ====================================================================== Marché
with tabs[7]:
    a, b = st.columns(2)
    with a:
        st.subheader("Volume du marché par séance")
        lm_p = liq_m[liq_m["Date"].isin(periode)].copy()
        lm_p["Seance"] = lm_p["Date"].dt.strftime("%d/%m")
        long = lm_p.melt(id_vars=["Seance"], value_vars=["Volume_fond", "Volume_blocs"], var_name="Type", value_name="v")
        long["Type"] = long["Type"].map({"Volume_fond": "Volume de fond", "Volume_blocs": "Grosses transactions"})
        long["M"] = long["v"] / 1e6
        st.altair_chart(alt.Chart(long).mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(
            x=alt.X("Seance:N", title=None, sort=None), y=alt.Y("sum(M):Q", title="M MAD"),
            color=alt.Color("Type:N", scale=alt.Scale(domain=["Volume de fond", "Grosses transactions"], range=[INK, GOLD]), legend=alt.Legend(orient="top", title=None)),
            tooltip=["Seance", "Type", alt.Tooltip("M", title="M MAD", format=",.1f")]), use_container_width=True)
    with b:
        st.subheader("MASI et volume")
        mk_p2 = mk_p.assign(Seance=mk_p["Date"].dt.strftime("%d/%m"), M=mk_p["Volume_MAD"] / 1e6)
        base = alt.Chart(mk_p2).encode(x=alt.X("Seance:N", title=None, sort=None))
        st.altair_chart(alt.layer(
            base.mark_bar(color="#D5DDE8", cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(y=alt.Y("M:Q", title="Volume (M MAD)")),
            base.mark_line(color=GOLD, strokeWidth=3, point=alt.OverlayMarkDef(color=GOLD, size=60)).encode(
                y=alt.Y("MASI:Q", scale=alt.Scale(zero=False), title="MASI"),
                tooltip=["Seance", alt.Tooltip("MASI", format=",.2f"), alt.Tooltip("MASI_var_pct", title="Var. %", format="+.2f"),
                         alt.Tooltip("M", title="Volume M MAD", format=",.1f")])
        ).resolve_scale(y="independent"), use_container_width=True)
    st.subheader("Où est allé l'argent")
    top = sig.head(15)
    day_cols = [c for c in sig.columns if "/" in str(c)]
    long = top.melt(id_vars="Valeur", value_vars=day_cols, var_name="Séance", value_name="Volume")
    long["M"] = long["Volume"] / 1e6
    st.altair_chart(alt.Chart(long).mark_bar().encode(
        y=alt.Y("Valeur:N", sort=list(top["Valeur"]), title=None),
        x=alt.X("sum(M):Q", title="M MAD"),
        color=alt.Color("Séance:N", sort=day_cols, legend=alt.Legend(orient="top")),
        order=alt.Order("Séance:N"),
        tooltip=["Valeur", "Séance", alt.Tooltip("M", title="M MAD", format=",.1f")]).properties(height=420),
        use_container_width=True)
    st.caption("Une couleur dominante = pic ponctuel (souvent un bloc), pas un intérêt régulier.")

# ====================================================================== Signaux
with tabs[2]:
    st.subheader(f"Signaux prix / volume sur {nb} séances")
    counts = sig["Signal"].value_counts()
    emo = {"Hausse confirmée": "🟢", "Baisse sous volume": "🔴", "Hausse sans volume": "🟡",
           "Baisse sans volume": "🟠", "Absorption": "🔵"}
    st.markdown(" &nbsp; ".join(f"{emo.get(k, '⚪')} **{k}** : {v}" for k, v in counts.items()))
    filtre = st.multiselect("Filtrer par signal", sorted(sig["Signal"].unique()))
    s = sig[sig["Signal"].isin(filtre)] if filtre else sig
    s = s.assign(Signal=s["Signal"].map(lambda x: f"{emo.get(x, '⚪')} {x}"))
    st.dataframe(
        s[["Valeur", "Secteur", "Volume_periode_MAD", "Part_periode_pct", "Var_periode_pct", "Tendance_volume",
           "Pic_vs_mediane", "Ratio_dernier_jour_vs_moy20", "Ticket_moyen", "Flux", "Signal"]],
        hide_index=True, use_container_width=True, height=520,
        column_config={
            "Volume_periode_MAD": st.column_config.NumberColumn("Volume (MAD)", format="%,.0f"),
            "Part_periode_pct": st.column_config.ProgressColumn("Part", format="%.1f %%", min_value=0,
                                                                 max_value=float(max(1, sig["Part_periode_pct"].max()))),
            "Var_periode_pct": st.column_config.NumberColumn("Var. période", format="%+.1f %%"),
            "Tendance_volume": st.column_config.NumberColumn("Tendance vol.", format="%.2f×"),
            "Pic_vs_mediane": st.column_config.NumberColumn("Pic", format="%.1f×"),
            "Ratio_dernier_jour_vs_moy20": st.column_config.NumberColumn("Dernier jour / moy. 20", format="%.1f×"),
            "Ticket_moyen": st.column_config.NumberColumn("Ticket moyen", format="%,.0f"),
        })
    with st.expander("Comment lire les signaux"):
        st.markdown(
            "- 🟢 **Hausse confirmée** : le cours monte (≥ +2 %) et le volume augmente (≥ +20 %).\n"
            "- 🟡 **Hausse sans volume** : le cours monte mais l'activité baisse, la hausse est fragile.\n"
            "- 🔴 **Baisse sous volume** : le cours baisse et le volume grossit, pression vendeuse réelle.\n"
            "- 🔵 **Absorption** : gros volume (≥ 20 M MAD) et prix quasi immobile (< 1 %).\n"
            "- **Tendance vol.** : moyenne de la 2ᵉ moitié de la période ÷ moyenne de la 1ʳᵉ moitié.\n"
            "- **Pic** : plus forte séance ÷ séance médiane. **Dernier jour / moy. 20** : fiable à partir de ~20 séances.\n"
            "- **Flux** : ticket moyen ≥ 250 000 MAD = institutionnel, ≤ 60 000 MAD = particuliers.")

# ====================================================================== Liquidité
with tabs[3]:
    n_liq = min(nb_liq, len(dates))
    hc1, hc2 = st.columns([3, 1])
    hc1.subheader(f"Score de liquidité · {n_liq} dernières séances")
    with hc2.popover("ⓘ La formule, valeur par valeur", use_container_width=True):
        v_ = st.selectbox("Valeur", liq["Valeur"].tolist(), key="liq_formule")
        x_ = liq[liq["Valeur"] == v_].iloc[0]
        st.markdown(f"#### {v_} : {fr(x_.Score, 0)} / 100 · {x_.Classe}")
        st.markdown("On classe toutes les valeurs sur 4 ingrédients. Être le meilleur sur un ingrédient donne "
                    "tous ses points, être le dernier n'en donne presque aucun. **Score = somme des points.**")
        ing_ = pd.DataFrame({
            "Ingrédient": ["Argent échangé / jour (hors grosses transactions)", "Régularité",
                           "Écart acheteurs / vendeurs", "Coût pour faire bouger le cours"],
            "Mesure": [mm(x_.Volume_fond_median) + " MAD", f"{fr(x_.Frequence_pct, 0)} % des séances",
                       "—" if np.isnan(x_.Ecart_median_pct) else f"{fr(x_.Ecart_median_pct, 2)} %",
                       "—" if np.isnan(x_.MAD_pour_1pct_impact) else mm(x_.MAD_pour_1pct_impact) + " MAD pour 1 %"],
            "Points": [x_.Pts_volume, x_.Pts_frequence, x_.Pts_ecart, x_.Pts_impact],
            "Maximum": [30, 20, 20, 30]})
        ing_["Lbl"] = ing_.apply(lambda r: f"{fr(r.Points, 0)} / {r.Maximum}", axis=1)
        b_ = alt.Chart(ing_).encode(y=alt.Y("Ingrédient:N", sort=None, title=None, axis=alt.Axis(labelLimit=280)))
        st.altair_chart(alt.layer(
            b_.mark_bar(color="#E3E8F0", cornerRadius=4).encode(x=alt.X("Maximum:Q", axis=None)),
            b_.mark_bar(color=GOLD, cornerRadius=4).encode(x="Points:Q"),
            b_.mark_text(align="left", dx=4, color=INK).encode(x="Maximum:Q", text="Lbl:N"),
        ).properties(height=170), use_container_width=True)
        st.dataframe(ing_[["Ingrédient", "Mesure", "Points"]], hide_index=True, use_container_width=True,
                     column_config={"Points": st.column_config.NumberColumn(format="%.1f")})
        st.caption("Régularité : part des séances où la valeur a été échangée. Écart : (meilleure offre − meilleure "
                   "demande) ÷ milieu, à la clôture. Coût pour faire bouger le cours : argent qu'il faut échanger pour "
                   "que le cours varie de 1 % (médiane sur la période). Grosses transactions : au moins 5 M MAD et "
                   "1 M MAD par contrat.")
    if n_liq < 20:
        st.warning(f"Seulement {n_liq} séance(s) d'historique : le score se stabilise à partir de 20 séances.")
    classes = liq["Classe"].value_counts()
    st.markdown('<div class="kpis">' +
                kpi("Très liquides", str(classes.get("Très liquide", 0)), "score ≥ 75", gold=True) +
                kpi("Liquides", str(classes.get("Liquide", 0)), "score 50 à 75") +
                kpi("Peu liquides", str(classes.get("Peu liquide", 0)), "score 25 à 50") +
                kpi("Illiquides", str(classes.get("Illiquide", 0)), "score < 25") +
                kpi("Largeur du marché", f"{fr(lm.Largeur_pct, 0)} %", "des valeurs traitées, dernière séance") +
                '</div>', unsafe_allow_html=True)
    a, b = st.columns([3, 2])
    with a:
        st.dataframe(
            liq[["Rang", "Valeur", "Secteur", "Score", "Classe", "Volume_fond_median", "Frequence_pct",
                 "Ecart_median_pct", "Amihud_median", "MAD_pour_1pct_impact", "Part_blocs_pct"]],
            hide_index=True, use_container_width=True, height=560,
            column_config={
                "Score": st.column_config.ProgressColumn("Score", format="%.0f", min_value=0, max_value=100),
                "Volume_fond_median": st.column_config.NumberColumn("Volume de fond / jour", format="%,.0f"),
                "Frequence_pct": st.column_config.NumberColumn("Fréquence", format="%.0f %%"),
                "Ecart_median_pct": st.column_config.NumberColumn("Écart achat/vente", format="%.2f %%"),
                "Amihud_median": st.column_config.NumberColumn("Impact (% / M MAD)", format="%.3f"),
                "MAD_pour_1pct_impact": st.column_config.NumberColumn("MAD pour bouger de 1 %", format="%,.0f"),
                "Part_blocs_pct": st.column_config.NumberColumn("Part des blocs", format="%.0f %%"),
            })
    with b:
        sc = liq.dropna(subset=["Amihud_median"])
        sc = sc[(sc["Volume_fond_median"] > 0) & (sc["Amihud_median"] > 0)]
        st.altair_chart(alt.Chart(sc).mark_circle(size=90, opacity=.85, stroke="#fff", strokeWidth=1).encode(
            x=alt.X("Volume_fond_median:Q", scale=alt.Scale(type="log"), title="Volume de fond / jour (MAD, log)"),
            y=alt.Y("Amihud_median:Q", scale=alt.Scale(type="log", reverse=True), title="Impact prix (log, haut = faible)"),
            color=alt.Color("Classe:N", scale=alt.Scale(domain=["Très liquide", "Liquide", "Peu liquide", "Illiquide"],
                                                        range=[GREEN, "#3E7CB1", GOLD, RED]),
                            legend=alt.Legend(orient="bottom", title=None)),
            tooltip=["Valeur", alt.Tooltip("Score", format=".0f"), alt.Tooltip("Volume_fond_median", title="Volume de fond", format=",.0f"),
                     alt.Tooltip("Amihud_median", title="Impact", format=".3f"),
                     alt.Tooltip("MAD_pour_1pct_impact", title="MAD pour 1 %", format=",.0f")]
        ).properties(height=430, title="Volume et impact"), use_container_width=True)
        st.caption("En haut à droite : beaucoup de volume et peu d'impact, c'est-à-dire facile à traiter. "
                   "En bas à gauche : un petit montant suffit à faire bouger le cours.")
    st.subheader("Liquidité du marché séance par séance")
    lmm = liq_m.assign(Seance=liq_m["Date"].dt.strftime("%d/%m"))
    base = alt.Chart(lmm).encode(x=alt.X("Seance:N", sort=None, title=None))
    st.altair_chart(alt.layer(
        base.mark_bar(color=INK, opacity=.85, cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(
            y=alt.Y("Volume_fond:Q", title="Volume de fond (MAD)", axis=alt.Axis(format="~s"))),
        base.mark_line(color=GOLD, strokeWidth=3, point=True).encode(
            y=alt.Y("Largeur_pct:Q", title="Valeurs traitées (%)", scale=alt.Scale(domain=[0, 100])),
            tooltip=["Seance", alt.Tooltip("Volume_fond", title="Volume de fond", format=",.0f"),
                     alt.Tooltip("Valeurs_traitees", title="Valeurs traitées"),
                     alt.Tooltip("Ecart_median_pct", title="Écart médian %", format=".2f")])
    ).resolve_scale(y="independent").properties(height=280), use_container_width=True)

# ====================================================================== Zoom valeur
with tabs[4]:
    choix = st.selectbox("Valeur", sig["Valeur"].tolist())
    r = sig[sig["Valeur"] == choix].iloc[0]
    lq = liq[liq["Valeur"] == choix]
    st.markdown('<div class="kpis">' +
                kpi("Volume période", mm(r["Volume_periode_MAD"]), "MAD", gold=True) +
                kpi("Variation", pct(r["Var_periode_pct"], 1), f"sur {nb} séances",
                    "up" if r["Var_periode_pct"] >= 0 else "dn") +
                kpi("Ticket moyen", fr(r["Ticket_moyen"] / 1e3, 0) + " k", r["Flux"]) +
                kpi("Signal", r["Signal"], "", txt=True) +
                (kpi("Liquidité", fr(lq["Score"].iloc[0], 0) + " / 100", lq["Classe"].iloc[0]) if len(lq) else "") +
                '</div>', unsafe_allow_html=True)
    a, b = st.columns(2)
    with a:
        st.subheader("Prix et volume")
        d = vj[(vj["Valeur"] == choix) & (vj["Date"].isin(periode))].copy()
        d["Seance"] = d["Date"].dt.strftime("%d/%m")
        d["M"] = d["Volume_MAD"] / 1e6
        base = alt.Chart(d).encode(x=alt.X("Seance:N", title=None, sort=None))
        vol = base.mark_bar(color="#C9D3E0", cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
            y=alt.Y("M:Q", title="Volume (M MAD)"))
        pmin, pmax = d[["Cloture", "VWAP"]].min().min(), d[["Cloture", "VWAP"]].max().max()
        scl = alt.Scale(domain=[pmin * 0.99, pmax * 1.01])
        close = base.mark_line(color=GOLD, strokeWidth=3, point=alt.OverlayMarkDef(color=GOLD, size=60)).encode(
            y=alt.Y("Cloture:Q", title="Cours", scale=scl),
            tooltip=[alt.Tooltip("Seance", title="Séance"), alt.Tooltip("Cloture", title="Clôture", format=",.2f"),
                     alt.Tooltip("VWAP", format=",.2f"), alt.Tooltip("M", title="Volume (M MAD)", format=",.1f"),
                     alt.Tooltip("Contrats", format=",")])
        vwap = base.mark_tick(color=INK, thickness=2, size=30).encode(y=alt.Y("VWAP:Q", scale=scl))
        st.altair_chart(alt.layer(vol, alt.layer(close, vwap)).resolve_scale(y="independent"), use_container_width=True)
        st.caption("Barres : volume · ligne dorée : clôture · trait bleu : VWAP du jour.")
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
        prof["M"] = prof["Volume_MAD"] / 1e6
        prof["Niveau"] = np.where(prof["M"] == prof["M"].max(), "Le plus chargé", "Autres")
        st.altair_chart(alt.Chart(prof).mark_bar(cornerRadiusTopRight=3, cornerRadiusBottomRight=3).encode(
            y=alt.Y("Prix:N", sort="descending", title=None), x=alt.X("M:Q", title="M MAD"),
            color=alt.Color("Niveau:N", scale=alt.Scale(domain=["Le plus chargé", "Autres"], range=[GOLD, INK]), legend=None),
            tooltip=["Prix", alt.Tooltip("M", title="M MAD", format=",.2f")]), use_container_width=True)
        st.caption("Le niveau doré concentre le plus de volume : support ou résistance probable.")

# ====================================================================== Blocs
with tabs[5]:
    st.subheader("Grosses transactions (≥ 5 M MAD et ≥ 1 M MAD par contrat)")
    bp = blocs[blocs["Date"].isin(periode)]
    st.dataframe(bp, hide_index=True, use_container_width=True,
                 column_config={"Date": st.column_config.DateColumn(format="DD/MM/YYYY"),
                                "Volume_MAD": st.column_config.NumberColumn(format="%,.0f"),
                                "MAD_par_contrat": st.column_config.NumberColumn(format="%,.0f"),
                                "Part_du_jour_pct": st.column_config.ProgressColumn("Part du jour", format="%.0f %%",
                                                                                     min_value=0, max_value=100),
                                "Prix_vs_cloture_pct": st.column_config.NumberColumn(format="%+.2f %%")})
    st.caption("Un bloc payé au-dessus de la clôture montre un acheteur décidé ; en dessous, un vendeur pressé.")

st.markdown('<div class="foot">Outil indépendant, non affilié à la Bourse de Casablanca. '
            'Données issues des bulletins de la cote publics. Ne constitue pas un conseil en investissement.</div>',
            unsafe_allow_html=True)
