# dashboard.py  
import streamlit as st  
import numpy as np  
import pandas as pd  
import plotly.graph_objects as go  
import plotly.express as px  
from plotly.subplots import make_subplots  

# ── Import de votre module existant ──────────────────────  
import optimizer as opt  # Votre fichier renommé optimizer.py  

# ══════════════════════════════════════════════════════════  
# CONFIG PAGE  
# ══════════════════════════════════════════════════════════  
st.set_page_config(  
    page_title="Portfolio Optimizer",  
    page_icon="📊",  
    layout="wide",  
    initial_sidebar_state="expanded"  
)  

# ── CSS custom minimal ────────────────────────────────────  
st.markdown("""  
<style>  
    .main { background-color: #0e1117; }  
    .metric-card {  
        background: #1e2130;  
        border-radius: 10px;  
        padding: 16px;  
        border-left: 4px solid #4f8ef7;  
    }  
    .stSidebar { background-color: #161b2e; }  
    h1, h2, h3 { color: #e8eaf6; }  
    .section-title {  
        color: #7986cb;  
        font-size: 0.75rem;  
        font-weight: 700;  
        letter-spacing: 0.1em;  
        text-transform: uppercase;  
        margin-top: 1.2rem;  
        margin-bottom: 0.3rem;  
    }  
</style>  
""", unsafe_allow_html=True)  

# ══════════════════════════════════════════════════════════  
# UNIVERSE — Chargé depuis universe.py (source unique)  
# ══════════════════════════════════════════════════════════  
from universe import UNIVERSE as _BASE_UNIVERSE 
# Mapping continent (dashboard en a besoin pour les filtres UI)  
_CONTINENT_MAP = {  
    "US":     "Amérique",  
    "EU":     "Europe",  
    "EM":     "Asie / EM",  
    "GLOBAL": "Global",  
    "OTHER":  "Global",  
}  

# Construction de FULL_UNIVERSE avec continent garanti  
FULL_UNIVERSE = []  
for _a in _BASE_UNIVERSE:  
    _entry = dict(_a)  # copie propre, ne pas muter universe.py  
    # continent : déjà présent dans universe.py ou on le mappe  
    if "continent" not in _entry:  
        _entry["continent"] = _CONTINENT_MAP.get(_entry.get("region", "OTHER"), "Global")  
    # sector : déjà présent dans universe.py  
    if "sector" not in _entry:  
        _entry["sector"] = "Autre"  
    FULL_UNIVERSE.append(_entry)   

CONTINENTS    = sorted(set(a["continent"] for a in FULL_UNIVERSE))  
SECTORS       = sorted(set(a["sector"]    for a in FULL_UNIVERSE))  
KINDS         = sorted(set(a["kind"]      for a in FULL_UNIVERSE))  
ALL_REGIONS   = sorted(set(a["region"]    for a in FULL_UNIVERSE))  

# Map continent → pays/régions disponibles  
CONTINENT_TO_REGIONS = {}  
for a in FULL_UNIVERSE:  
    c = a["continent"]  
    r = a["region"]  
    CONTINENT_TO_REGIONS.setdefault(c, set()).add(r)  
CONTINENT_TO_REGIONS = {k: sorted(v) for k, v in CONTINENT_TO_REGIONS.items()}  

# ══════════════════════════════════════════════════════════  
# SIDEBAR — PANNEAU DE CONTRÔLE  
# ══════════════════════════════════════════════════════════  
with st.sidebar:  
    st.markdown("## ⚙️ Panneau de Contrôle")  
    st.markdown("---")  

    # ── 1. ZONE GÉOGRAPHIQUE ──────────────────────────────  
    st.markdown('<p class="section-title">🌍 Zone Géographique</p>', unsafe_allow_html=True)  

    selected_continents = st.multiselect(  
        "Continents",  
        options=CONTINENTS,  
        default=["Amérique", "Europe"],  
        help="Sélectionnez les continents à inclure dans l'univers"  
    )  

    available_regions = []  
    for c in selected_continents:  
        available_regions += CONTINENT_TO_REGIONS.get(c, [])  
    available_regions = sorted(set(available_regions))  

    selected_regions = st.multiselect(  
        "Régions / Pays",  
        options=available_regions,  
        default=available_regions,  
        help="Filtrage fin par région"  
    )  

    # ── 2. SECTEURS ───────────────────────────────────────  
    st.markdown('<p class="section-title">🏭 Secteurs d\'activité</p>', unsafe_allow_html=True)  

    selected_sectors = st.multiselect(  
        "Secteurs",  
        options=SECTORS,  
        default=SECTORS,  
        help="Inclure/exclure des secteurs entiers"  
    )  

    # ── 3. TYPE D'ACTIFS ──────────────────────────────────  
    st.markdown('<p class="section-title">📦 Types d\'actifs</p>', unsafe_allow_html=True)  

    selected_kinds = st.multiselect(  
        "Catégories",  
        options=KINDS,  
        default=KINDS,  
        help="ETF, STOCK, GOLD, CASH..."  
    )  

# ── 4. PARAMÈTRES DE RISQUE ───────────────────────────
    st.markdown('<p class="section-title">📊 Risque & Volatilité</p>', unsafe_allow_html=True)

    RISK_PRESETS = {
        "Très Défensif": {"vol_target": 0.06, "cash_floor_risk_on": 0.15, "cash_floor_risk_off": 0.35},
        "Défensif":      {"vol_target": 0.08, "cash_floor_risk_on": 0.05, "cash_floor_risk_off": 0.25},
        "Équilibré":     {"vol_target": 0.12, "cash_floor_risk_on": 0.00, "cash_floor_risk_off": 0.20},
        "Dynamique":     {"vol_target": 0.16, "cash_floor_risk_on": 0.00, "cash_floor_risk_off": 0.15},
        "Agressif":      {"vol_target": 0.22, "cash_floor_risk_on": 0.00, "cash_floor_risk_off": 0.10},
    }

    def apply_risk_profile():
        profile = st.session_state["risk_profile"]
        preset = RISK_PRESETS[profile]
        st.session_state["vol_target_pct"] = int(round(preset["vol_target"] * 100))
        st.session_state["cash_floor_risk_on_pct"] = int(round(preset["cash_floor_risk_on"] * 100))
        st.session_state["cash_floor_risk_off_pct"] = int(round(preset["cash_floor_risk_off"] * 100))

    if "risk_profile" not in st.session_state:
        st.session_state["risk_profile"] = "Équilibré"
    if "vol_target_pct" not in st.session_state:
        st.session_state["vol_target_pct"] = 12
    if "cash_floor_risk_on_pct" not in st.session_state:
        st.session_state["cash_floor_risk_on_pct"] = 0
    if "cash_floor_risk_off_pct" not in st.session_state:
        st.session_state["cash_floor_risk_off_pct"] = 20

    risk_profile = st.select_slider(
        "Profil de risque",
        options=["Très Défensif", "Défensif", "Équilibré", "Dynamique", "Agressif"],
        key="risk_profile",
        on_change=apply_risk_profile,
        help="Charge un preset de volatilité cible et de niveaux de cash"
    )

    vol_target_pct = st.slider(
        "Volatilité cible (annuelle)",
        min_value=4,
        max_value=30,
        step=1,
        key="vol_target_pct",
        format="%d%%",
        help="Plafond de volatilité annualisée du portefeuille"
    )

    cash_floor_risk_on_pct = st.slider(
        "Cash minimum en régime Risk-On",
        min_value=0,
        max_value=40,
        step=1,
        key="cash_floor_risk_on_pct",
        format="%d%%"
    )

    cash_floor_risk_off_pct = st.slider(
        "Cash minimum en régime Risk-Off",
        min_value=0,
        max_value=60,
        step=1,
        key="cash_floor_risk_off_pct",
        format="%d%%"
    )

    vol_target = vol_target_pct / 100.0
    cash_floor_risk_on = cash_floor_risk_on_pct / 100.0
    cash_floor_risk_off = cash_floor_risk_off_pct / 100.0

    st.caption(
        f"Actif : vol cible {vol_target:.0%} | cash risk-on {cash_floor_risk_on:.0%} | "
        f"cash risk-off {cash_floor_risk_off:.0%}"
    )
    # ── 5. CORRÉLATION MARCHÉ ─────────────────────────────  
    st.markdown('<p class="section-title">🔗 Corrélation Marché</p>', unsafe_allow_html=True)  

    market_corr_mode = st.radio(  
        "Exposition marché",  
        options=["Libre", "Beta contrôlé", "Market Neutral"],  
        index=0,  
        horizontal=True,  
        help="Contrôle le beta vs benchmark"  
    )  

    # ── 6. BORNES RÉGIONALES ──────────────────────────────  
    st.markdown('<p class="section-title">⚖️ Contraintes Régionales</p>', unsafe_allow_html=True)  

    region_bounds = {}  
    active_regions_for_bounds = [r for r in ["US", "EU", "EM"] if r in (selected_regions or ALL_REGIONS)]  

    for reg in active_regions_for_bounds:  
        defaults = {"US": (0.25, 0.55), "EU": (0.20, 0.65), "EM": (0.00, 0.20)}  
        lo_def, hi_def = defaults.get(reg, (0.0, 1.0))  
        col1, col2 = st.columns(2)  
        with col1:  
            lo = st.number_input(f"{reg} Min", 0.0, 1.0, lo_def, 0.05, key=f"lo_{reg}", format="%.2f")  
        with col2:  
            hi = st.number_input(f"{reg} Max", 0.0, 1.0, hi_def, 0.05, key=f"hi_{reg}", format="%.2f")  
        if lo <= hi:  
            region_bounds[reg] = (lo, hi)  

    # ── 7. CAPS PAR TYPE ─────────────────────────────────  
    st.markdown('<p class="section-title">🎚️ Caps par Type d\'actif</p>', unsafe_allow_html=True)  

    cap_etf_pct   = st.slider("Cap ETF",   10, 100, 40, 5, format="%d%%")
    cap_stock_pct = st.slider("Cap STOCK",  5,  30, 10, 1, format="%d%%")
    cap_gold_pct  = st.slider("Cap GOLD",   0,  30, 15, 1, format="%d%%")

    cap_etf   = cap_etf_pct / 100.0
    cap_stock = cap_stock_pct / 100.0
    cap_gold  = cap_gold_pct / 100.0

    caps_by_kind = {  
        "ETF": cap_etf,  
        "STOCK": cap_stock,  
        "GOLD": cap_gold,  
        "CASH": 1.00,  
        "OTHER": 0.10  
    }  

    # ── 8. PARAMÈTRES AVANCÉS ────────────────────────────  
    with st.expander("🔬 Paramètres Avancés"):  
        window_years   = st.slider("Fenêtre estimation (années)", 2, 10, 5)  
        mean_shrink    = st.slider("Shrinkage mu", 0.0, 1.0, 0.5, 0.05)  
        max_turnover_pct = st.slider("Max Turnover (L1)", 2, 50, 10, 1, format="%d%%")
        min_rebal_l1_pct = st.slider("Min Rebal L1 (skip si <)", 0, 10, 3, 1, format="%d%%")
        max_turnover = max_turnover_pct / 100.0
        min_rebal_l1 = min_rebal_l1_pct / 100.0
        turn_gamma     = st.slider("Turnover Gamma (lissage)", 0.0, 50.0, 10.0, 1.0)  
        rebalance_freq = st.selectbox("Fréquence de rebalancement", ["ME", "QE", "YE"], index=0)  
        cost_bps       = st.slider("Coûts de transaction (bps)", 0, 50, 5, 1)  
        cov_method     = st.radio("Méthode Covariance", ["lw", "sample"], horizontal=True)  
        risk_off_enabled = st.toggle("Activer Risk-Off Signal", value=True)  

    # ── 9. LANCEMENT ──────────────────────────────────────  
    st.markdown("---")  
    run_button = st.button("🚀 Lancer l'optimisation", type="primary", use_container_width=True)  

# ══════════════════════════════════════════════════════════  
# FILTRAGE DE L'UNIVERS  
# ══════════════════════════════════════════════════════════  
def filter_universe(universe, continents, regions, sectors, kinds):  
    filtered = []  
    for a in universe:  
        if continents and a.get("continent") not in continents:  
            continue  
        if regions and a.get("region") not in (regions + ["OTHER"]):  
            continue  
        if sectors and a.get("sector") not in sectors:  
            continue  
        if kinds and a.get("kind") not in kinds:  
            continue  
        filtered.append(a)  
    # CASH toujours présent si CASH sélectionné  
    names = [a["name"] for a in filtered]  
    if "CASH" not in names:  
        cash = next((a for a in universe if a["name"] == "CASH"), None)  
        if cash:  
            filtered.append(cash)  
    return filtered  

# ══════════════════════════════════════════════════════════  
# MAIN PANEL  
# ══════════════════════════════════════════════════════════  
st.title("📊 Portfolio Optimizer — Dashboard")  
st.markdown("*Optimisation Mean-Variance avec contraintes régionales, sectorielles et rails de risque*")  

# ── Header KPIs (vides avant run) ────────────────────────  
kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)  

# ── Tabs ──────────────────────────────────────────────────  
tab_nav, tab_alloc, tab_perf, tab_weights, tab_universe = st.tabs([  
    "📈 NAV & Performance",  
    "🥧 Allocation",  
    "📊 Métriques",  
    "🗓️ Historique Poids",  
    "🔭 Univers Filtré"  
])  

# ══════════════════════════════════════════════════════════  
# ONGLET UNIVERS (toujours visible)  
# ══════════════════════════════════════════════════════════  
with tab_universe:  
    active_universe = filter_universe(  
        FULL_UNIVERSE, selected_continents, selected_regions, selected_sectors, selected_kinds  
    )  
    df_univ = pd.DataFrame(active_universe)[["name","ticker","kind","region","continent","sector","ccy"]]  
    
    st.markdown(f"### {len(active_universe)} actifs dans l'univers filtré")  

    # Stats de l'univers  
    c1, c2, c3 = st.columns(3)  
    with c1:  
        kind_counts = df_univ["kind"].value_counts()  
        fig_k = px.pie(values=kind_counts.values, names=kind_counts.index,  
                       title="Répartition par Type",  
                       color_discrete_sequence=px.colors.sequential.Blues_r,  
                       hole=0.4)  
        fig_k.update_layout(paper_bgcolor="#0e1117", font_color="#e8eaf6", height=280)  
        st.plotly_chart(fig_k, use_container_width=True)  
    with c2:  
        reg_counts = df_univ["region"].value_counts()  
        fig_r = px.pie(values=reg_counts.values, names=reg_counts.index,  
                       title="Répartition par Région",  
                       color_discrete_sequence=px.colors.sequential.Greens_r,  
                       hole=0.4)  
        fig_r.update_layout(paper_bgcolor="#0e1117", font_color="#e8eaf6", height=280)  
        st.plotly_chart(fig_r, use_container_width=True)  
    with c3:  
        sec_counts = df_univ["sector"].value_counts()  
        fig_s = px.pie(values=sec_counts.values, names=sec_counts.index,  
                       title="Répartition par Secteur",  
                       color_discrete_sequence=px.colors.sequential.Oranges_r,  
                       hole=0.4)  
        fig_s.update_layout(paper_bgcolor="#0e1117", font_color="#e8eaf6", height=280)  
        st.plotly_chart(fig_s, use_container_width=True)  

    st.dataframe(  
        df_univ.style.set_properties(**{"background-color": "#1e2130", "color": "#e8eaf6"}),  
        use_container_width=True, height=400  
    )  

# ══════════════════════════════════════════════════════════  
# LANCEMENT DE L'OPTIMISATION  
# ══════════════════════════════════════════════════════════  
if "opt_results" not in st.session_state:
    st.session_state["opt_results"] = None

if run_button:  
    active_universe = filter_universe(  
        FULL_UNIVERSE, selected_continents, selected_regions, selected_sectors, selected_kinds  
    )  

    # ── Chargement des données ────────────────────────────  
    with st.spinner("⏳ Téléchargement des données depuis Stooq..."):  
        try:  
            eur_px = opt.build_eur_price_matrix(active_universe)  
        except Exception as e:  
            st.error(f"❌ Erreur de chargement des données : {e}")  
            st.stop()  
    # ── Filtrage historique minimal ──────────────────────────
    min_start = eur_px.index.max() - pd.DateOffset(years=window_years + 1)

    eligible_cols = [
        c for c in eur_px.columns
        if eur_px[c].first_valid_index() is not None
        and eur_px[c].first_valid_index() <= min_start
    ]

    eur_px = eur_px[eligible_cols]

    if eur_px.empty or eur_px.shape[1] < 3:
        st.error("❌ Pas assez d'actifs avec un historique suffisant pour cette fenêtre.")
        st.stop()  
    # Re-sync universe with actual retained columns
    valid_names = set(eur_px.columns)
    active_universe = [a for a in active_universe if a["name"] in valid_names]

    if len(active_universe) < 3:
        st.error("❌ Univers trop restreint après filtrage historique. Sélectionnez plus d'actifs.")
        st.stop()

    # Rebuild benchmark on retained assets only
    bench_w = {
        a["name"]: 1.0
        for a in active_universe
        if a["kind"] == "ETF" and a["name"] in valid_names
    }

    if not bench_w:
        bench_w = {active_universe[0]["name"]: 1.0}

    # Rebuild constraints on retained universe
    max_weight_by_kind = {"STOCK": 0.40, "GOLD": float(cap_gold)}
    min_weight_by_kind = {"ETF": 0.30} if any(a["kind"] == "ETF" for a in active_universe) else {}


    st.success(f"✅ Données chargées : {eur_px.index.min().date()} → {eur_px.index.max().date()} | {eur_px.shape[1]} actifs")  

    # ── Optimisation ──────────────────────────────────────  
    with st.spinner("🧮 Optimisation en cours..."):  
        try:  
            nav, weights, lambdas = opt.backtest_allocator(
                eur_px, active_universe,
                window_years=window_years,
                rebalance=rebalance_freq,
                rf=0.0,
                mean_shrink=mean_shrink,
                cov_method=cov_method,
                winsor_mode="std",
                winsor_param=4.0,
                region_bounds=region_bounds if region_bounds else None,
                caps_by_kind=caps_by_kind,
                max_weight_by_kind=max_weight_by_kind,
                min_weight_by_kind=min_weight_by_kind,
                turn_gamma=turn_gamma,
                l2=1e-4,
                lambda_mode="grid",
                lam_grid=np.logspace(-1, 2, 25),
                tactical_em=False,
                cost_bps=float(cost_bps),
                require_full_window=True,
                max_turnover=float(max_turnover),
                min_rebalance_l1=float(min_rebal_l1),
                vol_target=float(vol_target),
                risk_off_enabled=risk_off_enabled,
                bench_weights_for_risk=bench_w,
                risk_off_ma_days=200,
                cash_floor_risk_off=float(cash_floor_risk_off),
                cash_floor_risk_on=float(cash_floor_risk_on),
            )
        except Exception as e:  
            st.error(f"❌ Erreur d'optimisation : {e}")  
            st.stop()  

    # ── Calculs de performance ────────────────────────────  
    port_rets  = nav.pct_change().dropna()
    asset_rets = opt.simple_returns_from_prices(eur_px).dropna(how="any")

    bench_w = {k: v for k, v in bench_w.items() if k in asset_rets.columns}
    if not bench_w:
        st.error("❌ Aucun composant de benchmark disponible dans les rendements calculés.")
        st.stop()

    bench_rets = opt.make_custom_benchmark(asset_rets, bench_w)  

    df_rb = pd.concat([port_rets.rename("PORT"), bench_rets.rename("BENCH")], axis=1).dropna()  
    p_ret, p_vol, p_sh = opt.ann_return_vol_sharpe(df_rb["PORT"])  
    b_ret, b_vol, b_sh = opt.ann_return_vol_sharpe(df_rb["BENCH"])  
    te, corr           = opt.tracking_error(df_rb["PORT"], df_rb["BENCH"])  
    b_full             = opt.beta(df_rb["PORT"], df_rb["BENCH"])  
    max_dd             = float(((nav / nav.cummax()) - 1).min())  
    to_stats           = opt.turnover(weights)  
    
    st.session_state["opt_results"] = {
    "nav": nav,
    "weights": weights,
    "lambdas": lambdas,
    "eur_px": eur_px,
    "bench_rets": bench_rets,
    "port_rets": port_rets,
    "p_ret": p_ret,
    "p_vol": p_vol,
    "p_sh": p_sh,
    "b_ret": b_ret,
    "b_vol": b_vol,
    "b_sh": b_sh,
    "te": te,
    "corr": corr,
    "b_full": b_full,
    "max_dd": max_dd,
    "to_stats": to_stats,
    "vol_target": vol_target,
}
results = st.session_state.get("opt_results")

if results is not None:
    nav = results["nav"]
    weights = results["weights"]
    lambdas = results["lambdas"]
    eur_px = results["eur_px"]
    bench_rets = results["bench_rets"]
    port_rets = results["port_rets"]
    p_ret = results["p_ret"]
    p_vol = results["p_vol"]
    p_sh = results["p_sh"]
    b_ret = results["b_ret"]
    b_vol = results["b_vol"]
    b_sh = results["b_sh"]
    te = results["te"]
    corr = results["corr"]
    b_full = results["b_full"]
    max_dd = results["max_dd"]
    to_stats = results["to_stats"]
    vol_target = results["vol_target"]

    # ── KPIs ─────────────────────────────────────────────
    with kpi1:
        st.metric("📈 Rendement Annuel", f"{p_ret:.2%}", f"{p_ret - b_ret:+.2%} vs bench")
    with kpi2:
        st.metric("📉 Volatilité", f"{p_vol:.2%}", f"cible ≤ {vol_target:.0%}")
    with kpi3:
        st.metric("⚡ Sharpe", f"{p_sh:.2f}", f"bench: {b_sh:.2f}")
    with kpi4:
        st.metric("💧 Max Drawdown", f"{max_dd:.2%}")
    with kpi5:
        st.metric("🔄 Turnover Moyen", f"{to_stats.mean():.2%}")

    # ══════════════════════════════════════════════════════
    # TAB 1 — NAV
    # ══════════════════════════════════════════════════════
    with tab_nav:
        bench_nav = (1 + bench_rets).cumprod()
        bench_nav = bench_nav / bench_nav.iloc[0]
        port_nav = nav / nav.iloc[0]

        fig_nav = go.Figure()
        fig_nav.add_trace(go.Scatter(
            x=port_nav.index, y=port_nav.values,
            name="Portfolio", line=dict(color="#4f8ef7", width=2.5)
        ))
        fig_nav.add_trace(go.Scatter(
            x=bench_nav.index, y=bench_nav.values,
            name="Benchmark", line=dict(color="#f7c34f", width=1.5, dash="dot")
        ))
        fig_nav.update_layout(
            title="NAV — Walk-forward (base 1)",
            paper_bgcolor="#0e1117", plot_bgcolor="#161b2e",
            font=dict(color="#e8eaf6"),
            xaxis=dict(gridcolor="#2a2e3f"),
            yaxis=dict(gridcolor="#2a2e3f"),
            legend=dict(bgcolor="#1e2130"),
            height=420
        )
        st.plotly_chart(fig_nav, use_container_width=True)

        dd_series = (nav / nav.cummax()) - 1
        fig_dd = go.Figure()
        fig_dd.add_trace(go.Scatter(
            x=dd_series.index, y=dd_series.values,
            fill="tozeroy", name="Drawdown",
            line=dict(color="#ef5350"), fillcolor="rgba(239,83,80,0.25)"
        ))
        fig_dd.update_layout(
            title="Drawdown",
            paper_bgcolor="#0e1117", plot_bgcolor="#161b2e",
            font=dict(color="#e8eaf6"),
            xaxis=dict(gridcolor="#2a2e3f"),
            yaxis=dict(gridcolor="#2a2e3f", tickformat=".1%"),
            height=220
        )
        st.plotly_chart(fig_dd, use_container_width=True)

    # ══════════════════════════════════════════════════════
    # TAB 2 — ALLOCATION
    # ══════════════════════════════════════════════════════
    with tab_alloc:
        last_w = weights.iloc[-1].dropna()
        last_w = last_w[last_w > 0.001]

        col_pie, col_bar = st.columns([1, 1])

        with col_pie:
            fig_pie = go.Figure(go.Pie(
                labels=last_w.index.tolist(),
                values=last_w.values.tolist(),
                hole=0.45,
                textinfo="label+percent",
                marker=dict(colors=px.colors.qualitative.Plotly)
            ))
            fig_pie.update_layout(
                title="Allocation actuelle",
                paper_bgcolor="#0e1117",
                font=dict(color="#e8eaf6"),
                height=450
            )
            st.plotly_chart(fig_pie, use_container_width=True)

        with col_bar:
            alloc_df = last_w.sort_values(ascending=True)
            fig_bar = go.Figure(go.Bar(
                x=alloc_df.values,
                y=alloc_df.index,
                orientation="h"
            ))
            fig_bar.update_layout(
                title="Poids par actif",
                paper_bgcolor="#0e1117",
                plot_bgcolor="#161b2e",
                font=dict(color="#e8eaf6"),
                xaxis=dict(gridcolor="#2a2e3f", tickformat=".0%"),
                yaxis=dict(gridcolor="#2a2e3f"),
                height=450
            )
            st.plotly_chart(fig_bar, use_container_width=True)

    # ══════════════════════════════════════════════════════
    # TAB 3 — MÉTRIQUES
    # ══════════════════════════════════════════════════════
    with tab_perf:
        metrics_df = pd.DataFrame({
            "Metric": ["Annual Return", "Volatility", "Sharpe", "Tracking Error", "Correlation", "Beta", "Max Drawdown"],
            "Portfolio": [p_ret, p_vol, p_sh, te, corr, b_full, max_dd]
        })

        st.dataframe(metrics_df, use_container_width=True)

    # ══════════════════════════════════════════════════════
    # TAB 4 — HISTORIQUE POIDS
    # ══════════════════════════════════════════════════════
    with tab_weights:
        if not weights.empty:
            weights_plot = weights.fillna(0.0)
            fig_w = go.Figure()
            for col in weights_plot.columns:
                if weights_plot[col].abs().max() > 0:
                    fig_w.add_trace(go.Scatter(
                        x=weights_plot.index,
                        y=weights_plot[col],
                        mode="lines",
                        name=col,
                        stackgroup="one"
                    ))
            fig_w.update_layout(
                title="Historique des poids",
                paper_bgcolor="#0e1117",
                plot_bgcolor="#161b2e",
                font=dict(color="#e8eaf6"),
                xaxis=dict(gridcolor="#2a2e3f"),
                yaxis=dict(gridcolor="#2a2e3f", tickformat=".0%"),
                height=500
            )
            st.plotly_chart(fig_w, use_container_width=True)
        else:
            st.info("Aucun historique de poids disponible.")
else:
    with tab_nav:
        st.info("Lancez une optimisation pour afficher les résultats.")
    with tab_alloc:
        st.info("Lancez une optimisation pour afficher les résultats.")
    with tab_perf:
        st.info("Lancez une optimisation pour afficher les résultats.")
    with tab_weights:
        st.info("Lancez une optimisation pour afficher les résultats.")