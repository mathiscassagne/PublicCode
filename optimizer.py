import numpy as np
import pandas as pd
import cvxpy as cp
from sklearn.covariance import LedoitWolf
import matplotlib.pyplot as plt
from universe import UNIVERSE, get_universe  
from data_loader import build_price_matrix  

def indices_by_filter(universe, columns, key, value):
    """
    Returns indices in 'columns' for assets in universe with a.get(key)==value.
    """
    name_to_idx = {name: i for i, name in enumerate(columns)}
    idx = []
    for a in universe:
        if a.get(key) == value and a.get("name") in name_to_idx:
            idx.append(name_to_idx[a["name"]])
    return idx


def build_cap_vector(universe, columns, caps_by_kind, default_cap=0.10):
    """
    Per-asset cap vector based on asset 'kind' (ETF/STOCK/GOLD/CASH/etc).
    """
    name_to_kind = {a.get("name"): a.get("kind", "OTHER") for a in universe}
    return np.array(
        [caps_by_kind.get(name_to_kind.get(col, "OTHER"), default_cap) for col in columns],
        dtype=float
    )

def prepare_window(px_slice, min_coverage=0.95, max_ffill=3):
    coverage = px_slice.notna().mean()
    keep_cols = coverage[coverage >= min_coverage].index.tolist()

    if not keep_cols:
        return pd.DataFrame()

    out = px_slice[keep_cols].copy()
    out = out.ffill(limit=max_ffill)
    out = out.dropna(how="any")
    return out

def feasibility_check(universe, columns, cap_vec,
                      region_bounds=None,
                      max_weight_by_kind=None,
                      min_weight_by_kind=None):
    """
    Basic feasibility checks to catch impossible constraints early.
    This does NOT guarantee feasibility, but it catches obvious contradictions.
    """
    # 1) Sum of max per-asset caps must allow reaching 100%
    max_total = float(np.sum(cap_vec))

    # 2) If you have sleeve caps (max_weight_by_kind), adjust theoretical max_total
    if max_weight_by_kind:
        for kind, cap in max_weight_by_kind.items():
            idx = indices_by_filter(universe, columns, "kind", kind)
            if idx:
                # These assets collectively cannot exceed 'cap'
                # If their per-asset caps sum to more than 'cap', subtract the excess from max_total
                excess = float(np.sum(cap_vec[idx]) - cap)
                if excess > 0:
                    max_total -= excess

    if max_total < 1 - 1e-9:
        raise RuntimeError(
            f"Infeasible: maximum reachable total weight is {max_total:.3f} < 1. "
            f"Loosen per-asset caps and/or sleeve caps."
        )

    # 3) Kind floors must not exceed 100%
    if min_weight_by_kind:
        floors_sum = float(sum(min_weight_by_kind.values()))
        if floors_sum > 1 + 1e-9:
            raise RuntimeError(f"Infeasible: kind floors sum to {floors_sum:.3f} > 1.")

    # 4) Region bounds should allow sum to reach 1
    if region_bounds:
        sum_lo = float(sum(lo for lo, hi in region_bounds.values()))
        sum_hi = float(sum(hi for lo, hi in region_bounds.values()))
        if sum_lo > 1 + 1e-9:
            raise RuntimeError(f"Infeasible: region lower bounds sum to {sum_lo:.3f} > 1.")
        if sum_hi < 1 - 1e-9:
            raise RuntimeError(f"Infeasible: region upper bounds sum to {sum_hi:.3f} < 1.")
        
def add_turnover_constraint(cons, w, w_prev, max_turnover):
    """
    Adds L1 turnover constraint:
        sum_i |w_i - w_prev_i| <= max_turnover

    Implemented with epigraph variables so it stays convex.
    """
    if w_prev is None or max_turnover is None:
        return cons

    w_prev = np.asarray(w_prev).ravel()
    n = int(w.shape[0])

    if len(w_prev) != n:
        raise ValueError(f"w_prev length {len(w_prev)} != n {n}")

    t = cp.Variable(n)  # |w - w_prev| slack
    cons += [
        t >= w - w_prev,
        t >= -(w - w_prev),
        cp.sum(t) <= float(max_turnover)
    ]
    return cons


def add_volatility_cap(cons, w, Sigma, vol_target):
    if vol_target is None:
        return cons

    vol_target = float(vol_target)
    if vol_target <= 0:
        raise ValueError("vol_target must be > 0")

    Sigma_sym = 0.5 * (Sigma + Sigma.T)

    # Build A such that A.T @ A = Sigma
    try:
        L = np.linalg.cholesky(Sigma_sym)  # Sigma = L L.T
        A = L.T                             # A.T A = L L.T = Sigma
    except np.linalg.LinAlgError:
        vals, vecs = np.linalg.eigh(Sigma_sym)
        vals = np.maximum(vals, 1e-12)
        A = (vecs * np.sqrt(vals)) @ vecs.T  # symmetric sqrt => A.T A = Sigma

    cons += [cp.norm(A @ w, 2) <= vol_target]
    return cons


def add_min_cash_floor(cons, w, columns, cash_floor):
    """
    Enforces w['CASH'] >= cash_floor if CASH exists in columns.
    """
    if cash_floor is None:
        return cons

    cash_floor = float(cash_floor)
    if cash_floor <= 0:
        return cons

    if "CASH" not in columns:
        raise ValueError("cash_floor was requested but 'CASH' is not in columns/universe.")

    cash_idx = columns.index("CASH")
    cons += [w[cash_idx] >= cash_floor]
    return cons


def add_max_cash_cap(cons, w, columns, cash_cap):
    """
    Enforces w['CASH'] <= cash_cap if CASH exists.
    """
    if cash_cap is None:
        return cons
    cash_cap = float(cash_cap)
    if cash_cap >= 1.0:
        return cons
    if cash_cap < 0:
        raise ValueError("cash_cap must be >= 0")

    if "CASH" not in columns:
        raise ValueError("cash_cap requested but 'CASH' not in columns/universe.")

    cash_idx = columns.index("CASH")
    cons += [w[cash_idx] <= cash_cap]
    return cons


def relax_region_floors_keep_caps(region_bounds: dict) -> dict:
    """
    Risk-off behavior:
      - Keep region upper bounds (caps)
      - Set all lower bounds to 0
    """
    if not region_bounds:
        return region_bounds
    out = {}
    for reg, (lo, hi) in region_bounds.items():
        out[reg] = (0.0, float(hi))
    return out


def compute_risk_off_signal(px_window: pd.DataFrame,
                            bench_weights_for_risk: dict,
                            ma_days: int = 200) -> bool:
    """
    Simple risk-off signal:
      - Build benchmark from assets in px_window (weighted)
      - Risk-off if benchmark price < its MA(ma_days)

    px_window must contain the benchmark components as columns.
    """
    if not bench_weights_for_risk:
        return False

    w = pd.Series(bench_weights_for_risk, dtype=float)
    w = w / w.sum()

    missing = [c for c in w.index if c not in px_window.columns]
    if missing:
        # If the risk benchmark can't be built, fail safe to risk-on
        return False

    bench_px = (px_window[w.index] * w).sum(axis=1).dropna()
    if len(bench_px) < ma_days + 5:
        return False

    ma = bench_px.rolling(int(ma_days)).mean()
    return bool(bench_px.iloc[-1] < ma.iloc[-1])

def min_variance_weights(Sigma, universe, columns,
                         region_bounds=None,
                         caps_by_kind=None,
                         max_weight_by_kind=None,
                         min_weight_by_kind=None,
                         cash_floor=None,
                         cash_cap=None,
                         w_prev=None,
                         max_turnover=None):
    """
    Solve a pure QP: minimize variance w' Sigma w subject to the SAME linear constraints.
    This ignores vol_target and lambda/objective. It tells you the minimum achievable vol.
    """
    n = Sigma.shape[0]
    w = cp.Variable(n)

    cap_vec = build_cap_vector(universe, columns, caps_by_kind or {}, default_cap=0.10)
    feasibility_check(universe, columns, cap_vec, region_bounds, max_weight_by_kind, min_weight_by_kind)

    cons = [cp.sum(w) == 1, w >= 0, w <= cap_vec]

    cons = add_turnover_constraint(cons, w, w_prev, max_turnover)

    if region_bounds:
        for reg, (lo, hi) in region_bounds.items():
            idx = indices_by_filter(universe, columns, "region", reg)
            if idx:
                cons += [cp.sum(w[idx]) >= float(lo), cp.sum(w[idx]) <= float(hi)]

    if max_weight_by_kind:
        for kind, cap in max_weight_by_kind.items():
            idx = indices_by_filter(universe, columns, "kind", kind)
            if idx:
                cons += [cp.sum(w[idx]) <= float(cap)]

    if min_weight_by_kind:
        for kind, floor in min_weight_by_kind.items():
            idx = indices_by_filter(universe, columns, "kind", kind)
            if idx:
                cons += [cp.sum(w[idx]) >= float(floor)]

    cons = add_min_cash_floor(cons, w, columns, cash_floor)
    cons = add_max_cash_cap(cons, w, columns, cash_cap)

    prob = cp.Problem(cp.Minimize(cp.quad_form(w, Sigma)), cons)
    prob.solve(solver=cp.OSQP, eps_abs=1e-7, eps_rel=1e-7, max_iter=200000, verbose=False)

    if w.value is None:
        return None
    wv = np.array(w.value).ravel()
    s = wv.sum()
    if s <= 0:
        return None
    return wv / s

def choose_solver_for_problem(needs_conic: bool) -> str:
    """
    needs_conic=True  -> SOC constraints present (e.g., vol cap) => conic solver required.
    needs_conic=False -> pure QP => OSQP preferred.
    """
    installed = set(cp.installed_solvers())

    if needs_conic:
        # Order: ECOS (fast), CLARABEL (very stable), SCS (robust fallback)
        for s in ["ECOS", "CLARABEL", "SCS"]:
            if s in installed:
                return s
        raise RuntimeError("Conic solver required (ECOS/CLARABEL/SCS) but none installed.")
    else:
        if "OSQP" in installed:
            return "OSQP"
        # Fallbacks can solve QPs too
        for s in ["ECOS", "CLARABEL", "SCS"]:
            if s in installed:
                return s
        raise RuntimeError("No suitable solver installed.")
    
def solve_problem_with_fallback(prob: cp.Problem, needs_conic: bool) -> None:
    installed = set(cp.installed_solvers())

    if needs_conic:
        candidates = ["ECOS", "CLARABEL", "SCS"]
    else:
        candidates = ["OSQP", "ECOS", "CLARABEL", "SCS"]

    candidates = [s for s in candidates if s in installed]
    if not candidates:
        raise RuntimeError("No candidate solvers available in this environment.")

    last_err = None
    for s in candidates:
        try:
            if s == "OSQP":
                prob.solve(solver=cp.OSQP, eps_abs=1e-6, eps_rel=1e-6, max_iter=200000, verbose=False)
            elif s == "ECOS":
                prob.solve(solver=cp.ECOS, abstol=1e-8, reltol=1e-8, feastol=1e-8, max_iters=50000, verbose=False)
            elif s == "CLARABEL":
                prob.solve(solver=cp.CLARABEL, verbose=False)
            elif s == "SCS":
                prob.solve(solver=cp.SCS, eps=1e-5, max_iters=200000, verbose=False)
            else:
                prob.solve(solver=s, verbose=False)

            # Accept optimal and optimal_inaccurate; reject infeasible/failed
            if prob.status in ("optimal", "optimal_inaccurate"):
                return
            last_err = RuntimeError(f"Solver {s} ended with status={prob.status}")
        except Exception as e:
            last_err = e

    # If we got here: nothing worked
    raise RuntimeError(f"All candidate solvers failed. Last error: {last_err}")



def min_l1_needed_for_cash_rail(w_prev, columns, cash_floor=None, cash_cap=None) -> float:
    """
    Minimum L1 turnover needed to satisfy cash_floor/cash_cap given previous weights.
    If you must increase CASH by delta, L1 >= 2*delta (increase cash + decrease others).
    Same if you must decrease CASH by delta.
    """
    if w_prev is None:
        return 0.0
    if "CASH" not in columns:
        return 0.0

    cash_idx = columns.index("CASH")
    w_prev_cash = float(np.asarray(w_prev).ravel()[cash_idx])

    req = 0.0
    if cash_floor is not None:
        delta_up = max(0.0, float(cash_floor) - w_prev_cash)
        req = max(req, 2.0 * delta_up)

    if cash_cap is not None:
        delta_down = max(0.0, w_prev_cash - float(cash_cap))
        req = max(req, 2.0 * delta_down)

    return float(req)

def normalize_rebalance(freq: str) -> str:
    # try as-is
    try:
        pd.tseries.frequencies.to_offset(freq)
        return freq
    except Exception:
        pass

    # try common alias swaps (old <-> new)
    swaps = {"M":"ME","Q":"QE","Y":"YE","BM":"BME","BQ":"BQE","BY":"BYE",
             "ME":"M","QE":"Q","YE":"Y","BME":"BM","BQE":"BQ","BYE":"BY"}
    if freq in swaps:
        return swaps[freq]

    # last resort: raise
    pd.tseries.frequencies.to_offset(freq)  # will raise with good message
    return freq

def assert_required_sleeves_exist(universe, columns, region_bounds):
    missing = []
    for reg, (lo, hi) in (region_bounds or {}).items():
        if lo > 0 and not indices_by_filter(universe, columns, "region", reg):
            missing.append(reg)
    if missing:
        raise RuntimeError(f"Missing assets for regions with positive floors: {missing}")


TRADING_DAYS = 252


# ============================================================
# 1) DATA: STOOQ CLOSE LOADER
# ============================================================
def stooq_close(ticker: str) -> pd.Series:
    """
    Downloads daily close from Stooq.
    Returns a Series indexed by Date.
    """
    urls = [
        f"https://stooq.com/q/d/l/?s={ticker}&i=d",
        f"https://stooq.pl/q/d/l/?s={ticker}&i=d",
    ]
    last_err = None
    for url in urls:
        try:
            df = pd.read_csv(url)
            if "Date" not in df.columns or "Close" not in df.columns:
                raise ValueError(f"Bad response. Columns: {df.columns.tolist()[:8]} (url: {url})")
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df = df.dropna(subset=["Date"]).set_index("Date").sort_index()
            if df.empty:
                raise ValueError(f"Empty after parsing dates (url: {url})")
            s = df["Close"].astype(float)
            s = s[~s.index.duplicated(keep="last")]
            return s.rename(ticker)
        except Exception as e:
            last_err = e
    raise ValueError(f"No usable data for {ticker}. Last error: {last_err}")

# ============================================================
# 2) FX CONVERSION (EUR BASE) — Stooq FX only
#    Assume FX series quoted as CCY per EUR (EURUSD = USD per EUR)
#    => EUR price = CCY price / FX
# ============================================================
def fx_candidates_stooq(ccy: str):
    ccy = ccy.upper()
    if ccy == "USD":
        return ["eurusd", "eurusd.fx", "eurusd.f", "eurusd.x"]
    if ccy == "GBP":
        return ["eurgbp", "eurgbp.fx", "eurgbp.f", "eurgbp.x"]
    if ccy == "HKD":
        return ["eurhkd", "eurhkd.fx", "eurhkd.f", "eurhkd.x"]
    if ccy == "JPY":
        return ["eurjpy", "eurjpy.fx", "eurjpy.f", "eurjpy.x"]
    return []

def get_fx_series_to_eur(ccy: str, fx_cache: dict) -> pd.Series:
    ccy = ccy.upper()
    if ccy in fx_cache:
        return fx_cache[ccy]
    last_err = None
    for t in fx_candidates_stooq(ccy):
        try:
            s = stooq_close(t).rename(f"FX_EUR{ccy}")
            fx_cache[ccy] = s
            return s
        except Exception as e:
            last_err = e
    raise ValueError(f"No FX series found for EUR{ccy} on Stooq. Last error: {last_err}")

def convert_prices_to_eur(px_ccy: pd.Series, ccy: str, fx_cache: dict) -> pd.Series:
    ccy = (ccy or "EUR").upper()
    if ccy == "EUR":
        return px_ccy

    fx = get_fx_series_to_eur(ccy, fx_cache)
    df = pd.concat([px_ccy, fx], axis=1, sort=True)

    # FX gaps: forward fill is usually reasonable
    df.iloc[:, 1] = df.iloc[:, 1].ffill()

    df = df.dropna(subset=[df.columns[0], df.columns[1]])
    eur = (df.iloc[:, 0] / df.iloc[:, 1]).rename(px_ccy.name)
    return eur


# ============================================================
# 3) BUILD EUR PRICE MATRIX (outer-join; CASH supported)
# ============================================================
def build_eur_price_matrix(universe):
    """
    universe items: {name, ticker, ccy, scale(optional), kind, region}
    Supports synthetic CASH asset with ticker=None.
    """
    fx_cache = {}
    dropped = []
    series = {}

    wants_cash = False

    for a in universe:
        name = a["name"]
        ticker = a.get("ticker", None)
        ccy = a.get("ccy", "EUR")
        scale = a.get("scale", 1.0)

        if ticker is None:
            # Synthetic asset (e.g. CASH)
            wants_cash = wants_cash or (name.upper() == "CASH")
            continue

        try:
            s = (stooq_close(ticker) * scale).rename(name)
            s_eur = convert_prices_to_eur(s, ccy, fx_cache)
            series[name] = s_eur
        except Exception as e:
            dropped.append((name, ticker, str(e)))

    if not series:
        raise RuntimeError("No assets downloaded. Check tickers / internet / Stooq.")

    eur_px = pd.concat(series.values(), axis=1, sort=True)

    # Add synthetic CASH as constant price series
    if wants_cash:
        eur_px["CASH"] = 1.0

    if dropped:
        print("Dropped assets:")
        for name, tick, msg in dropped:
            print(f"  - {name} ({tick}): {msg}")

    return eur_px


# ============================================================
# 4) RETURNS (simple + log), winsorization
# ============================================================
def simple_returns_from_prices(px: pd.DataFrame) -> pd.DataFrame:
    return px.pct_change().dropna(how="all")

def log_returns_from_prices(px: pd.DataFrame) -> pd.DataFrame:
    return np.log(px).diff().dropna(how="all")

def log_returns_from_nav(nav: pd.Series) -> pd.Series:
    nav = nav.dropna()
    return np.log(nav).diff().dropna()

def winsorize_std(df: pd.DataFrame, z=4.0) -> pd.DataFrame:
    x = df.copy()
    for c in x.columns:
        s = x[c].std(ddof=1)
        if s and s > 0:
            x[c] = x[c].clip(lower=-z*s, upper=z*s)
    return x

def winsorize_quantile(df: pd.DataFrame, q=0.01) -> pd.DataFrame:
    x = df.copy()
    lo = x.quantile(q)
    hi = x.quantile(1 - q)
    return x.clip(lower=lo, upper=hi, axis=1)


# ============================================================
# 5) ESTIMATION: mu, Sigma (rolling window)
# ============================================================
def estimate_mu_sigma(ret_window: pd.DataFrame,
                      mean_shrink=0.5,
                      cov_method="lw",
                      winsor_mode="std",
                      winsor_param=4.0):
    """
    ret_window: daily SIMPLE returns recommended.
    mu is annualized arithmetic mean.
    Sigma is annualized covariance.
    """
    rw = ret_window.copy()

    if winsor_mode == "std":
        rw = winsorize_std(rw, z=float(winsor_param))
    elif winsor_mode == "quantile":
        rw = winsorize_quantile(rw, q=float(winsor_param))
    elif winsor_mode in (None, "none"):
        pass
    else:
        raise ValueError("winsor_mode must be 'std', 'quantile', or None")

    mu_sample = (rw.mean() * TRADING_DAYS).to_numpy()
    mu_bar = float(mu_sample.mean())
    mu0 = np.full_like(mu_sample, mu_bar)
    mu = (1 - mean_shrink) * mu_sample + mean_shrink * mu0

    if cov_method == "lw":
        lw = LedoitWolf().fit(rw.to_numpy())
        Sigma = lw.covariance_ * TRADING_DAYS
    elif cov_method == "sample":
        Sigma = (rw.cov() * TRADING_DAYS).to_numpy()
    else:
        raise ValueError("cov_method must be 'lw' or 'sample'")

    # PSD guard
    vals, vecs = np.linalg.eigh(Sigma)
    vals = np.maximum(vals, 1e-10)
    Sigma = (vecs * vals) @ vecs.T
    return mu, Sigma


# ============================================================
# 6) CONSTRAINT HELPERS
# ============================================================
def _solve_qp_single(mu, Sigma, universe, columns,
                     rf=0.0,
                     lam=10.0,
                     region_bounds=None,
                     caps_by_kind=None,
                     max_weight_by_kind=None,
                     min_weight_by_kind=None,
                     w_prev=None,
                     turn_gamma=10.0,
                     l2=1e-4,
                     max_turnover=None,          # hard turnover L1
                     vol_target=None,            # annual vol cap (SOC)
                     cash_floor=None,            # CASH min
                     cash_cap=None,              # CASH max
                     zero_tol=1e-6,
                     _retry_relax_turnover=True):

    n = len(mu)
    w = cp.Variable(n)
    excess = mu - float(rf)

    cap_vec = build_cap_vector(universe, columns, caps_by_kind or {}, default_cap=0.10)
    feasibility_check(universe, columns, cap_vec, region_bounds, max_weight_by_kind, min_weight_by_kind)

    cons = [cp.sum(w) == 1, w >= 0, w <= cap_vec]

    # Turnover rail
    cons = add_turnover_constraint(cons, w, w_prev, max_turnover)

    # Region bounds
    if region_bounds:
        for reg, (lo, hi) in region_bounds.items():
            idx = indices_by_filter(universe, columns, "region", reg)
            if idx:
                cons += [cp.sum(w[idx]) >= float(lo), cp.sum(w[idx]) <= float(hi)]

    # Sleeve caps/floors by kind
    if max_weight_by_kind:
        for kind, cap in max_weight_by_kind.items():
            idx = indices_by_filter(universe, columns, "kind", kind)
            if idx:
                cons += [cp.sum(w[idx]) <= float(cap)]

    if min_weight_by_kind:
        for kind, floor in min_weight_by_kind.items():
            idx = indices_by_filter(universe, columns, "kind", kind)
            if idx:
                cons += [cp.sum(w[idx]) >= float(floor)]

    # Vol + cash rails
    needs_conic = vol_target is not None
    cons = add_volatility_cap(cons, w, Sigma, vol_target)
    cons = add_min_cash_floor(cons, w, columns, cash_floor)
    cons = add_max_cash_cap(cons, w, columns, cash_cap)

    # Objective (mean-variance with small L2 + turnover smoothing)
    obj = (excess @ w
           - float(lam) * cp.quad_form(w, Sigma)
           - float(l2) * cp.sum_squares(w))
    if (w_prev is not None) and (turn_gamma is not None) and (float(turn_gamma) > 0):
        obj = obj - float(turn_gamma) * cp.sum_squares(w - np.asarray(w_prev).ravel())

    prob = cp.Problem(cp.Maximize(obj), cons)

    try:
        solve_problem_with_fallback(prob, needs_conic=needs_conic)
    except Exception:
        # Optional one-shot retry by relaxing turnover (common infeasibility culprit)
        if _retry_relax_turnover and (max_turnover is not None) and (w_prev is not None):
            relaxed = min(1.0, max(float(max_turnover) * 4.0, float(max_turnover) + 0.25))
            return _solve_qp_single(
                mu, Sigma, universe, columns,
                rf=rf, lam=lam,
                region_bounds=region_bounds,
                caps_by_kind=caps_by_kind,
                max_weight_by_kind=max_weight_by_kind,
                min_weight_by_kind=min_weight_by_kind,
                w_prev=w_prev,
                turn_gamma=turn_gamma,
                l2=l2,
                max_turnover=relaxed,
                vol_target=vol_target,
                cash_floor=cash_floor,
                cash_cap=cash_cap,
                zero_tol=zero_tol,
                _retry_relax_turnover=False
            )
        return None

    if w.value is None:
        return None

    wv = np.asarray(w.value).ravel()
    wv[wv < float(zero_tol)] = 0.0
    s = float(wv.sum())
    if s <= 0:
        return None
    return wv / s

def solve_qp_best_sharpe(mu, Sigma, universe, columns,
                         rf=0.0,
                         region_bounds=None,
                         caps_by_kind=None,
                         max_weight_by_kind=None,
                         min_weight_by_kind=None,
                         w_prev=None,
                         turn_gamma=10.0,
                         l2=1e-4,
                         lam_grid=None,
                         max_turnover=None,
                         vol_target=None,
                         cash_floor=None,
                         cash_cap=None,
                         zero_tol=1e-6):
    if lam_grid is None:
        lam_grid = np.logspace(-1, 2, 25)

    best = None
    for lam in lam_grid:
        wv = _solve_qp_single(
            mu, Sigma, universe, columns,
            rf=rf, lam=float(lam),
            region_bounds=region_bounds,
            caps_by_kind=caps_by_kind,
            max_weight_by_kind=max_weight_by_kind,
            min_weight_by_kind=min_weight_by_kind,
            w_prev=w_prev,
            turn_gamma=turn_gamma,
            l2=l2,
            max_turnover=max_turnover,
            vol_target=vol_target,
            cash_floor=cash_floor,
            cash_cap=cash_cap,
            zero_tol=zero_tol
        )
        if wv is None:
            continue

        pret = float(mu @ wv)
        pvol = float(np.sqrt(max(1e-18, wv @ Sigma @ wv)))
        sharpe = (pret - float(rf)) / pvol if pvol > 0 else -np.inf

        if (best is None) or (sharpe > best["sharpe"]):
            best = {"w": wv, "ret": pret, "vol": pvol, "sharpe": sharpe, "lam": float(lam)}

    if best is None:
        raise RuntimeError("No feasible solution across lambda grid with current rails/constraints.")
    return best



def compute_risk_off_signal(px_window: pd.DataFrame,
                            bench_weights_for_risk: dict,
                            ma_days: int = 200) -> bool:
    """
    Simple risk-off signal:
      - Build benchmark from assets in px_window (weighted)
      - Risk-off if benchmark price < its MA(ma_days)

    px_window must contain the benchmark components as columns.
    """
    if not bench_weights_for_risk:
        return False

    w = pd.Series(bench_weights_for_risk, dtype=float)
    w = w / w.sum()

    missing = [c for c in w.index if c not in px_window.columns]
    if missing:
        # If the risk benchmark can't be built, fail safe to risk-on
        return False

    bench_px = (px_window[w.index] * w).sum(axis=1).dropna()
    if len(bench_px) < ma_days + 5:
        return False

    ma = bench_px.rolling(int(ma_days)).mean()
    return bool(bench_px.iloc[-1] < ma.iloc[-1])



# ============================================================
# 8) Tactical EM floor (optional)
# ============================================================
def momentum_signal(prices: pd.Series, lookback_days=126) -> bool:
    s = prices.dropna()
    if len(s) < lookback_days + 5:
        return False
    return float(s.iloc[-1] / s.iloc[-lookback_days] - 1) > 0.0

def dynamic_region_bounds(base_bounds, px_window, em_asset_name="EEM",
                          em_floor_when_on=0.05, em_mom_lookback=126):
    b = dict(base_bounds) if base_bounds else {}
    if em_asset_name in px_window.columns:
        em_on = momentum_signal(px_window[em_asset_name], lookback_days=em_mom_lookback)
        lo, hi = b.get("EM", (0.0, 0.20))
        lo = max(lo, em_floor_when_on) if em_on else 0.0
        b["EM"] = (lo, hi)
    return b


# ============================================================
# 8b) Risk regime proxy for cash floor (Rail)
# ============================================================
def bench_price_proxy(px_window: pd.DataFrame, bench_weights: dict) -> pd.Series:
    w = pd.Series(bench_weights, dtype=float)
    w = w / w.sum()
    missing = [c for c in w.index if c not in px_window.columns]
    if missing:
        return None
    df = px_window[w.index].dropna(how="any")
    if df.empty:
        return None
    return (df * w).sum(axis=1)

def is_risk_off(bench_px: pd.Series, ma_days=200) -> bool:
    if bench_px is None:
        return False
    s = bench_px.dropna()
    if len(s) < ma_days + 5:
        return False
    ma = s.rolling(ma_days).mean()
    return bool(s.iloc[-1] < ma.iloc[-1])


# ============================================================
# 9) WALK-FORWARD BACKTEST (rails included)
# ============================================================
def backtest_allocator(eur_px: pd.DataFrame, universe,
                       window_years=5,
                       rebalance="ME",
                       rf=0.0,
                       mean_shrink=0.5,
                       cov_method="lw",
                       winsor_mode="std",
                       winsor_param=4.0,
                       region_bounds=None,
                       caps_by_kind=None,
                       max_weight_by_kind=None,
                       min_weight_by_kind=None,
                       turn_gamma=10.0,
                       l2=1e-4,
                       lambda_mode="grid",
                       lam_fixed=10.0,
                       lam_grid=None,
                       val_months=12,
                       tactical_em=False,
                       em_floor_when_on=0.05,
                       cost_bps=0.0,
                       require_full_window=True,
                       # NEW rails:
                       max_turnover=None,
                       min_rebalance_l1=None,          # skip rebalance if |Δw|_1 too small
                       vol_target=None,
                       risk_off_enabled=False,
                       bench_weights_for_risk=None,
                       risk_off_ma_days=200,
                       cash_floor_risk_off=0.20,
                       cash_floor_risk_on=0.00,
                       cash_cap_risk_off=0.30,         # NEW
                       cash_cap_risk_on=0.25):         # NEW

    """
    Returns:
        nav (Series), weights (DataFrame), lambda_used (Series)
    """
    if rebalance == "M":
        rebalance = "ME"

    px = eur_px.sort_index()
    rebal_dates = px.resample(normalize_rebalance(rebalance)).last().index

    cols = list(px.columns)
    assert_required_sleeves_exist(universe, cols, region_bounds)
    w_prev = None
    w_hist = []
    nav_parts = []
    lam_hist = []
    nav_level = 1.0
    skip_empty_win = 0
    skip_short_full_window = 0
    skip_short_returns = 0
    skip_empty_fwd = 0
    used_periods = 0
    for i, dt in enumerate(rebal_dates):
        start = dt - pd.DateOffset(years=window_years)
        px_win = px.loc[start:dt]

        if require_full_window:
            if px_win.empty:
                skip_empty_win += 1
                continue
            if (dt - px_win.index.min()) < pd.Timedelta(days=int(365 * (window_years - 0.05))):
                skip_short_full_window += 1
                continue

        px_win = px.loc[start:dt].dropna(how="any")

        if require_full_window:
            if px_win.empty:
                skip_empty_win += 1
                continue
            if (dt - px_win.index.min()) < pd.Timedelta(days=int(365 * (window_years - 0.05))):
                skip_short_full_window += 1
                continue

        r_win = simple_returns_from_prices(px_win).dropna(how="any")
        if len(r_win) < TRADING_DAYS * 2:
            skip_short_returns += 1
            continue
        
        

        # --- RISK-OFF SIGNAL ---
        risk_off = False
        if risk_off_enabled:
            risk_off = compute_risk_off_signal(
                px_win,
                bench_weights_for_risk=bench_weights_for_risk,
                ma_days=risk_off_ma_days
            )

        # --- REGION BOUNDS (relax floors in risk-off) ---
        rb = region_bounds
        if tactical_em and region_bounds is not None:
            rb = dynamic_region_bounds(region_bounds, px_win, em_floor_when_on=em_floor_when_on)

        if risk_off and rb is not None:
            rb = relax_region_floors_keep_caps(rb)


        # 1) estimate mu, Sigma first  (Sigma must exist before any vol checks)
        mu, Sigma = estimate_mu_sigma(r_win, mean_shrink, cov_method, winsor_mode, winsor_param)

        # 2) regime -> cash rails
        cash_floor = cash_floor_risk_off if (risk_off_enabled and risk_off) else cash_floor_risk_on
        cash_cap   = cash_cap_risk_off   if (risk_off_enabled and risk_off) else cash_cap_risk_on

        # 3) turnover cap effective (relax if cash rail would force impossible trade)
        max_turnover_eff = max_turnover
        if (max_turnover is not None) and (w_prev is not None):
            req_l1 = min_l1_needed_for_cash_rail(w_prev, cols, cash_floor=cash_floor, cash_cap=cash_cap)
            if req_l1 > float(max_turnover):
                max_turnover_eff = min(1.0, req_l1 * 1.01)

        # 4) volatility cap effective (do NOT mutate the global vol_target)
        vol_cap_eff = vol_target

        # 5) optional vol feasibility pre-check (now Sigma exists)
        if vol_cap_eff is not None:
            w_minvar = min_variance_weights(
                Sigma, universe, cols,
                region_bounds=rb,
                caps_by_kind=caps_by_kind,
                max_weight_by_kind=max_weight_by_kind,
                min_weight_by_kind=min_weight_by_kind,
                cash_floor=cash_floor,
                cash_cap=cash_cap,
                w_prev=w_prev,
                max_turnover=max_turnover_eff
            )
            if w_minvar is not None:
                min_vol = float(np.sqrt(w_minvar @ Sigma @ w_minvar))
                if min_vol > float(vol_cap_eff) + 1e-6:
                    vol_cap_eff = min_vol * 1.01
            else:
                vol_cap_eff = None  # safest fallback if min-var itself cannot be solved

        # choose lambda + solve
        if lambda_mode == "grid":
            try:
                best = solve_qp_best_sharpe(
                    mu, Sigma, universe, cols,
                    rf=rf,
                    region_bounds=rb,
                    caps_by_kind=caps_by_kind,
                    max_weight_by_kind=max_weight_by_kind,
                    min_weight_by_kind=min_weight_by_kind,
                    w_prev=w_prev,
                    turn_gamma=turn_gamma,
                    l2=l2,
                    lam_grid=lam_grid,
                    max_turnover=max_turnover_eff,
                    vol_target=vol_cap_eff,
                    cash_floor=cash_floor,
                    cash_cap=cash_cap
                )
                w = best["w"]
                lam_used = best["lam"]
            except RuntimeError:
                # Minimal, “safety-first” fallback:
                # 1) if we have previous weights, keep them (no trade) rather than crash
                if w_prev is not None:
                    w = w_prev.copy()
                    lam_used = np.nan
                else:
                    # 2) if no previous weights yet, try disabling vol cap for first allocation
                    best = solve_qp_best_sharpe(
                        mu, Sigma, universe, cols,
                        rf=rf,
                        region_bounds=rb,
                        caps_by_kind=caps_by_kind,
                        max_weight_by_kind=max_weight_by_kind,
                        min_weight_by_kind=min_weight_by_kind,
                        w_prev=w_prev,
                        turn_gamma=turn_gamma,
                        l2=l2,
                        lam_grid=lam_grid,
                        max_turnover=max_turnover_eff,
                        vol_target=None,          # relax only the SOC rail
                        cash_floor=cash_floor,
                        cash_cap=cash_cap
                    )
                    w = best["w"]
                    lam_used = best["lam"]
        elif lambda_mode == "fixed":
            lam_used = float(lam_fixed)
            w = _solve_qp_single(
                mu, Sigma, universe, cols,
                rf=rf, lam=lam_used,
                region_bounds=rb,
                caps_by_kind=caps_by_kind,
                max_weight_by_kind=max_weight_by_kind,
                min_weight_by_kind=min_weight_by_kind,
                w_prev=w_prev,
                turn_gamma=turn_gamma,
                l2=l2,
                max_turnover=max_turnover,
                vol_target=vol_target,
                cash_floor=cash_floor,
                cash_cap=cash_cap
            )
            if w is None:
                continue
        else:
            raise ValueError("lambda_mode must be 'grid' or 'fixed' for this version.")

        # --- SKIP TINY REBALANCES ---
        if w_prev is not None and min_rebalance_l1 is not None:
            dw = float(np.sum(np.abs(w - w_prev)))
            if dw < float(min_rebalance_l1):
                # keep previous weights, do not trade, keep same lambda record
                w = w_prev.copy()

        # costs at rebalance (still uses w vs w_prev)
        if w_prev is not None and cost_bps > 0:
            to = float(np.sum(np.abs(w - w_prev)))
            traded_notional = to / 2.0
            nav_level *= (1.0 - (cost_bps / 10000.0) * traded_notional)

        w_prev = w.copy()
        w_hist.append(pd.Series(w, index=cols, name=dt))
        lam_hist.append(pd.Series({"lambda": lam_used}, name=dt))

        # Forward slice uses complete-data intersection too (execution realism)
        dt_next = rebal_dates[i + 1] if (i + 1) < len(rebal_dates) else px.index.max()
        px_fwd = px.loc[(px.index > dt) & (px.index <= dt_next)].dropna(how="any")
        if len(px_fwd) < 2:
            skip_empty_fwd += 1
            continue

        r_fwd = simple_returns_from_prices(px_fwd).dropna(how="any")
        if r_fwd.empty:
            skip_empty_fwd += 1
            continue

        rp = pd.Series(r_fwd.to_numpy() @ w, index=r_fwd.index)
        nav_path = nav_level * (1.0 + rp).cumprod()
        nav_level = float(nav_path.iloc[-1])
        nav_parts.append(nav_path)
        used_periods += 1

    if not nav_parts:
        raise RuntimeError(
            "No backtest periods produced. "
            f"empty_win={skip_empty_win}, "
            f"short_full_window={skip_short_full_window}, "
            f"short_returns={skip_short_returns}, "
            f"empty_fwd={skip_empty_fwd}, "
            f"rebal_dates={len(rebal_dates)}"
        )
    nav = pd.concat(nav_parts).sort_index()
    weights = pd.DataFrame(w_hist).sort_index()
    lambdas = pd.DataFrame(lam_hist).sort_index()["lambda"]
    return nav, weights, lambdas


# ============================================================
# 10) DIAGNOSTICS + BENCHMARK + BETA
# ============================================================
def turnover(weights: pd.DataFrame) -> pd.Series:
    return weights.diff().abs().sum(axis=1)

def bound_hit_rate(weights: pd.DataFrame, universe, region_bounds, atol=1e-3):
    if not region_bounds:
        return None
    cols = list(weights.columns)
    out = {}
    for reg, (lo, hi) in region_bounds.items():
        idx = indices_by_filter(universe, cols, "region", reg)
        if not idx:
            continue
        s = weights.iloc[:, idx].sum(axis=1)
        out[f"{reg}_near_lo"] = float(np.isclose(s, lo, atol=atol).mean())
        out[f"{reg}_near_hi"] = float(np.isclose(s, hi, atol=atol).mean())
    return pd.Series(out)

def make_custom_benchmark(asset_rets: pd.DataFrame, bench_weights: dict) -> pd.Series:
    w = pd.Series(bench_weights, dtype=float)
    w = w / w.sum()
    missing = [c for c in w.index if c not in asset_rets.columns]
    if missing:
        raise ValueError(f"Benchmark components missing from columns: {missing}")
    r_m = (asset_rets[w.index] * w).sum(axis=1)
    return r_m.rename("BENCH")

def beta(x: pd.Series, m: pd.Series) -> float:
    df = pd.concat([x, m], axis=1, sort=False).dropna()
    if len(df) < 60:
        return np.nan
    x = df.iloc[:, 0]
    m = df.iloc[:, 1]
    var_m = float(np.var(m, ddof=1))
    if var_m <= 0:
        return np.nan
    cov_xm = float(np.cov(x, m, ddof=1)[0, 1])
    return cov_xm / var_m

def rolling_beta(x: pd.Series, m: pd.Series, window=252) -> pd.Series:
    df = pd.concat([x, m], axis=1, sort=False).dropna()
    x = df.iloc[:, 0]
    m = df.iloc[:, 1]
    rb = x.rolling(window).cov(m) / m.rolling(window).var()
    return rb.rename(f"beta_{window}d")

def ann_return_vol_sharpe(r: pd.Series, rf=0.0):
    mu = float(r.mean() * TRADING_DAYS)
    vol = float(r.std(ddof=1) * np.sqrt(TRADING_DAYS))
    sh = (mu - rf) / vol if vol > 0 else np.nan
    return mu, vol, sh

def tracking_error(port_r: pd.Series, bench_r: pd.Series):
    df = pd.concat([port_r, bench_r], axis=1, sort=False).dropna()
    diff = df.iloc[:, 0] - df.iloc[:, 1]
    te = float(diff.std(ddof=1) * np.sqrt(TRADING_DAYS))
    corr = float(df.iloc[:, 0].corr(df.iloc[:, 1]))
    return te, corr


# ============================================================
# MINI MONTE CARLO / MULTI-RUN EXPERIMENTS (rails supported)
# ============================================================
def _max_drawdown(nav: pd.Series) -> float:
    s = nav.dropna()
    peak = s.cummax()
    dd = (s / peak) - 1.0
    return float(dd.min()) if len(dd) else np.nan

def _cvar(r: pd.Series, alpha=0.05) -> float:
    x = r.dropna().to_numpy()
    if len(x) == 0:
        return np.nan
    q = np.quantile(x, alpha)
    tail = x[x <= q]
    return float(tail.mean()) if len(tail) else float(q)

def run_mini_monte_carlo(
    eur_px: pd.DataFrame,
    universe: list,
    region_bounds: dict,
    caps_by_kind: dict,
    max_weight_by_kind: dict,
    min_weight_by_kind: dict,
    bench_w: dict,
    experiments: list,
    rf: float = 0.0,
    verbose: bool = True,
) -> pd.DataFrame:

    asset_rets = simple_returns_from_prices(eur_px).dropna(how="any")
    bench_rets = make_custom_benchmark(asset_rets, bench_w)

    rows = []
    for i, cfg in enumerate(experiments, start=1):
        label = cfg.get("label", f"run_{i}")

        if verbose:
            print(f"\n=== [{i}/{len(experiments)}] {label} ===")

        params = dict(
            window_years=5,
            rebalance="ME",
            rf=rf,
            mean_shrink=0.5,
            cov_method="lw",
            winsor_mode="std",
            winsor_param=4.0,
            region_bounds=region_bounds,
            caps_by_kind=caps_by_kind,
            max_weight_by_kind=max_weight_by_kind,
            min_weight_by_kind=min_weight_by_kind,
            turn_gamma=10.0,
            l2=1e-4,
            lambda_mode="grid",
            lam_fixed=10.0,
            lam_grid=np.logspace(-1, 2, 25),
            val_months=12,
            tactical_em=False,
            em_floor_when_on=0.05,
            cost_bps=0.0,
            require_full_window=True,
            # rails
            max_turnover=0.10,
            min_rebalance_l1=0.03,
            vol_target=0.12,
            risk_off_enabled=True,
            bench_weights_for_risk=bench_w,
            risk_off_ma_days=200,
            cash_floor_risk_off=0.20,
            cash_floor_risk_on=0.00
        )
        params.update(cfg)

        try:
            nav, weights, lambdas = backtest_allocator(
                eur_px, universe,
                window_years=params["window_years"],
                rebalance=params["rebalance"],
                rf=params["rf"],
                mean_shrink=params["mean_shrink"],
                cov_method=params["cov_method"],
                winsor_mode=params["winsor_mode"],
                winsor_param=params["winsor_param"],
                region_bounds=params["region_bounds"],
                caps_by_kind=params["caps_by_kind"],
                max_weight_by_kind=params["max_weight_by_kind"],
                min_weight_by_kind=params["min_weight_by_kind"],
                turn_gamma=params["turn_gamma"],
                l2=params["l2"],
                lambda_mode=params["lambda_mode"],
                lam_fixed=params["lam_fixed"],
                lam_grid=params["lam_grid"],
                val_months=params["val_months"],
                tactical_em=params["tactical_em"],
                em_floor_when_on=params["em_floor_when_on"],
                cost_bps=params["cost_bps"],
                require_full_window=params["require_full_window"],
                max_turnover=params["max_turnover"],
                min_rebalance_l1=params["min_rebalance_l1"],
                vol_target=params["vol_target"],
                risk_off_enabled=params["risk_off_enabled"],
                bench_weights_for_risk=params["bench_weights_for_risk"],
                risk_off_ma_days=params["risk_off_ma_days"],
                cash_floor_risk_off=params["cash_floor_risk_off"],
                cash_floor_risk_on=params["cash_floor_risk_on"],
            )

            port_r = nav.pct_change().dropna()
            df = pd.concat([port_r.rename("PORT"), bench_rets.rename("BENCH")], axis=1, sort=False).dropna()
            port_al = df["PORT"]
            bench_al = df["BENCH"]

            ann_ret, ann_vol, sh = ann_return_vol_sharpe(port_al, rf=rf)
            b_full = beta(port_al, bench_al)
            te, corr = tracking_error(port_al, bench_al)

            to = turnover(weights)
            to_mean = float(to.mean()) if len(to) else np.nan
            to_p95 = float(to.quantile(0.95)) if len(to) else np.nan

            row = dict(
                run=i,
                label=label,
                start=str(nav.index.min().date()),
                end=str(nav.index.max().date()),
                ann_return=ann_ret,
                ann_vol=ann_vol,
                sharpe=sh,
                max_dd=_max_drawdown(nav),
                cvar_5pct=_cvar(port_al, 0.05),
                beta_vs_bench=b_full,
                tracking_error=te,
                corr_vs_bench=corr,
                turnover_mean=to_mean,
                turnover_p95=to_p95,
                lambda_mode=params["lambda_mode"],
                cost_bps=params["cost_bps"],
                tactical_em=params["tactical_em"],
                mean_shrink=params["mean_shrink"],
                winsor_mode=params["winsor_mode"],
                winsor_param=params["winsor_param"],
                turn_gamma=params["turn_gamma"],
                l2=params["l2"],
                max_turnover=params["max_turnover"],
                vol_target=params["vol_target"],
                min_rebalance_l1=params["min_rebalance_l1"],
                cash_floor_risk_off=params["cash_floor_risk_off"],
            )

            if lambdas is not None and len(lambdas):
                row["lambda_median"] = float(pd.Series(lambdas).median())
                row["lambda_p90"] = float(pd.Series(lambdas).quantile(0.90))
            else:
                row["lambda_median"] = np.nan
                row["lambda_p90"] = np.nan

            rows.append(row)

            if verbose:
                print(f"Sharpe={row['sharpe']:.3f}  Ret={row['ann_return']:.3%}  Vol={row['ann_vol']:.3%}  "
                      f"TE={row['tracking_error']:.3%}  Turn(mean)={row['turnover_mean']:.3%}")

        except Exception as e:
            rows.append(dict(run=i, label=label, error=str(e)))
            if verbose:
                print(f"FAILED: {e}")

    out = pd.DataFrame(rows)
    if "sharpe" in out.columns:
        out = out.sort_values(by="sharpe", ascending=False, na_position="last")
    return out


# ============================================================
# 11) RUN
# ============================================================
if __name__ == "__main__":  

    from universe import UNIVERSE as universe   # ← 4 espaces, import local  

    eur_px = build_eur_price_matrix(universe)  
    print(f"EUR price matrix: {eur_px.index.min().date()} -> {eur_px.index.max().date()} | assets: {eur_px.shape[1]}")  

    region_bounds = {  
        "US": (0.25, 0.55),  
        "EU": (0.30, 0.65),  
        "EM": (0.00, 0.20),  
    }  

    caps_by_kind        = {"ETF": 0.40, "STOCK": 0.10, "GOLD": 0.15, "OTHER": 0.10, "CASH": 1.00}  
    max_weight_by_kind  = {"STOCK": 0.40, "GOLD": 0.15}  
    min_weight_by_kind  = {"ETF": 0.50}  
    bench_w             = {"MEUD": 0.50, "SPY": 0.30, "EEM": 0.10, "GLD": 0.10}  

    MAX_TURNOVER       = 0.10  
    MIN_REBAL_L1       = 0.03  
    VOL_TARGET         = 0.12  
    CASH_FLOOR_RISK_OFF = 0.20  
    CASH_FLOOR_RISK_ON  = 0.00  

    nav, weights, lambdas = backtest_allocator(
        eur_px, universe,
        window_years=5,
        rebalance="ME",             # consider "QE" for even calmer trading
        rf=0.0,
        mean_shrink=0.5,
        cov_method="lw",
        winsor_mode="std",
        winsor_param=4.0,
        region_bounds=region_bounds,
        caps_by_kind=caps_by_kind,
        max_weight_by_kind=max_weight_by_kind,
        min_weight_by_kind=min_weight_by_kind,
        turn_gamma=10.0,
        l2=1e-4,
        lambda_mode="grid",
        lam_grid=np.logspace(-1, 2, 25),
        tactical_em=False,
        em_floor_when_on=0.05,
        cost_bps=0.0,
        require_full_window=True,
        # rails
        max_turnover=MAX_TURNOVER,
        min_rebalance_l1=MIN_REBAL_L1,
        vol_target=VOL_TARGET,
        risk_off_enabled=True,
        bench_weights_for_risk=bench_w,
        risk_off_ma_days=200,
        cash_floor_risk_off=CASH_FLOOR_RISK_OFF,
        cash_floor_risk_on=CASH_FLOOR_RISK_ON,
        px_raw = px.loc[start:dt],
        px_win = prepare_window(px_raw, min_coverage=0.95, max_ffill=3)
    )

    print(f"Backtest NAV: {nav.index.min().date()} -> {nav.index.max().date()}")
    print(weights.tail(3))

    plt.figure()
    plt.plot(nav.index, nav.values)
    plt.title("Walk-forward NAV (EUR base)")
    plt.xlabel("Date")
    plt.ylabel("NAV (start=1)")
    plt.show()

    to = turnover(weights)
    print("\nTurnover stats (sum abs weight change per rebalance):")
    print(to.describe())

    bh = bound_hit_rate(weights, universe, region_bounds)
    if bh is not None:
        print("\nBound hit rates:")
        print(bh)

    # Benchmark + beta report
    asset_rets = simple_returns_from_prices(eur_px).dropna(how="any")
    port_rets = nav.pct_change().dropna()
    bench_rets = make_custom_benchmark(asset_rets, bench_w)

    df_rb = pd.concat([port_rets.rename("PORT"), bench_rets.rename("BENCH")], axis=1, sort=False).dropna()
    port_rets_al = df_rb["PORT"]
    bench_rets_al = df_rb["BENCH"]

    p_ret, p_vol, p_sh = ann_return_vol_sharpe(port_rets_al, rf=0.0)
    b_ret, b_vol, b_sh = ann_return_vol_sharpe(bench_rets_al, rf=0.0)

    b_full = beta(port_rets_al, bench_rets_al)
    b_roll = rolling_beta(port_rets_al, bench_rets_al, window=252)

    te, corr = tracking_error(port_rets_al, bench_rets_al)

    facts = pd.DataFrame({
        "ann_return": [p_ret, b_ret],
        "ann_vol":    [p_vol, b_vol],
        "sharpe_rf0": [p_sh,  b_sh],
        "beta_vs_bench":[b_full, 1.0],
        "tracking_error":[te, np.nan],
        "corr_vs_bench":[corr, np.nan],
    }, index=["PORTFOLIO", "BENCHMARK"])

    print("\nPerformance table:")
    print(facts)

    plt.figure()
    plt.plot(b_roll.index, b_roll.values)
    plt.title("Rolling 252d beta: Portfolio vs Custom Benchmark")
    plt.xlabel("Date")
    plt.ylabel("Beta")
    plt.show()

    # Mini Monte Carlo runs (now includes rails)
    experiments = [
        {"label":"trusted_grid", "lambda_mode":"grid", "cost_bps":5.0, "tactical_em":False,
         "max_turnover":0.10, "min_rebalance_l1":0.03, "vol_target":0.12, "cash_floor_risk_off":0.20},

        {"label":"trusted_quarterly", "rebalance":"QE", "lambda_mode":"grid", "cost_bps":5.0,
         "max_turnover":0.10, "min_rebalance_l1":0.03, "vol_target":0.12, "cash_floor_risk_off":0.20},

        {"label":"tighter_turnover", "lambda_mode":"grid", "cost_bps":5.0,
         "max_turnover":0.05, "min_rebalance_l1":0.03, "vol_target":0.12, "cash_floor_risk_off":0.20},

        {"label":"lower_vol", "lambda_mode":"grid", "cost_bps":5.0,
         "max_turnover":0.10, "min_rebalance_l1":0.03, "vol_target":0.10, "cash_floor_risk_off":0.20},
    ]

    summary = run_mini_monte_carlo(
        eur_px, universe,
        region_bounds, caps_by_kind, max_weight_by_kind, min_weight_by_kind,
        bench_w,
        experiments,
        rf=0.0,
        verbose=True
    )

    print("\n=== MINI MONTE CARLO SUMMARY (sorted by Sharpe) ===")
    print(summary.to_string(index=False))