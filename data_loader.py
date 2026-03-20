import pandas as pd  
import numpy as np  
import requests  
import os  
import time  
from datetime import datetime, timedelta  

CACHE_DIR = "cache_prices"  
os.makedirs(CACHE_DIR, exist_ok=True)  

# ── Fetch depuis stooq ──────────────────────────────────────────────────────  

def fetch_stooq(ticker: str,   
                start: str = "2018-01-01",   
                end: str = None) -> pd.Series:  
    """  
    Télécharge les prix de clôture depuis stooq.com.  
    Retourne une pd.Series avec index DatetimeIndex.  
    """  
    if end is None:  
        end = datetime.today().strftime("%Y%m%d")  
    start_fmt = start.replace("-", "")  
    end_fmt   = end.replace("-", "")  

    url = (  
        f"https://stooq.com/q/d/l/"  
        f"?s={ticker}&d1={start_fmt}&d2={end_fmt}&i=d"  
    )  

    try:  
        df = pd.read_csv(url, parse_dates=["Date"], index_col="Date")  
        if df.empty or "Close" not in df.columns:  
            return pd.Series(dtype=float, name=ticker)  
        series = df["Close"].sort_index().dropna()  
        series.name = ticker  
        return series  
    except Exception as e:  
        print(f"[stooq] Erreur pour {ticker}: {e}")  
        return pd.Series(dtype=float, name=ticker)  

# ── Cache local ─────────────────────────────────────────────────────────────  

def _cache_path(ticker: str) -> str:  
    safe = ticker.replace("/", "_").replace(".", "_")  
    return os.path.join(CACHE_DIR, f"{safe}.parquet")  

def load_with_cache(ticker: str,  
                    start: str = "2018-01-01",  
                    max_age_hours: int = 12) -> pd.Series:  
    """  
    Charge depuis le cache si < max_age_hours,  
    sinon re-télécharge depuis stooq.  
    """  
    path = _cache_path(ticker)  

    if os.path.exists(path):  
        age = time.time() - os.path.getmtime(path)  
        if age < max_age_hours * 3600:  
            try:  
                df = pd.read_parquet(path)  
                return df["Close"]  
            except Exception:  
                pass  

    # Téléchargement frais  
    series = fetch_stooq(ticker, start=start)  
    if not series.empty:  
        pd.DataFrame({"Close": series}).to_parquet(path)  
    return series  

# ── Chargement de la matrice de prix ────────────────────────────────────────  

def build_price_matrix(universe: list,  
                       start: str = "2018-01-01",  
                       delay: float = 0.3) -> pd.DataFrame:  
    """  
    Construit la matrice de prix pour tout l'univers.  
    Paramètres:  
        universe  : liste de dicts {"name", "ticker", ...}  
        start     : date de début  
        delay     : délai entre requêtes (éviter ban stooq)  
    Retourne:  
        DataFrame avec les noms comme colonnes, index DatetimeIndex  
    """  
    prices = {}  

    for asset in universe:  
        name   = asset["name"]  
        ticker = asset["ticker"]  

        if ticker is None:  # CASH  
            prices[name] = None  
            continue  

        print(f"[loader] Chargement {name} ({ticker})...")  
        series = load_with_cache(ticker, start=start)  

        if series.empty:  
            print(f"[loader] ⚠️  Pas de données pour {name}")  
        else:  
            prices[name] = series  

        time.sleep(delay)  

    # Assemblage  
    valid = {k: v for k, v in prices.items() if v is not None and not v.empty}  
    if not valid:  
        return pd.DataFrame()  

    df = pd.DataFrame(valid)  
    df.index = pd.to_datetime(df.index)  
    df = df.sort_index().dropna(how="all")  
    return df  