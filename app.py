"""
IKE Portfolio Analyzer — XTB Edition
=====================================
Aplikacja Streamlit do analizy portfela inwestycyjnego IKE Maklerskie (XTB):
- parsowanie eksportu XLSX (historia rachunku) z XTB i mapowanie tickerów na format Yahoo Finance,
- pobieranie danych rynkowych na żywo (yfinance) wraz z prognozami analityków,
- autorski silnik "Smart Buy-The-Dip",
- kalkulator rebalansu (klasyczny oraz przez nową wpłatę),
- moduł limitu IKE, podatku Belki i projekcji długoterminowej.

Uwaga: aplikacja ma charakter analityczno-edukacyjny i NIE stanowi porady
inwestycyjnej, podatkowej ani rekomendacji w rozumieniu przepisów o obrocie
instrumentami finansowymi.
"""

from __future__ import annotations

import io
import math
import logging
import re
import unicodedata
from datetime import datetime
from typing import Optional

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import yfinance as yf

# pandas_ta bywa niekompatybilne z nowszymi wersjami numpy (numpy>=2.0 usunęło
# alias `numpy.NaN`, z którego korzysta pandas_ta<=0.3.14b0). Importujemy je
# w trybie best-effort — jeśli się nie uda, aplikacja korzysta z własnych,
# czystych implementacji RSI/SMA i działa dalej bez przerywania.
try:
    import pandas_ta as ta  # noqa: F401
    PANDAS_TA_AVAILABLE = True
except Exception:  # pragma: no cover
    PANDAS_TA_AVAILABLE = False

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("ike_portfolio_app")

st.set_page_config(
    page_title="IKE Portfolio Analyzer — XTB",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================================
# KONFIGURACJA
# ============================================================================

BELKA_TAX_RATE = 0.19

# Roczne limity wpłat na IKE (źródło: obwieszczenia Ministra Rodziny, Pracy
# i Polityki Społecznej publikowane w Monitorze Polskim / gov.pl).
# Zaktualizuj ten słownik, gdy zostanie ogłoszony nowy limit na kolejny rok.
IKE_LIMITS = {
    2022: 17766.0,
    2023: 20805.0,
    2024: 23472.0,
    2025: 26019.0,
    2026: 28260.0,
}

# Aliasy nazw kolumn w eksporcie XLSX z XTB (arkusze "Open Positions" /
# "Closed Positions"; dopasowywane po znormalizowaniu: małe litery, bez
# polskich znaków diakrytycznych). Kolejność ma znaczenie — bardziej precyzyjne
# nazwy są sprawdzane w pierwszej kolejności.
SYMBOL_COLS = ["symbol", "ticker", "instrument"]
TYPE_COLS = ["typ", "rodzaj transakcji", "type", "side", "kierunek"]
VOLUME_COLS = ["wolumen", "ilosc", "volume", "quantity", "qty", "shares"]
PRICE_COLS = ["cena otwarcia", "open price", "cena zakupu", "price", "cena"]
AMOUNT_COLS = ["kwota", "wartosc", "amount", "purchase value", "total"]
DATE_COLS = ["data otwarcia", "open time", "czas otwarcia", "data", "date"]

# Arkusze eksportu historii rachunku XTB, po jakich fragmentach nazwy (po
# normalizacji) je rozpoznajemy.
CASH_OPERATIONS_SHEET_KEYWORDS = ["cash", "operations"]
CASH_OPERATIONS_SHEET_KEYWORDS = ["cash", "operations"]

# Dodatkowe aliasy kolumn potrzebne wyłącznie w arkuszu "Cash Operations".
COMMENT_COLS = ["comment", "komentarz"]
CASH_TIME_COLS = ["time", "data", "date"]

# Słowa kluczowe używane do zlokalizowania wiersza nagłówka tabeli wewnątrz
# arkusza XLSX (przed właściwą tabelą XTB umieszcza kilka wierszy metadanych
# takich jak numer konta czy zakres dat).
HEADER_ROW_KEYWORDS = ["ticker", "instrument", "symbol"]

SELL_KEYWORDS = ["sell", "sprzedaz"]

# "Cash Operations" nie zawiera osobnych kolumn Wolumen/Cena — trzeba je
# wyłuskać z pola Comment, np. "OPEN BUY 4 @ 78.010" lub, przy częściowym
# zamknięciu pozycji, "CLOSE BUY 3/4 @ 659.00" (wolumen = licznik ułamka).
TRANSACTION_COMMENT_RE = re.compile(
    r"(open|close)\s+(buy|sell)\s+([\d]+(?:[.,]\d+)?)(?:\s*/\s*[\d]+(?:[.,]\d+)?)?\s*@\s*([\d]+(?:[.,]\d+)?)",
    re.IGNORECASE,
)

# Mapowanie sufiksów giełd XTB -> sufiksów używanych przez Yahoo Finance.
# "" oznacza brak sufiksu (np. akcje US: AAPL.US -> AAPL).
XTB_TO_YAHOO_SUFFIX = {
    "US": "",
    "PL": "WA",
    "DE": "DE",
    "UK": "L",
    "FR": "PA",
    "NL": "AS",
    "ES": "MC",
    "IT": "MI",
    "PT": "LS",
    "AT": "VI",
    "CH": "SW",
    "BE": "BR",
    "SE": "ST",
    "NO": "OL",
    "DK": "CO",
    "FI": "HE",
    "HU": "BD",
    "CZ": "PR",
}

HISTORY_PERIOD_OPTIONS = ["6mo", "1y", "2y", "5y"]

# ============================================================================
# CSS — MOTYW "TERMINAL" (DARK MODE)
# ============================================================================


def inject_custom_css() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

        .stApp {
            background: radial-gradient(circle at 15% 0%, #10192E 0%, #0B1120 55%) !important;
            color: #E6EDF7;
        }

        h1, h2, h3 {
            font-family: 'Space Grotesk', sans-serif !important;
            letter-spacing: -0.01em;
        }

        /* Wartości liczbowe w stylu "ticker tape" terminala giełdowego */
        [data-testid="stMetricValue"] {
            font-family: 'IBM Plex Mono', monospace !important;
            font-weight: 600;
        }
        [data-testid="stMetricLabel"] {
            color: #8B96AC !important;
        }
        [data-testid="stMetric"] {
            background: #131B2E;
            border: 1px solid #232F49;
            border-left: 3px solid #2DD4BF;
            border-radius: 8px;
            padding: 14px 16px;
        }

        div[data-testid="stDataFrame"] * {
            font-family: 'IBM Plex Mono', monospace !important;
            font-size: 0.85rem;
        }

        .stTabs [data-baseweb="tab-list"] {
            gap: 4px;
            border-bottom: 1px solid #232F49;
        }
        .stTabs [data-baseweb="tab"] {
            background-color: transparent;
            font-family: 'Space Grotesk', sans-serif;
            color: #8B96AC;
        }
        .stTabs [aria-selected="true"] {
            color: #2DD4BF !important;
            border-bottom: 2px solid #2DD4BF !important;
        }

        section[data-testid="stSidebar"] {
            background-color: #0D1424;
            border-right: 1px solid #232F49;
        }

        .stButton button {
            background-color: #16233F;
            border: 1px solid #2DD4BF;
            color: #E6EDF7;
            border-radius: 6px;
        }
        .stButton button:hover {
            background-color: #1B2C4D;
            border-color: #2DD4BF;
            color: #2DD4BF;
        }

        div[data-testid="stExpander"] {
            border: 1px solid #232F49;
            border-radius: 8px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _apply_dark_theme(fig: go.Figure) -> go.Figure:
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="IBM Plex Mono, monospace", color="#E6EDF7"),
        margin=dict(l=10, r=10, t=50, b=10),
        legend=dict(bgcolor="rgba(0,0,0,0)"),
    )
    return fig


def _color_scale(value, vmin: float, vmax: float, reverse: bool = False) -> str:
    """Lekki, samodzielny odpowiednik background_gradient (bez zależności od
    matplotlib) — czerwony/żółty/zielony gradient dla wartości liczbowych."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    v = max(vmin, min(vmax, value))
    ratio = (v - vmin) / (vmax - vmin) if vmax > vmin else 0.5
    if reverse:
        ratio = 1 - ratio
    if ratio < 0.5:
        t = ratio / 0.5
        r, g, b = 239 + (250 - 239) * t, 68 + (204 - 68) * t, 68 + (21 - 68) * t
    else:
        t = (ratio - 0.5) / 0.5
        r, g, b = 250 + (34 - 250) * t, 204 + (197 - 204) * t, 21 + (94 - 21) * t
    return f"background-color: rgba({int(r)},{int(g)},{int(b)},0.30); color: #F5F7FA;"


def _styler_apply(styler, func, subset=None):
    """Kompatybilne z różnymi wersjami pandas (Styler.map vs applymap)."""
    if hasattr(styler, "map"):
        return styler.map(func, subset=subset)
    return styler.applymap(func, subset=subset)  # starsze pandas


def fmt_number(value, decimals: int = 2, suffix: str = "") -> str:
    """Formatuje liczbę w polskim stylu (spacja jako separator tysięcy,
    przecinek jako separator dziesiętny). Zwraca '—' dla braku danych."""
    if value is None:
        return "—"
    if isinstance(value, float) and math.isnan(value):
        return "—"
    try:
        text = f"{value:,.{decimals}f}"
        text = text.replace(",", " ").replace(".", ",")
        return f"{text}{suffix}"
    except (TypeError, ValueError):
        return "—"


# ============================================================================
# MODUŁ 1: PARSOWANIE CSV I MAPOWANIE SYMBOLI (XTB -> YAHOO FINANCE)
# ============================================================================


def _normalize(text: str) -> str:
    text = str(text).strip().lower()
    text = "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))
    return text


def _find_sheet(sheet_map: dict, keywords: list) -> Optional[str]:
    """Znajduje nazwę arkusza (oryginalną), której znormalizowana wersja
    zawiera wszystkie podane słowa kluczowe."""
    for norm_name, original in sheet_map.items():
        if all(k in norm_name for k in keywords):
            return original
    return None


def _read_sheet_table(xls: "pd.ExcelFile", sheet_name: str) -> Optional[pd.DataFrame]:
    """Wczytuje arkusz XLSX i lokalizuje właściwy wiersz nagłówka tabeli —
    eksporty XTB zawierają na górze arkusza kilka wierszy metadanych (numer
    konta, zakres dat), zanim pojawi się właściwa tabela z kolumnami takimi
    jak 'Ticker', 'Volume' itd."""
    raw = xls.parse(sheet_name, header=None)
    header_idx = None
    for idx in range(min(20, len(raw))):
        row_vals = [v for v in raw.iloc[idx].tolist() if pd.notna(v)]
        if len(row_vals) < 2:
            continue
        low_joined = " ".join(_normalize(v) for v in row_vals)
        if any(k in low_joined for k in HEADER_ROW_KEYWORDS):
            header_idx = idx
            break
    if header_idx is None:
        return None

    header = raw.iloc[header_idx].tolist()
    data = raw.iloc[header_idx + 1:].copy()
    data.columns = [str(c).strip() if pd.notna(c) else f"col_{i}" for i, c in enumerate(header)]
    data = data.dropna(axis=1, how="all").dropna(how="all")
    return data.reset_index(drop=True)


def find_column(columns, aliases) -> Optional[str]:
    norm_map = {_normalize(c): c for c in columns}
    for alias in aliases:
        if alias in norm_map:
            return norm_map[alias]
    for norm_col, original in norm_map.items():
        for alias in aliases:
            if alias in norm_col:
                return original
    return None


def map_xtb_symbol_to_yahoo(xtb_symbol: str) -> str:
    """Konwertuje ticker w formacie XTB (np. 'PKN.PL', 'AAPL.US', 'EUNL.DE')
    na format rozpoznawany przez Yahoo Finance."""
    if not isinstance(xtb_symbol, str) or not xtb_symbol.strip():
        return ""
    symbol = xtb_symbol.strip().upper()

    if "." in symbol:
        base, suffix = symbol.rsplit(".", 1)
        yahoo_suffix = XTB_TO_YAHOO_SUFFIX.get(suffix)
        if yahoo_suffix is None:
            logger.warning("Nieznany sufiks giełdy '.%s' dla %s — pozostawiono bez zmian.", suffix, xtb_symbol)
            return symbol
        return base if yahoo_suffix == "" else f"{base}.{yahoo_suffix}"
    return symbol


def _parse_transaction_comment(comment) -> Optional[tuple]:
    """Wyłuskuje z pola 'Comment' arkusza Cash Operations informacje o
    transakcji, np. 'OPEN BUY 4 @ 78.010' -> ('open', 4.0, 78.010) lub
    'CLOSE BUY 3/4 @ 659.00' -> ('close', 3.0, 659.00) — przy częściowym
    zamknięciu pozycji liczy się tylko licznik ułamka (wolumen tej
    konkretnej transakcji zamknięcia)."""
    if not isinstance(comment, str):
        return None
    match = TRANSACTION_COMMENT_RE.search(comment)
    if not match:
        return None
    action = match.group(1).lower()
    try:
        volume = float(match.group(3).replace(",", "."))
        price = float(match.group(4).replace(",", "."))
    except ValueError:
        return None
    return action, volume, price


def _parse_cash_operations(df: pd.DataFrame) -> pd.DataFrame:
    """Buduje znormalizowaną tabelę transakcji z arkusza 'Cash Operations',
    który zawiera KAŻDĄ pojedynczą operację na koncie (zakupy, sprzedaże,
    dywidendy, wpłaty, odsetki...). Filtrujemy tylko wiersze będące realnymi
    transakcjami giełdowymi ('Stock purchase' / 'Stock sell'), a wolumen i
    cenę odczytujemy z pola Comment, bo arkusz nie ma osobnych kolumn na te
    wartości."""
    col_type = find_column(df.columns, TYPE_COLS)
    col_symbol = find_column(df.columns, SYMBOL_COLS)
    col_amount = find_column(df.columns, AMOUNT_COLS)
    col_comment = find_column(df.columns, COMMENT_COLS)
    col_date = find_column(df.columns, CASH_TIME_COLS)

    if col_type is None or col_symbol is None or col_comment is None:
        return pd.DataFrame()

    work = pd.DataFrame()
    work["type_raw"] = df[col_type].astype(str)
    work["symbol_xtb"] = df[col_symbol].astype(str).str.strip()
    work["amount_raw"] = pd.to_numeric(df[col_amount], errors="coerce") if col_amount else np.nan
    work["comment"] = df[col_comment].astype(str) if col_comment else ""
    work["date"] = pd.to_datetime(df[col_date], errors="coerce") if col_date else pd.NaT

    # Zostają wyłącznie wiersze opisujące transakcje giełdowe (kupno/sprzedaż
    # akcji/ETF) — odrzucamy dywidendy, wpłaty, odsetki od wolnych środków itp.
    type_norm = work["type_raw"].apply(_normalize)
    is_stock_row = type_norm.str.contains("stock") & (
        type_norm.str.contains("purchase") | type_norm.str.contains("sell") | type_norm.str.contains("sprzedaz")
    )
    work = work[is_stock_row & (work["symbol_xtb"].str.len() > 0)]
    work = work[~work["symbol_xtb"].str.lower().isin(["nan", "none", ""])]

    parsed = work["comment"].apply(_parse_transaction_comment)
    work = work[parsed.notna()].copy()
    if work.empty:
        return pd.DataFrame()
    parsed = parsed[parsed.notna()]

    work["action"] = parsed.apply(lambda p: p[0])
    work["volume"] = parsed.apply(lambda p: p[1])
    work["price"] = parsed.apply(lambda p: p[2])
    work["amount"] = work["amount_raw"].abs()
    missing_amount = work["amount"].isna()
    work.loc[missing_amount, "amount"] = work.loc[missing_amount, "volume"] * work.loc[missing_amount, "price"]
    # "OPEN" = wejście w pozycję (kupno), "CLOSE" = wyjście z pozycji
    # (sprzedaż) — niezależnie od słowa BUY/SELL w komentarzu, bo ono opisuje
    # kierunek pierwotnej pozycji, a nie tej konkretnej operacji.
    work["type"] = np.where(work["action"] == "open", "BUY", "SELL")

    return work[["symbol_xtb", "volume", "price", "amount", "type", "date"]].reset_index(drop=True)


@st.cache_data(show_spinner=False)
def parse_xtb_xlsx(file_bytes: bytes) -> pd.DataFrame:
    """Parsuje eksport XLSX historii rachunku z XTB do znormalizowanej tabeli
    transakcji, korzystając z arkusza 'Cash Operations' — zawiera on komplet
    pojedynczych transakcji (i zakupy, i sprzedaże), z czego po zagregowaniu
    (`aggregate_positions`) wynika aktualny, otwarty portfel."""
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    except Exception as exc:
        raise ValueError(f"Nie udało się odczytać pliku XLSX ({exc}). Sprawdź, czy to poprawny eksport z XTB.")

    sheet_map = {_normalize(name): name for name in xls.sheet_names}
    cash_sheet = _find_sheet(sheet_map, CASH_OPERATIONS_SHEET_KEYWORDS)

    if cash_sheet is None:
        raise ValueError(
            "Nie znaleziono arkusza 'Cash Operations' w pliku. "
            "Sprawdź, czy to poprawny eksport historii rachunku z XTB "
            "(w platformie XTB: Historia rachunku -> Eksport -> XLSX)."
        )

    df = _read_sheet_table(xls, cash_sheet)
    if df is None or df.empty:
        raise ValueError("Arkusz 'Cash Operations' jest pusty lub nie udało się odczytać jego nagłówka.")

    clean = _parse_cash_operations(df)
    if clean.empty:
        raise ValueError(
            "Nie znaleziono żadnych transakcji giełdowych ('Stock purchase' / 'Stock sell') "
            "w arkuszu 'Cash Operations'. Sprawdź strukturę pliku XLSX."
        )

    clean["type_norm"] = clean["type"].apply(_normalize)
    clean["is_sell"] = clean["type_norm"].apply(lambda t: any(k in t for k in SELL_KEYWORDS))
    clean["symbol_yahoo"] = clean["symbol_xtb"].apply(map_xtb_symbol_to_yahoo)

    return clean.reset_index(drop=True)


def aggregate_positions(clean_df: pd.DataFrame) -> pd.DataFrame:
    """Agreguje transakcje do aktualnie otwartych pozycji: wolumen netto,
    średnia cena zakupu (metoda average cost) i koszt zainwestowany w
    POZOSTAŁY (jeszcze niesprzedany) wolumen per ticker."""
    if clean_df.empty:
        return pd.DataFrame(columns=["symbol_xtb", "symbol_yahoo", "volume", "avg_price", "total_spent"])

    buys = clean_df[~clean_df["is_sell"]]
    sells = clean_df[clean_df["is_sell"]]

    buy_agg = buys.groupby(["symbol_xtb", "symbol_yahoo"], as_index=False).agg(
        buy_volume=("volume", "sum"), total_bought=("amount", "sum")
    )
    sell_agg = sells.groupby(["symbol_xtb", "symbol_yahoo"], as_index=False).agg(sell_volume=("volume", "sum"))

    positions = buy_agg.merge(sell_agg, on=["symbol_xtb", "symbol_yahoo"], how="left")
    positions["sell_volume"] = positions["sell_volume"].fillna(0.0)
    positions["volume"] = positions["buy_volume"] - positions["sell_volume"]
    positions["avg_price"] = np.where(
        positions["buy_volume"] > 0, positions["total_bought"] / positions["buy_volume"], np.nan
    )

    # WAŻNE: "total_bought" to koszt WSZYSTKICH historycznych zakupów danego
    # tickera — jeśli część z nich już sprzedano (częściowe zamknięcie
    # pozycji), użycie tej sumy jako "zainwestowano" zaniżałoby wynik
    # procentowy zysku/straty. Koszt zainwestowany w pozostały wolumen
    # liczymy metodą average cost: średnia cena zakupu * wolumen wciąż
    # posiadany.
    positions["total_spent"] = positions["avg_price"] * positions["volume"]

    positions = positions[positions["volume"] > 1e-9].reset_index(drop=True)
    return positions.drop(columns=["buy_volume", "sell_volume", "total_bought"])


# ============================================================================
# MODUŁ 2: DANE RYNKOWE NA ŻYWO (yfinance) + WSKAŹNIKI TECHNICZNE
# ============================================================================


def compute_rsi(close: pd.Series, length: int = 14) -> pd.Series:
    if PANDAS_TA_AVAILABLE:
        try:
            result = ta.rsi(close, length=length)
            if result is not None:
                return result
        except Exception as exc:  # pragma: no cover
            logger.warning("pandas_ta.rsi nie powiodło się, użyto fallbacku: %s", exc)
    delta = close.diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.ewm(alpha=1 / length, min_periods=length, adjust=False).mean()
    avg_loss = loss.ewm(alpha=1 / length, min_periods=length, adjust=False).mean()
    rs = avg_gain / avg_loss.replace(0, np.nan)
    return 100 - (100 / (1 + rs))


def compute_sma(close: pd.Series, length: int) -> pd.Series:
    if PANDAS_TA_AVAILABLE:
        try:
            result = ta.sma(close, length=length)
            if result is not None:
                return result
        except Exception as exc:  # pragma: no cover
            logger.warning("pandas_ta.sma nie powiodło się, użyto fallbacku: %s", exc)
    return close.rolling(window=length, min_periods=max(1, length // 2)).mean()


def _safe_float(value) -> float:
    try:
        return np.nan if value is None else float(value)
    except (TypeError, ValueError):
        return np.nan


@st.cache_data(show_spinner=False, ttl=3600)
def fetch_market_data(symbol: str, period: str = "2y") -> dict:
    """Pobiera cenę, wskaźniki techniczne i dane fundamentalne dla jednego
    tickera Yahoo Finance. Zawsze zwraca słownik — błędy trafiają do klucza
    'error' zamiast przerywać działanie aplikacji."""
    result = {
        "symbol": symbol, "current_price": np.nan, "currency": "",
        "target_mean": np.nan, "target_high": np.nan, "target_low": np.nan,
        "rsi14": np.nan, "sma50": np.nan, "sma200": np.nan,
        "week52_high": np.nan, "week52_low": np.nan, "dividend_yield": np.nan,
        "sector": "Nieznany", "error": None,
    }
    if not symbol:
        result["error"] = "Brak symbolu."
        return result

    try:
        ticker = yf.Ticker(symbol)
        hist = ticker.history(period=period, auto_adjust=True)
        if hist is None or hist.empty:
            result["error"] = f"Brak danych historycznych dla {symbol}."
            return result

        close = hist["Close"].dropna()
        result["current_price"] = float(close.iloc[-1]) if len(close) else np.nan

        rsi_series = compute_rsi(close, 14).dropna()
        result["rsi14"] = float(rsi_series.iloc[-1]) if len(rsi_series) else np.nan

        sma50_series = compute_sma(close, 50).dropna()
        result["sma50"] = float(sma50_series.iloc[-1]) if len(sma50_series) else np.nan

        sma200_series = compute_sma(close, 200).dropna()
        result["sma200"] = float(sma200_series.iloc[-1]) if len(sma200_series) else np.nan

        lookback = close.tail(252)
        result["week52_high"] = float(lookback.max()) if len(lookback) else np.nan
        result["week52_low"] = float(lookback.min()) if len(lookback) else np.nan

        try:
            info = ticker.info or {}
        except Exception as exc:
            logger.warning("ticker.info nie powiodło się dla %s: %s", symbol, exc)
            info = {}

        result["currency"] = info.get("currency") or ""
        result["target_mean"] = _safe_float(info.get("targetMeanPrice"))
        result["target_high"] = _safe_float(info.get("targetHighPrice"))
        result["target_low"] = _safe_float(info.get("targetLowPrice"))
        result["sector"] = info.get("sector") or "Nieznany"

        div_yield = _safe_float(info.get("dividendYield"))
        if not math.isnan(div_yield) and div_yield > 1:
            div_yield = div_yield / 100.0  # yfinance bywa niespójne: ułamek vs procent
        result["dividend_yield"] = div_yield

        if math.isnan(result["week52_high"]):
            result["week52_high"] = _safe_float(info.get("fiftyTwoWeekHigh"))
        if math.isnan(result["week52_low"]):
            result["week52_low"] = _safe_float(info.get("fiftyTwoWeekLow"))

    except Exception as exc:
        logger.error("Błąd pobierania danych dla %s: %s", symbol, exc)
        result["error"] = str(exc)

    return result


# ============================================================================
# MODUŁ 3: SILNIK "SMART BUY-THE-DIP"
# ============================================================================


def compute_dip_score(rsi: float, price: float, sma200: float, week52_high: float) -> dict:
    """Autorski wskaźnik heurystyczny 0-100 na podstawie trzech sygnałów:
    RSI<35, cena poniżej SMA200, spadek >=15% od 52-tygodniowego szczytu.
    To narzędzie analityczne, NIE rekomendacja inwestycyjna."""
    score = 0.0
    reasons = []

    if not math.isnan(rsi) and rsi < 35:
        rsi_pts = min(40.0, (35 - rsi) / 35 * 40 + 10)
        score += rsi_pts
        reasons.append(f"RSI={rsi:.1f} (wyprzedanie)")

    if not math.isnan(price) and not math.isnan(sma200) and sma200 > 0 and price < sma200:
        distance_pct = (sma200 - price) / sma200 * 100
        score += min(30.0, 15 + distance_pct)
        reasons.append(f"cena {distance_pct:.1f}% poniżej SMA200")

    drop_from_high_pct = np.nan
    if not math.isnan(price) and not math.isnan(week52_high) and week52_high > 0:
        drop_from_high_pct = (week52_high - price) / week52_high * 100
        if drop_from_high_pct >= 15:
            score += min(30.0, drop_from_high_pct)
            reasons.append(f"spadek {drop_from_high_pct:.1f}% od 52W High")

    score = float(min(100.0, max(0.0, score)))
    if score >= 60:
        label = "🔥 Głęboki dołek"
    elif score >= 35:
        label = "🟡 Możliwa okazja"
    else:
        label = "⚪ Brak sygnału"

    return {
        "dip_score": round(score, 1),
        "label": label,
        "reasons": "; ".join(reasons) if reasons else "brak spełnionych warunków",
        "drop_from_high_pct": drop_from_high_pct,
    }


# ============================================================================
# MODUŁ 4: KALKULATOR REBALANSU
# ============================================================================


def classic_rebalance(df: pd.DataFrame, target_col: str = "target_pct") -> pd.DataFrame:
    """Klasyczny rebalans: sprzedaj nadwyżki, dokup niedowagi, aż osiągnąć cel."""
    df = df.copy()
    total_value = df["current_value"].sum()
    df["current_pct"] = np.where(total_value > 0, df["current_value"] / total_value * 100, 0)
    df["target_value"] = total_value * df[target_col] / 100.0
    df["diff_value"] = df["target_value"] - df["current_value"]
    df["action"] = np.where(df["diff_value"] > 1e-6, "KUP", np.where(df["diff_value"] < -1e-6, "SPRZEDAJ", "OK"))
    return df


def contribution_rebalance(df: pd.DataFrame, contribution_amount: float, target_col: str = "target_pct") -> pd.DataFrame:
    """Rozdziela nową wpłatę proporcjonalnie do tego, jak bardzo dany składnik
    jest niedoważony względem celu — bez sprzedawania istniejących pozycji.

    1) policz deficyt = max(0, cel - obecna_wartość) przy wartości portfela
       powiększonej o wpłatę,
    2) rozdziel wpłatę proporcjonalnie do deficytu,
    3) jeśli wpłata przewyższa sumę deficytów, nadwyżkę rozdziel wg wag celu.
    """
    df = df.copy()
    current_total = df["current_value"].sum()
    total_after = current_total + contribution_amount

    df["target_value_after"] = total_after * df[target_col] / 100.0
    df["deficit"] = (df["target_value_after"] - df["current_value"]).clip(lower=0)
    total_deficit = df["deficit"].sum()

    if total_deficit <= 0:
        df["allocation"] = contribution_amount * df[target_col] / 100.0
    elif total_deficit >= contribution_amount:
        df["allocation"] = contribution_amount * (df["deficit"] / total_deficit)
    else:
        remainder = contribution_amount - total_deficit
        weight_sum = df[target_col].sum()
        if weight_sum > 0:
            remainder_alloc = remainder * (df[target_col] / weight_sum)
        else:
            remainder_alloc = remainder / len(df) if len(df) else 0.0
        df["allocation"] = df["deficit"] + remainder_alloc

    df["current_pct"] = np.where(current_total > 0, df["current_value"] / current_total * 100, 0)
    df["new_value"] = df["current_value"] + df["allocation"]
    df["new_pct"] = np.where(total_after > 0, df["new_value"] / total_after * 100, 0)
    return df


# ============================================================================
# MODUŁ 5: IKE, PODATEK BELKI, PROJEKCJA DŁUGOTERMINOWA
# ============================================================================


def get_ike_limit(year: int) -> float:
    if year in IKE_LIMITS:
        return IKE_LIMITS[year]
    return IKE_LIMITS[max(IKE_LIMITS.keys())]


def belka_tax_saved(profit: float) -> float:
    return max(0.0, profit) * BELKA_TAX_RATE


def project_capital(initial: float, annual_contribution: float, years: int, cagr: float) -> pd.DataFrame:
    """Uproszczona projekcja kapitału rok do roku (odsetki składane rocznie +
    stała roczna wpłata) — porównanie IKE (bez podatku Belki) vs konto
    opodatkowane (podatek 19% naliczany od wypracowanego zysku na koniec)."""
    rows = []
    capital_ike = initial
    capital_taxed = initial
    for year in range(1, years + 1):
        capital_ike = capital_ike * (1 + cagr) + annual_contribution
        capital_taxed = capital_taxed * (1 + cagr) + annual_contribution
        rows.append({"year": year, "ike_value": capital_ike, "taxed_value": capital_taxed})

    df = pd.DataFrame(rows)
    total_contributed = initial + annual_contribution * years
    df["taxed_profit"] = (df["taxed_value"] - total_contributed).clip(lower=0)
    df["taxed_value_net"] = total_contributed + (df["taxed_value"] - total_contributed) - df["taxed_profit"] * BELKA_TAX_RATE
    df["ike_advantage"] = df["ike_value"] - df["taxed_value_net"]
    return df


# ============================================================================
# UI: KOMPONENTY WSPÓLNE
# ============================================================================


def render_kpi_row(total_value, total_invested, profit_value, profit_pct, belka_saved, ike_contributed, ike_limit) -> None:
    ike_used_pct = (ike_contributed / ike_limit * 100) if ike_limit else 0.0

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("💼 Całkowita wartość", fmt_number(total_value, 2, " PLN"))
    with col2:
        delta = f"{profit_pct:+.2f}%" if not (isinstance(profit_pct, float) and math.isnan(profit_pct)) else None
        st.metric("📈 Zysk / Strata", fmt_number(profit_value, 2, " PLN"), delta=delta)
    with col3:
        st.metric(
            "🧾 Zaoszczędzony podatek Belki",
            fmt_number(belka_saved, 2, " PLN"),
            help="19% od niezrealizowanego zysku — kwota podatku, której unikasz dzięki opakowaniu IKE.",
        )
    with col4:
        st.metric(f"🎯 Limit IKE ({ike_used_pct:.0f}%)", fmt_number(ike_contributed, 0, " PLN"), delta=f"z {fmt_number(ike_limit, 0, ' PLN')}")

    st.progress(min(1.0, max(0.0, ike_used_pct / 100)))


# ============================================================================
# UI: ZAKŁADKA 1 — PRZEGLĄD PORTFELA
# ============================================================================


def render_overview_tab(portfolio: pd.DataFrame, raw_transactions: pd.DataFrame) -> None:
    st.subheader("Pozycje w portfelu")

    display_cols = {
        "symbol_xtb": "Symbol (XTB)", "symbol_yahoo": "Symbol (Yahoo)", "volume": "Wolumen",
        "avg_price": "Śr. cena zakupu", "current_price": "Cena aktualna", "total_spent": "Zainwestowano",
        "current_value": "Wartość obecna", "profit_value": "Zysk/Strata", "profit_pct": "Zysk/Strata %",
    }
    view = portfolio[list(display_cols.keys())].rename(columns=display_cols)
    styler = view.style.format(
        {
            "Wolumen": "{:.4f}", "Śr. cena zakupu": "{:.2f}", "Cena aktualna": "{:.2f}",
            "Zainwestowano": "{:.2f}", "Wartość obecna": "{:.2f}", "Zysk/Strata": "{:.2f}",
            "Zysk/Strata %": "{:+.2f}%",
        },
        na_rep="—",
    )
    styler = _styler_apply(styler, lambda v: _color_scale(v, -30, 30), subset=["Zysk/Strata %"])
    st.dataframe(styler, use_container_width=True)

    missing = portfolio[portfolio["current_price"].isna()]
    if not missing.empty:
        st.warning("⚠️ Brak aktualnych danych rynkowych dla: " + ", ".join(missing["symbol_yahoo"].tolist()))

    col1, col2 = st.columns(2)
    with col1:
        pie_df = portfolio.dropna(subset=["current_value"])
        if not pie_df.empty and pie_df["current_value"].sum() > 0:
            fig = px.pie(pie_df, names="symbol_xtb", values="current_value", title="Alokacja portfela", hole=0.5)
            st.plotly_chart(_apply_dark_theme(fig), use_container_width=True)
    with col2:
        bar_df = portfolio.dropna(subset=["profit_value"])
        if not bar_df.empty:
            fig2 = px.bar(
                bar_df.sort_values("profit_value"), x="profit_value", y="symbol_xtb", orientation="h",
                color="profit_value", color_continuous_scale=["#F87171", "#232F49", "#34D399"],
                title="Zysk / Strata per pozycja (PLN)", labels={"profit_value": "PLN", "symbol_xtb": ""},
            )
            st.plotly_chart(_apply_dark_theme(fig2), use_container_width=True)

    with st.expander("🔍 Podgląd rozpoznanych transakcji z pliku XLSX"):
        preview_cols = {
            "symbol_xtb": "Symbol (XTB)", "symbol_yahoo": "Symbol (Yahoo)", "type": "Typ",
            "volume": "Wolumen", "price": "Cena", "amount": "Kwota", "date": "Data",
        }
        available = [c for c in preview_cols if c in raw_transactions.columns]
        st.dataframe(raw_transactions[available].rename(columns=preview_cols), use_container_width=True)


# ============================================================================
# UI: ZAKŁADKA 2 — PROGNOZY ANALITYKÓW
# ============================================================================


def render_analyst_tab(portfolio: pd.DataFrame) -> None:
    st.subheader("🎯 Prognozy analityków (6-12 miesięcy)")

    df = portfolio.copy()
    df["upside_pct"] = np.where(
        df["current_price"].notna() & df["target_mean"].notna() & (df["current_price"] > 0),
        (df["target_mean"] - df["current_price"]) / df["current_price"] * 100,
        np.nan,
    )

    if not df["target_mean"].notna().any():
        st.info("Brak dostępnych prognoz analityków dla pozycji w portfelu (yfinance nie zwrócił tych danych — dotyczy to zwłaszcza wielu ETF-ów).")

    cols = {
        "symbol_xtb": "Symbol", "current_price": "Cena obecna", "target_low": "Cel min.",
        "target_mean": "Cel średni", "target_high": "Cel max.", "upside_pct": "Potencjał %",
    }
    view = df[list(cols.keys())].rename(columns=cols)
    styler = view.style.format(
        {"Cena obecna": "{:.2f}", "Cel min.": "{:.2f}", "Cel średni": "{:.2f}", "Cel max.": "{:.2f}", "Potencjał %": "{:+.2f}%"},
        na_rep="—",
    )
    styler = _styler_apply(styler, lambda v: _color_scale(v, -20, 20), subset=["Potencjał %"])
    st.dataframe(styler, use_container_width=True)

    chart_df = df.dropna(subset=["target_mean", "current_price"])
    if not chart_df.empty:
        fig = go.Figure()
        fig.add_trace(go.Bar(x=chart_df["symbol_xtb"], y=chart_df["current_price"], name="Cena obecna", marker_color="#2DD4BF"))
        fig.add_trace(go.Bar(x=chart_df["symbol_xtb"], y=chart_df["target_mean"], name="Cel średni (analitycy)", marker_color="#FBBF24"))
        fig.update_layout(barmode="group", title="Cena obecna vs cel analityków")
        st.plotly_chart(_apply_dark_theme(fig), use_container_width=True)

    st.caption("Prognozy pochodzą z agregacji konsensusu analityków dostępnej w Yahoo Finance i mają charakter wyłącznie informacyjny.")


# ============================================================================
# UI: ZAKŁADKA 3 — BUY THE DIP
# ============================================================================


def render_dip_tab(portfolio: pd.DataFrame) -> None:
    st.subheader("📉 Silnik Smart Buy-The-Dip")
    st.caption("Autorski wskaźnik heurystyczny (0–100) łączący RSI, SMA200 i dystans od 52W High — nie stanowi rekomendacji inwestycyjnej.")

    df = portfolio.sort_values("dip_score", ascending=False).copy()
    cols = {
        "symbol_xtb": "Symbol", "current_price": "Cena", "rsi14": "RSI(14)", "sma200": "SMA200",
        "week52_high": "52W High", "drop_from_high_pct": "Spadek od 52W High %", "dip_score": "Dip Score",
        "label": "Sygnał", "reasons": "Uzasadnienie",
    }
    view = df[list(cols.keys())].rename(columns=cols)
    styler = view.style.format(
        {
            "Cena": "{:.2f}", "RSI(14)": "{:.1f}", "SMA200": "{:.2f}", "52W High": "{:.2f}",
            "Spadek od 52W High %": "{:.1f}%", "Dip Score": "{:.0f}",
        },
        na_rep="—",
    )
    styler = _styler_apply(styler, lambda v: _color_scale(v, 0, 100, reverse=True), subset=["Dip Score"])
    st.dataframe(styler, use_container_width=True)

    deep_dip = df[df["dip_score"] >= 60]
    if not deep_dip.empty:
        st.error("🔥 Głęboki dołek wykryty w: " + ", ".join(deep_dip["symbol_xtb"].tolist()))

    chart_df = df.dropna(subset=["dip_score"])
    if not chart_df.empty:
        fig = px.bar(
            chart_df, x="symbol_xtb", y="dip_score", color="dip_score",
            color_continuous_scale=["#232F49", "#FBBF24", "#F87171"], range_color=[0, 100],
            title="Dip Score per aktywo", labels={"symbol_xtb": "", "dip_score": "Dip Score"},
        )
        fig.add_hline(y=60, line_dash="dash", line_color="#F87171", annotation_text="Głęboki dołek")
        st.plotly_chart(_apply_dark_theme(fig), use_container_width=True)


# ============================================================================
# UI: ZAKŁADKA 4 — KALKULATOR REBALANSU
# ============================================================================


def render_rebalance_tab(portfolio: pd.DataFrame) -> None:
    st.subheader("⚖️ Kalkulator Rebalansu")

    granularity = st.radio("Poziom rebalansu", ["Aktywa", "Sektory"], horizontal=True)
    mode = st.radio("Tryb kalkulatora", ["Rebalans klasyczny (sprzedaj / dokup)", "Rebalans wpłatą (nowe środki)"], horizontal=True)

    base = portfolio.dropna(subset=["current_value"]).copy()
    if base.empty:
        st.warning("Brak wycenionych pozycji do rebalansu — sprawdź dostępność danych rynkowych w zakładce Przegląd.")
        return

    if granularity == "Sektory":
        base["sector"] = base["sector"].fillna("Nieznany")
        grouped = base.groupby("sector", as_index=False)["current_value"].sum().rename(columns={"sector": "label"})
    else:
        grouped = base[["symbol_xtb", "current_value"]].rename(columns={"symbol_xtb": "label"})

    n = len(grouped)
    grouped["target_pct"] = round(100 / n, 2) if n else 0.0

    st.caption("Ustaw docelowe udziały procentowe (suma powinna wynosić 100%):")
    edited = st.data_editor(
        grouped,
        column_config={
            "label": st.column_config.TextColumn("Pozycja / Sektor", disabled=True),
            "current_value": st.column_config.NumberColumn("Wartość obecna (PLN)", format="%.2f", disabled=True),
            "target_pct": st.column_config.NumberColumn("Cel %", min_value=0.0, max_value=100.0, step=0.5),
        },
        hide_index=True,
        use_container_width=True,
        key=f"rebalance_editor_{granularity}",
    )

    target_sum = edited["target_pct"].sum()
    if abs(target_sum - 100) > 0.5:
        st.warning(f"⚠️ Suma docelowych udziałów wynosi {target_sum:.1f}% — dla poprawnych wyników powinna wynosić 100%.")

    if mode.startswith("Rebalans klasyczny"):
        result = classic_rebalance(edited.rename(columns={"label": "symbol"}))
        show = result.rename(
            columns={
                "symbol": "Pozycja", "current_value": "Wartość obecna", "current_pct": "Udział obecny %",
                "target_pct": "Udział docelowy %", "target_value": "Wartość docelowa",
                "diff_value": "Różnica (PLN)", "action": "Akcja",
            }
        )
        cols = ["Pozycja", "Wartość obecna", "Udział obecny %", "Udział docelowy %", "Wartość docelowa", "Różnica (PLN)", "Akcja"]
        st.dataframe(
            show[cols].style.format(
                {
                    "Wartość obecna": "{:.2f}", "Udział obecny %": "{:.1f}%", "Udział docelowy %": "{:.1f}%",
                    "Wartość docelowa": "{:.2f}", "Różnica (PLN)": "{:+.2f}",
                },
                na_rep="—",
            ),
            use_container_width=True,
        )
        fig = px.bar(
            result, x="symbol", y=["current_value", "target_value"], barmode="group",
            title="Obecna vs docelowa wartość pozycji", labels={"symbol": "", "value": "PLN"},
            color_discrete_sequence=["#2DD4BF", "#FBBF24"],
        )
        st.plotly_chart(_apply_dark_theme(fig), use_container_width=True)

    else:
        contribution = st.number_input("Kwota nowej wpłaty (PLN)", min_value=0.0, value=3000.0, step=100.0)
        if contribution > 0:
            result = contribution_rebalance(edited.rename(columns={"label": "symbol"}), contribution)
            show = result.rename(
                columns={
                    "symbol": "Pozycja", "current_value": "Wartość obecna", "current_pct": "Udział obecny %",
                    "allocation": "Kwota do dopłaty", "new_value": "Wartość po wpłacie", "new_pct": "Udział po wpłacie %",
                }
            )
            cols = ["Pozycja", "Wartość obecna", "Udział obecny %", "Kwota do dopłaty", "Wartość po wpłacie", "Udział po wpłacie %"]
            st.dataframe(
                show[cols].style.format(
                    {
                        "Wartość obecna": "{:.2f}", "Udział obecny %": "{:.1f}%", "Kwota do dopłaty": "{:.2f}",
                        "Wartość po wpłacie": "{:.2f}", "Udział po wpłacie %": "{:.1f}%",
                    },
                    na_rep="—",
                ),
                use_container_width=True,
            )
            fig = px.bar(
                result.sort_values("allocation", ascending=False), x="symbol", y="allocation",
                title=f"Rozdział wpłaty {fmt_number(contribution, 0, ' PLN')} wg odchylenia od celu",
                color="allocation", color_continuous_scale=["#131B2E", "#2DD4BF"],
                labels={"symbol": "", "allocation": "PLN"},
            )
            st.plotly_chart(_apply_dark_theme(fig), use_container_width=True)


# ============================================================================
# UI: ZAKŁADKA 5 — PROJEKCJA IKE I DYWIDENDY
# ============================================================================


def render_projection_tab(portfolio: pd.DataFrame, total_value: float, cagr_pct: float, ike_limit: float) -> None:
    st.subheader("💰 Projekcja kapitału: IKE vs konto opodatkowane")
    st.caption("Uproszczony model rocznego kapitalizowania odsetek — nie stanowi porady podatkowej ani inwestycyjnej.")

    cagr = cagr_pct / 100
    col1, col2 = st.columns(2)
    with col1:
        annual_contribution = st.number_input(
            "Zakładana roczna wpłata (PLN)", min_value=0.0, value=float(min(ike_limit, 10000.0)), step=500.0
        )
    with col2:
        st.metric("Kapitał startowy (obecna wartość portfela)", fmt_number(total_value, 2, " PLN"))

    tabs = st.tabs(["10 lat", "20 lat", "30 lat"])
    for horizon, tab in zip([10, 20, 30], tabs):
        with tab:
            proj = project_capital(total_value, annual_contribution, horizon, cagr)
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=proj["year"], y=proj["ike_value"], name="IKE (bez podatku Belki)", line=dict(color="#2DD4BF", width=3)))
            fig.add_trace(go.Scatter(x=proj["year"], y=proj["taxed_value_net"], name="Konto opodatkowane (netto)", line=dict(color="#F87171", width=3)))
            fig.update_layout(title=f"Projekcja {horizon} lat przy CAGR={cagr_pct}%", xaxis_title="Rok", yaxis_title="Wartość (PLN)")
            st.plotly_chart(_apply_dark_theme(fig), use_container_width=True)

            final = proj.iloc[-1]
            c1, c2, c3 = st.columns(3)
            c1.metric(f"IKE po {horizon} latach", fmt_number(final["ike_value"], 0, " PLN"))
            c2.metric("Konto opodatkowane (netto)", fmt_number(final["taxed_value_net"], 0, " PLN"))
            c3.metric("Przewaga IKE", fmt_number(final["ike_advantage"], 0, " PLN"))

    st.divider()
    st.subheader("Dywidendy")
    div_df = portfolio.dropna(subset=["dividend_yield"]).copy()
    if div_df.empty:
        st.info("Brak danych o stopie dywidendy dla pozycji w portfelu.")
    else:
        div_df["annual_dividend_income"] = div_df["current_value"] * div_df["dividend_yield"]
        total_div_income = div_df["annual_dividend_income"].sum()
        st.metric("Szacowany łączny roczny dochód z dywidend (brutto)", fmt_number(total_div_income, 2, " PLN"))

        show = div_df[["symbol_xtb", "dividend_yield", "current_value", "annual_dividend_income"]].rename(
            columns={
                "symbol_xtb": "Symbol", "dividend_yield": "Stopa dywidendy", "current_value": "Wartość pozycji",
                "annual_dividend_income": "Szacowany roczny dochód",
            }
        )
        st.dataframe(
            show.style.format({"Stopa dywidendy": "{:.2%}", "Wartość pozycji": "{:.2f}", "Szacowany roczny dochód": "{:.2f}"}),
            use_container_width=True,
        )


# ============================================================================
# MODUŁ 6: SKANER GPW (WIG20 & mWIG40)
# ============================================================================

# Reprezentatywna baza tickerów GPW (WIG20 + mWIG40) w formacie Yahoo Finance
# (sufiks .WA). Skład indeksów zmienia się okresowo — lista ma charakter
# poglądowy i można ją swobodnie edytować / rozszerzać.
GPW_SCANNER_TICKERS = [
    # --- WIG20 ---
    "PKN.WA", "PKO.WA", "PEO.WA", "PZU.WA", "KGH.WA",
    "DNP.WA", "CDR.WA", "ALE.WA", "LPP.WA", "CPS.WA",
    "SPL.WA", "MBK.WA", "OPL.WA", "PGE.WA", "TPE.WA",
    "JSW.WA", "KRU.WA", "BDX.WA", "ALR.WA", "KTY.WA",
    # --- mWIG40 ---
    "CCC.WA", "DVL.WA", "ENA.WA", "MIL.WA", "TEN.WA",
    "PLW.WA", "XTB.WA", "ATT.WA", "BFT.WA", "GPW.WA",
    "AMC.WA", "NEU.WA", "APR.WA", "ACP.WA", "R22.WA",
    "11B.WA", "COG.WA", "VRG.WA", "STS.WA", "PEP.WA",
    "WPL.WA", "TXT.WA", "MRB.WA", "ABE.WA", "ASB.WA",
    "GRX.WA", "PXM.WA", "MDG.WA", "ENT.WA", "ATC.WA",
    "SNT.WA", "HUG.WA", "CIG.WA", "ELB.WA", "GEA.WA",
    "OPN.WA", "PCO.WA", "ULM.WA", "WLT.WA", "ONE.WA",
]


@st.cache_data(show_spinner=False, ttl=3600)
def fetch_gpw_scanner_data(tickers: tuple, period: str = "1y") -> pd.DataFrame:
    """Pobiera dane rynkowe (cena, RSI14, spadek od szczytu 52W, docelowa
    cena analityków) dla listy tickerów GPW, wykorzystując tę samą funkcję
    `fetch_market_data`, co reszta aplikacji (jeden spójny cache)."""
    rows = []
    for sym in tickers:
        data = fetch_market_data(sym, period=period)
        rows.append(data)
    df = pd.DataFrame(rows)

    df["drawdown_pct"] = np.where(
        (df["week52_high"].notna()) & (df["week52_high"] > 0) & (df["current_price"].notna()),
        (df["week52_high"] - df["current_price"]) / df["week52_high"] * 100,
        np.nan,
    )
    df["upside_pct"] = np.where(
        (df["target_mean"].notna()) & (df["current_price"].notna()) & (df["current_price"] > 0),
        (df["target_mean"] - df["current_price"]) / df["current_price"] * 100,
        np.nan,
    )
    # Konsensus analityków jest zwyczajowo szacowany na horyzont ~12 miesięcy.
    # Cel 6-miesięczny to uproszczona interpolacja liniowa w połowie drogi —
    # ma charakter wyłącznie poglądowy, nie jest osobną prognozą.
    df["target_12m_pct"] = df["upside_pct"]
    df["target_6m_pct"] = df["upside_pct"] * 0.5
    df["target_12m_price"] = df["target_mean"]
    df["target_6m_price"] = np.where(
        df["current_price"].notna() & df["target_mean"].notna(),
        df["current_price"] + (df["target_mean"] - df["current_price"]) * 0.5,
        np.nan,
    )
    return df


def compute_scanner_score(rsi: float, drawdown_pct: float, upside_pct: float) -> dict:
    """Autorska, zagregowana Ocena Końcowa (1–10) skanera GPW, oparta o trzy
    składowe: wyprzedanie RSI, głębokość spadku od szczytu 52W oraz potencjał
    wzrostu wg średniej ceny docelowej analityków. To narzędzie analityczno-
    -edukacyjne, NIE stanowi rekomendacji inwestycyjnej."""
    # Składowa RSI: im niższe RSI (wyprzedanie), tym wyższy wynik.
    score_rsi = 5.0 if (rsi is None or (isinstance(rsi, float) and math.isnan(rsi))) else float(np.clip((80.0 - rsi) / 6.0, 0.0, 10.0))
    # Składowa spadku od szczytu: głębszy dołek = wyższy wynik (do granicy 50%).
    score_drawdown = 0.0 if (drawdown_pct is None or (isinstance(drawdown_pct, float) and math.isnan(drawdown_pct))) else float(np.clip(drawdown_pct / 5.0, 0.0, 10.0))
    # Składowa potencjału: wyższy szacowany upside wg analityków = wyższy wynik (do granicy 50%).
    score_upside = 0.0 if (upside_pct is None or (isinstance(upside_pct, float) and math.isnan(upside_pct))) else float(np.clip(upside_pct / 5.0, 0.0, 10.0))

    final = 0.40 * score_upside + 0.35 * score_rsi + 0.25 * score_drawdown
    final = float(np.clip(final, 1.0, 10.0))

    reasons = []
    if not math.isnan(rsi) if rsi is not None else False:
        if rsi < 30:
            reasons.append(f"silne wyprzedanie (RSI {rsi:.0f} < 30)")
        elif rsi < 40:
            reasons.append(f"umiarkowane wyprzedanie (RSI {rsi:.0f})")
    if drawdown_pct is not None and not (isinstance(drawdown_pct, float) and math.isnan(drawdown_pct)):
        if drawdown_pct > 25:
            reasons.append(f"głęboki dołek {drawdown_pct:.0f}% od szczytu 52W")
        elif drawdown_pct > 10:
            reasons.append(f"korekta {drawdown_pct:.0f}% od szczytu 52W")
    if upside_pct is not None and not (isinstance(upside_pct, float) and math.isnan(upside_pct)):
        if upside_pct > 25:
            reasons.append(f"{upside_pct:.0f}% potencjału wg analityków")
        elif upside_pct > 10:
            reasons.append(f"{upside_pct:.0f}% szacowanego potencjału")

    if reasons:
        rationale = "Silny sygnał: " + " połączone z ".join(reasons[:2]) + "." if len(reasons) > 1 else reasons[0].capitalize() + "."
    else:
        rationale = "Brak wyraźnych sygnałów — spółka blisko neutralnych poziomów wskaźników."

    return {"score": round(final, 1), "rationale": rationale}


def compute_scanner_status_labels(rsi: float, drawdown_pct: float, upside_pct: float) -> str:
    """Buduje listę etykiet/statusów dla wiersza tabeli skanera GPW."""
    labels = []
    if rsi is not None and not (isinstance(rsi, float) and math.isnan(rsi)) and rsi < 30:
        labels.append("Mocne Wyprzedanie (RSI < 30)")
    if upside_pct is not None and not (isinstance(upside_pct, float) and math.isnan(upside_pct)) and upside_pct > 25:
        labels.append("Duży potencjał (>25%)")
    if drawdown_pct is not None and not (isinstance(drawdown_pct, float) and math.isnan(drawdown_pct)) and drawdown_pct > 25:
        labels.append("Głęboki dołek od szczytu")
    return " | ".join(labels) if labels else "—"


@st.cache_data(show_spinner=False, ttl=3600)
def fetch_company_profile(symbol: str) -> dict:
    """Pobiera dodatkowe dane fundamentalne (wzrost przychodów/zysków,
    sektor, rekomendację analityków) używane WYŁĄCZNIE do wygenerowania
    autorskiego komentarza jakościowego dla TOP 5 okazji skanera GPW.
    Osobny, dedykowany cache — nie modyfikuje `fetch_market_data`."""
    result = {
        "symbol": symbol, "long_name": symbol, "sector": "Nieznany", "industry": "Nieznana",
        "revenue_growth": np.nan, "earnings_growth": np.nan, "recommendation": "", "num_analysts": np.nan,
    }
    try:
        info = yf.Ticker(symbol).info or {}
        result["long_name"] = info.get("longName") or info.get("shortName") or symbol
        result["sector"] = info.get("sector") or "Nieznany"
        result["industry"] = info.get("industry") or "Nieznana"
        result["revenue_growth"] = _safe_float(info.get("revenueGrowth"))
        result["earnings_growth"] = _safe_float(info.get("earningsGrowth"))
        result["recommendation"] = (info.get("recommendationKey") or "").lower()
        result["num_analysts"] = _safe_float(info.get("numberOfAnalystOpinions"))
    except Exception as exc:
        logger.warning("Nie udało się pobrać profilu fundamentalnego dla %s: %s", symbol, exc)
    return result


_RECOMMENDATION_PL = {
    "strong_buy": "zdecydowanie rekomendują kupno",
    "buy": "rekomendują kupno",
    "hold": "rekomendują utrzymanie pozycji",
    "underperform": "rekomendują ostrożność (poniżej rynku)",
    "sell": "rekomendują sprzedaż",
}


def generate_company_narrative(profile: dict, quant_rationale: str) -> str:
    """Buduje autorski komentarz opisowy do jednej z TOP 5 okazji — łączy
    twarde dane fundamentalne z yfinance (sektor, dynamika przychodów/zysków,
    rekomendacja analityków) z uzasadnieniem ilościowym (RSI/spadek/potencjał).
    To narzędzie analityczno-edukacyjne, a nie porada inwestycyjna — komentarz
    NIE jest interpretacją newsów ani sytuacji biznesowej spółki, tylko
    automatycznym opisem dostępnych wskaźników."""
    parts = []
    if profile.get("sector") and profile["sector"] != "Nieznany":
        industry_txt = f" ({profile['industry']})" if profile.get("industry") and profile["industry"] != "Nieznana" else ""
        parts.append(f"Spółka działa w sektorze {profile['sector']}{industry_txt}.")

    rev_g = profile.get("revenue_growth")
    if rev_g is not None and not (isinstance(rev_g, float) and math.isnan(rev_g)):
        if rev_g > 0:
            parts.append(f"Przychody rosną r/r o ok. {rev_g * 100:.0f}%.")
        else:
            parts.append(f"Przychody spadają r/r o ok. {abs(rev_g) * 100:.0f}%.")

    earn_g = profile.get("earnings_growth")
    if earn_g is not None and not (isinstance(earn_g, float) and math.isnan(earn_g)):
        if earn_g > 0:
            parts.append(f"Zyski rosną r/r o ok. {earn_g * 100:.0f}%.")
        else:
            parts.append(f"Zyski spadają r/r o ok. {abs(earn_g) * 100:.0f}%.")

    rec = profile.get("recommendation")
    if rec in _RECOMMENDATION_PL:
        n_analysts = profile.get("num_analysts")
        analysts_txt = f" (na podstawie {int(n_analysts)} analityków)" if n_analysts and not math.isnan(n_analysts) else ""
        parts.append(f"Analitycy {_RECOMMENDATION_PL[rec]}{analysts_txt}.")

    fundamentals_txt = " ".join(parts)
    if fundamentals_txt:
        return f"{fundamentals_txt} Od strony technicznej: {quant_rationale}"
    return quant_rationale


def generate_portfolio_tips(scanned: pd.DataFrame, portfolio_symbols: set) -> list:
    """Generuje 2-3 auto-sugestie ("tipy") łączące wynik skanera GPW z
    aktualnym portfelem użytkownika: (1) najlepiej oceniona posiadana pozycja,
    (2) najsłabiej oceniona posiadana pozycja, (3) najlepsza okazja spoza
    portfela. To narzędzie analityczno-edukacyjne, NIE stanowi rekomendacji
    inwestycyjnej ani porady dot. konkretnych transakcji."""
    tips = []
    if scanned.empty:
        return tips

    held = scanned[scanned["symbol"].isin(portfolio_symbols)].sort_values("score", ascending=False)
    not_held = scanned[~scanned["symbol"].isin(portfolio_symbols)].sort_values("score", ascending=False)

    if not held.empty:
        best_held = held.iloc[0]
        if best_held["score"] >= 6.5:
            tips.append(
                f"💡 Twoja pozycja **{best_held['symbol']}** ma wysoką ocenę skanera "
                f"({best_held['score']:.1f}/10) — {best_held['rationale']} Można rozważyć zwiększenie zaangażowania."
            )
        worst_held = held.iloc[-1]
        if worst_held["score"] <= 4.0 and worst_held["symbol"] != best_held["symbol"]:
            tips.append(
                f"⚠️ Pozycja **{worst_held['symbol']}** w portfelu ma niską ocenę skanera "
                f"({worst_held['score']:.1f}/10) — warto przeanalizować, czy warunki się nie pogorszyły."
            )

    if not not_held.empty and len(tips) < 3:
        top_new = not_held.iloc[0]
        tips.append(
            f"🆕 Poza portfelem wyróżnia się **{top_new['symbol']}** (ocena {top_new['score']:.1f}/10) — "
            f"{top_new['rationale']} Może warto dodać do watchlisty."
        )

    if not tips:
        tips.append("ℹ️ Brak wyraźnych rozbieżności między portfelem a skanerem — obecne pozycje są zbliżone do neutralnych ocen.")

    return tips[:3]


def render_gpw_scanner_tab(portfolio: pd.DataFrame) -> None:
    st.subheader("🔍 Skaner GPW (WIG20 & mWIG40)")
    st.caption(
        "Automatyczny skan ok. 60 spółek z WIG20 i mWIG40 wg RSI(14), spadku od szczytu 52-tygodniowego "
        "oraz prognoz analityków (yfinance). Narzędzie ma charakter analityczno-edukacyjny i NIE stanowi "
        "porady ani rekomendacji inwestycyjnej."
    )

    col_a, col_b = st.columns([3, 1])
    with col_a:
        st.info(f"Analizowana baza: {len(GPW_SCANNER_TICKERS)} spółek (WIG20 + mWIG40).")
    with col_b:
        if st.button("🔄 Odśwież skaner GPW"):
            fetch_gpw_scanner_data.clear()
            st.success("Cache skanera wyczyszczony.")

    progress = st.progress(0.0, text="Skanowanie GPW...")
    # fetch_market_data jest cache'owane per-ticker (@st.cache_data), więc
    # pasek postępu ma charakter informacyjny — realne wywołania sieciowe
    # następują tylko dla tickerów spoza cache.
    scan_df = fetch_gpw_scanner_data(tuple(GPW_SCANNER_TICKERS))
    progress.progress(1.0, text="Gotowe.")
    progress.empty()

    valid = scan_df[scan_df["current_price"].notna()].copy()
    failed = scan_df[scan_df["current_price"].isna()]

    if valid.empty:
        st.error("Nie udało się pobrać danych dla żadnej spółki z listy skanera.")
        return

    scored = valid.apply(
        lambda r: compute_scanner_score(r.get("rsi14"), r.get("drawdown_pct"), r.get("upside_pct")), axis=1, result_type="expand"
    )
    valid = pd.concat([valid.reset_index(drop=True), scored.reset_index(drop=True)], axis=1)
    valid["status_labels"] = valid.apply(
        lambda r: compute_scanner_status_labels(r.get("rsi14"), r.get("drawdown_pct"), r.get("upside_pct")), axis=1
    )
    valid = valid.sort_values("score", ascending=False).reset_index(drop=True)

    if not failed.empty:
        with st.expander(f"⚠️ Brak danych dla {len(failed)} spółek"):
            st.write(", ".join(failed["symbol"].tolist()))

    portfolio_symbols = set(portfolio["symbol_yahoo"].dropna().unique()) if portfolio is not None and not portfolio.empty else set()
    tips = generate_portfolio_tips(valid, portfolio_symbols)
    st.markdown("### 💡 Sugestie dla Twojego portfela")
    for tip in tips:
        st.info(tip)

    st.markdown("### 🔥 TOP 5 Najlepszych Okazji")
    top5 = valid.head(5)
    top_cols = st.columns(5)
    for col, (_, row) in zip(top_cols, top5.iterrows()):
        with col:
            st.metric(row["symbol"], f"{row['score']:.1f} / 10")
            st.caption(row["rationale"])
            currency_suffix = f" {row.get('currency', '')}" if row.get("currency") else ""
            price_txt = fmt_number(row.get("current_price"), 2, currency_suffix)
            target6_txt = fmt_number(row.get("target_6m_price"), 2)
            pct6_txt = fmt_number(row.get("target_6m_pct"), 1, "%")
            target12_txt = fmt_number(row.get("target_12m_price"), 2)
            pct12_txt = fmt_number(row.get("target_12m_pct"), 1, "%")
            st.write(f"**Cena obecna:** {price_txt}")
            st.write(f"**Target 6M:** {target6_txt} ({pct6_txt})")
            st.write(f"**Target 12M:** {target12_txt} ({pct12_txt})")
            profile = fetch_company_profile(row["symbol"])
            narrative = generate_company_narrative(profile, row["rationale"])
            st.caption(f"🗒️ {narrative}")

    st.divider()
    st.markdown("### 📋 Pełna Tabela Skanera")
    display = valid.rename(
        columns={
            "symbol": "Ticker", "current_price": "Cena", "rsi14": "RSI(14)",
            "drawdown_pct": "Spadek od szczytu 52W %", "target_mean": "Śr. cena docelowa",
            "upside_pct": "Potencjał %", "target_6m_price": "Target 6M", "target_6m_pct": "Potencjał 6M %",
            "target_12m_price": "Target 12M", "target_12m_pct": "Potencjał 12M %",
            "score": "Ocena Końcowa", "status_labels": "Statusy", "sector": "Sektor",
        }
    )
    cols = [
        "Ticker", "Ocena Końcowa", "Cena", "RSI(14)", "Spadek od szczytu 52W %",
        "Śr. cena docelowa", "Potencjał %", "Target 6M", "Potencjał 6M %",
        "Target 12M", "Potencjał 12M %", "Statusy", "Sektor",
    ]
    st.dataframe(
        display[cols].style.format(
            {
                "Ocena Końcowa": "{:.1f}", "Cena": "{:.2f}", "RSI(14)": "{:.1f}",
                "Spadek od szczytu 52W %": "{:.1f}%", "Śr. cena docelowa": "{:.2f}",
                "Potencjał %": "{:.1f}%", "Target 6M": "{:.2f}", "Potencjał 6M %": "{:.1f}%",
                "Target 12M": "{:.2f}", "Potencjał 12M %": "{:.1f}%",
            },
            na_rep="—",
        ),
        use_container_width=True,
        height=600,
    )


# ============================================================================
# GŁÓWNA APLIKACJA
# ============================================================================


def main() -> None:
    inject_custom_css()
    st.title("📈 IKE Portfolio Analyzer — XTB")
    st.caption("Analiza portfela, prognozy analityków, wykrywanie okazji i kalkulator rebalansu dla konta IKE Maklerskie.")

    with st.sidebar:
        st.header("⚙️ Ustawienia")
        uploaded_file = st.file_uploader("Wgraj plik XLSX z historii rachunku XTB", type=["xlsx"])
        history_period = st.selectbox(
            "Okres danych historycznych (RSI / SMA)", HISTORY_PERIOD_OPTIONS, index=2,
            help="Dłuższy okres = dokładniejsze SMA200, ale wolniejsze pobieranie.",
        )
        if st.button("🔄 Wymuś odświeżenie danych rynkowych"):
            fetch_market_data.clear()
            st.success("Cache wyczyszczony — dane zostaną pobrane ponownie.")

        st.divider()
        st.subheader("IKE")
        ike_year = st.selectbox("Rok podatkowy", sorted(IKE_LIMITS.keys(), reverse=True))
        ike_contributed_this_year = st.number_input("Wpłacono w tym roku na IKE (PLN)", min_value=0.0, value=0.0, step=100.0)

        st.divider()
        st.subheader("Projekcja długoterminowa")
        cagr_pct = st.slider("Zakładany CAGR (%)", 1, 15, 7)

        st.divider()
        st.caption("⚠️ Aplikacja ma charakter analityczno-edukacyjny. Nie stanowi porady inwestycyjnej ani podatkowej.")

    if uploaded_file is None:
        st.info("👋 Wgraj plik XLSX z historią rachunku XTB w panelu bocznym, aby rozpocząć analizę.")
        st.stop()

    try:
        clean_df = parse_xtb_xlsx(uploaded_file.getvalue())
    except ValueError as exc:
        st.error(f"❌ Błąd podczas parsowania pliku XLSX: {exc}")
        st.stop()
    except Exception as exc:  # zabezpieczenie przed nieoczekiwanymi błędami parsera
        logger.exception("Nieoczekiwany błąd parsowania XLSX")
        st.error(f"❌ Nieoczekiwany błąd podczas przetwarzania pliku: {exc}")
        st.stop()

    positions = aggregate_positions(clean_df)
    if positions.empty:
        st.warning("Nie znaleziono żadnych aktualnie otwartych pozycji w pliku (możliwe, że wszystkie transakcje to sprzedaże).")
        st.stop()

    symbols = positions["symbol_yahoo"].tolist()
    market_rows = []
    progress = st.progress(0.0, text="Pobieranie danych rynkowych...")
    for i, sym in enumerate(symbols):
        market_rows.append(fetch_market_data(sym, period=history_period))
        progress.progress((i + 1) / len(symbols), text=f"Pobieranie danych: {sym}")
    progress.empty()

    market_df = pd.DataFrame(market_rows)
    portfolio = positions.merge(market_df, left_on="symbol_yahoo", right_on="symbol", how="left")

    portfolio["current_value"] = portfolio["volume"] * portfolio["current_price"]
    portfolio["profit_value"] = portfolio["current_value"] - portfolio["total_spent"]
    portfolio["profit_pct"] = np.where(
        portfolio["total_spent"] > 0, portfolio["profit_value"] / portfolio["total_spent"] * 100, np.nan
    )

    dip_records = [
        compute_dip_score(row.get("rsi14", np.nan), row.get("current_price", np.nan), row.get("sma200", np.nan), row.get("week52_high", np.nan))
        for _, row in portfolio.iterrows()
    ]
    portfolio = pd.concat([portfolio.reset_index(drop=True), pd.DataFrame(dip_records)], axis=1)

    errors = portfolio[portfolio["error"].notna()]
    if not errors.empty:
        with st.expander(f"⚠️ Problemy z pobieraniem danych dla {len(errors)} pozycji"):
            for _, row in errors.iterrows():
                st.write(f"**{row['symbol_yahoo']}**: {row['error']}")

    total_value = float(portfolio["current_value"].sum(skipna=True))
    total_invested = float(portfolio["total_spent"].sum(skipna=True))
    total_profit = total_value - total_invested
    total_profit_pct = (total_profit / total_invested * 100) if total_invested else float("nan")
    belka_saved = belka_tax_saved(total_profit)
    ike_limit = get_ike_limit(ike_year)

    render_kpi_row(total_value, total_invested, total_profit, total_profit_pct, belka_saved, ike_contributed_this_year, ike_limit)

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
        [
            "📊 Przegląd Portfela & Ceny na żywo",
            "🎯 Prognozy Analityków (6-12M)",
            "📉 Okazje \"Buy The Dip\"",
            "⚖️ Kalkulator Rebalansu",
            "💰 Projekcja IKE & Dywidendy",
            "🔍 Skaner GPW (WIG20 & mWIG40)",
        ]
    )
    with tab1:
        render_overview_tab(portfolio, clean_df)
    with tab2:
        render_analyst_tab(portfolio)
    with tab3:
        render_dip_tab(portfolio)
    with tab4:
        render_rebalance_tab(portfolio)
    with tab5:
        render_projection_tab(portfolio, total_value, cagr_pct, ike_limit)
    with tab6:
        render_gpw_scanner_tab(portfolio)


if __name__ == "__main__":
    main()
