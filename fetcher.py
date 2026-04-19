"""
fetcher.py — Yahoo Finance price-fetching utilities
"""

from datetime import date, timedelta
from typing import Optional

import yfinance as yf

from config import SYMBOL_OVERRIDES


def build_ticker(symbol: str, exchange: str) -> str:
    """Build a Yahoo Finance ticker string from symbol + exchange."""
    if symbol in SYMBOL_OVERRIDES:
        return SYMBOL_OVERRIDES[symbol]
    suffix = ".NS" if exchange.upper() == "NSE" else ".BO"
    return symbol + suffix


def fetch_current_price(ticker: str) -> Optional[float]:
    """Fetch the latest available price (live or last close)."""
    try:
        info  = yf.Ticker(ticker).fast_info
        price = getattr(info, "last_price", None) or getattr(info, "regular_market_price", None)
        if price:
            return round(float(price), 2)
        hist = yf.Ticker(ticker).history(period="2d")
        if not hist.empty:
            return round(float(hist["Close"].iloc[-1]), 2)
    except Exception:
        pass
    return None


def fetch_price_on_date(ticker: str, target: date) -> Optional[float]:
    """Fetch the closing price on a specific calendar date."""
    try:
        start = target
        end   = target + timedelta(days=1)
        hist  = yf.Ticker(ticker).history(start=start.isoformat(), end=end.isoformat())
        if not hist.empty:
            return round(float(hist["Close"].iloc[-1]), 2)
    except Exception:
        pass
    return None


def fetch_price_range(ticker: str, from_date: date, to_date: date) -> dict[str, Optional[float]]:
    """Return {date_str: close_price} for every trading day in the range."""
    result: dict[str, Optional[float]] = {}
    try:
        end  = to_date + timedelta(days=1)
        hist = yf.Ticker(ticker).history(start=from_date.isoformat(), end=end.isoformat())
        for idx, row in hist.iterrows():
            day_str        = idx.strftime("%d-%b-%Y")
            result[day_str] = round(float(row["Close"]), 2)
    except Exception:
        pass
    return result
