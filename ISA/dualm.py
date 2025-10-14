#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dual Momentum (Korean ETFs) – Monthly 12M return comparison + pick
Requires: pip install finance-datareader pandas
"""
import pandas as pd
import FinanceDataReader as fdr
from datetime import datetime
from dateutil.relativedelta import relativedelta
from pathlib import Path

# ---- User settings ----
ETFS = [
    {"group": "미국", "name": "KODEX 미국S&P500", "code": "379800"},          # Equity 1
    {"group": "선진국", "name": "TIGER 유로스탁스50(합성 H)", "code": "195930"}, # Equity 2
    {"group": "선진국", "name": "TIGER 일본TOPIX(합성 H)", "code": "195920"},    # Equity 3
]
BOND = {"group": "채권", "name": "KODEX 미국종합채권SRI액티브(H)", "code": "437080"}  # Defensive asset

# Absolute momentum rule:
# - Pick the equity with the highest 12M return.
# - If its 12M return <= bond 12M return, hold the bond instead.
# You can switch to "0%-threshold" absolute momentum by changing ABS_THRESHOLD_MODE below.
ABS_THRESHOLD_MODE = "bond"   # "bond" | "zero"

# How many years of data to pull (for warm-up)
YEARS = 4

# Output files
OUT_DIR = Path("./dual_momentum_out")
OUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_CSV = OUT_DIR / "dm_signals.csv"
RET_CSV = OUT_DIR / "dm_12m_returns.csv"

def month_end(d: pd.Timestamp) -> pd.Timestamp:
    """Normalize to month-end (exchange calendar agnostic)."""
    d = pd.Timestamp(d).normalize()
    next_month = (d + pd.offsets.MonthBegin(1))
    return (next_month - pd.offsets.Day(1)).normalize()

def get_monthly_adjclose(code: str, start: str, end: str) -> pd.Series:
    """Fetch daily OHLCV and return month-end Adjusted Close series."""
    df = fdr.DataReader(code, start, end)  # index: Date, columns include 'Close'
    if df.empty:
        raise RuntimeError(f"No data for {code}.")
    # FDR for KRX ETFs provides 'Close' (dividends are minimal; use Close as proxy).
    s = df['Close'].copy()
    s.index = pd.to_datetime(s.index)
    m = s.resample('M').last()  # month-end close
    m.name = code
    return m

def compute_12m_return(series: pd.Series) -> pd.Series:
    """12M return: price / price_12m_ago - 1 (align to same index)."""
    return series / series.shift(12) - 1.0

def pick_asset(equity_12m: pd.Series, bond_12m: pd.Series) -> pd.Series:
    """Return a Series of chosen asset codes by month according to Dual Momentum."""
    # At each month, choose equity with highest 12M return
    best_equity_code = equity_12m.idxmax(axis=1)  # row-wise max: returns column (code) name
    best_equity_ret = equity_12m.max(axis=1)

    if ABS_THRESHOLD_MODE == "bond":
        # If best equity 12M <= bond 12M -> hold bond
        pick = best_equity_code.where(best_equity_ret.gt(bond_12m), other=BOND['code'])
    else:
        # If best equity 12M <= 0% -> hold bond
        pick = best_equity_code.where(best_equity_ret.gt(0.0), other=BOND['code'])
    pick.name = "PICK_CODE"
    return pick

def main():
    today = pd.Timestamp.today().normalize()
    start = (today - pd.DateOffset(years=YEARS)).strftime("%Y-%m-%d")
    end = today.strftime("%Y-%m-%d")

    # Fetch monthly prices
    monthly = {}
    for e in ETFS + [BOND]:
        monthly[e['code']] = get_monthly_adjclose(e['code'], start, end)

    monthly_df = pd.DataFrame(monthly).dropna(how='all')
    monthly_df.index.name = "DATE"

    # 12M returns
    ret12 = monthly_df.apply(compute_12m_return).dropna(how='all')
    # Split equities vs bond
    equity_codes = [e['code'] for e in ETFS]
    bond_code = BOND['code']

    equity_12m = ret12[equity_codes].dropna(how='all')
    bond_12m = ret12[bond_code].reindex(equity_12m.index)

    # Align index and compute picks
    ret12_aligned = ret12.reindex(equity_12m.index)
    picks = pick_asset(equity_12m, bond_12m)

    # Build pretty table with names
    code2name = {e['code']: e['name'] for e in ETFS + [BOND]}
    out = pd.DataFrame(index=equity_12m.index)
    for code in equity_codes + [bond_code]:
        out[code2name[code] + " (12M)"] = ret12_aligned[code]

    # Winner per month (name) and signal
    out["월별 최고 수익률(주식)"] = equity_12m.max(axis=1)
    out["월별 최고 종목코드(주식)"] = equity_12m.idxmax(axis=1)
    out["월별 최고 종목명(주식)"] = out["월별 최고 종목코드(주식)"].map(code2name)
    out["채권 12M"] = bond_12m
    out["선택(코드)"] = picks
    out["선택(이름)"] = picks.map(code2name)

    # Save returns and signals
    out.to_csv(RET_CSV, encoding="utf-8-sig")
    sig = out[["선택(코드)", "선택(이름)"]].copy()
    sig.to_csv(LOG_CSV, encoding="utf-8-sig")
    
    # Current month recommendation (last available row)
    last = out.dropna().iloc[-1]
    print("=== 듀얼 모멘텀 월간 선택 ===")
    print(f"기준월: {last.name.strftime('%Y-%m')}")
    print(f"선택: {last['선택(이름)']} ({last['선택(코드)']})")
    print("\n요약 (12M 수익률 기준):")
    for code in equity_codes + [bond_code]:
        nm = code2name[code] + " (12M)"
        print(f" - {nm}: {last[nm]:.2%}")
    print(f"\n룰: 최고 주식 12M {'>' if ABS_THRESHOLD_MODE=='bond' else '>'} {'채권 12M' if ABS_THRESHOLD_MODE=='bond' else '0%'}이면 그 주식, 아니면 채권")
    print(f"\nCSV 저장: {RET_CSV}")
    print(f"시그널 로그: {LOG_CSV}")

if __name__ == "__main__":
    main()
