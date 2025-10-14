#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LAA Timing Signals (Monthly, hardened)

타이밍 룰(포트의 25% 슬리브 가정):
  - (S&P500 가격 < SMA) AND (실업률 > 12M MA) → SHY(미국 단기국채)
  - 그 외 → QQQ(나스닥)

개선/방어:
  - FRED 발표 랙(룩어헤드) 방지: 실업률을 fred_lag개월 시프트 후 12M MA 계산
  - yfinance 컬럼 변화 대응: 'Adj Close' 없으면 'Close'
  - SMA 계산 데이터 부족 체크
  - resample('M') → 'ME'로 변경 (FutureWarning 제거)
  - 컬럼명 강제 지정 + 입력 정규화 (SPX 컬럼 유실/멀티인덱스 방지)
  - 동적 SMA 컬럼명 처리
  - CSV 반올림/포맷
  - CLI 인자 지원

설치(uv):
  uv venv .venv && source .venv/bin/activate
  uv pip install pandas yfinance pandas-datareader python-dateutil

실행 예:
  python laa_signals.py --years 25 --sma 200 --fred-lag 1
"""

from __future__ import annotations
import argparse
from pathlib import Path

import pandas as pd
import yfinance as yf
from pandas_datareader import data as pdr

# -----------------------
# 기본 파라미터
# -----------------------
SPX_TICKER_DEFAULT = "^GSPC"   # 필요 시 'SPY'
FRED_SERIES = "UNRATE"         # 미국 실업률(월별, %)
YEARS_DEFAULT = 25
SMA_WINDOW_DEFAULT = 200
FRED_LAG_DEFAULT = 1
OUT_DIR = Path("./laa_out")


# -----------------------
# 데이터 취득 / 정규화
# -----------------------
def fetch_spx(start: str, end: str, ticker: str, sma_window: int) -> pd.DataFrame:
    """
    S&P500(또는 SPY) 일별 종가 + SMA 계산.
    - auto_adjust=False 명시
    - 'Adj Close' 없으면 'Close'
    - 반환 시 컬럼명을 ["SPX", f"SPX_SMA{sma_window}"]로 강제 지정
    """
    df = yf.download(ticker, start=start, end=end, progress=False, auto_adjust=False)
    if df.empty:
        raise RuntimeError(f"{ticker} 데이터를 가져오지 못했습니다.")

    col = "Adj Close" if "Adj Close" in df.columns else "Close"
    px = df[col].copy().dropna()
    if px.shape[0] < sma_window + 5:
        raise RuntimeError(f"SMA{sma_window} 계산에 데이터가 부족합니다. 보유={px.shape[0]}개 일봉.")

    sma = px.rolling(sma_window).mean()
    # 강제 컬럼명 지정 (이름 유실/멀티인덱스 방지)
    out = pd.concat([px, sma], axis=1).dropna()
    out.columns = ["SPX", f"SPX_SMA{sma_window}"]
    return out


def safe_read_fred(start: str, end: str) -> pd.Series:
    """FRED(UNRATE) 안전 호출."""
    try:
        s = pdr.DataReader(FRED_SERIES, "fred", start=start, end=end).iloc[:, 0]
        s.name = "UNRATE"
        return s
    except Exception as e:
        raise RuntimeError(f"FRED(UNRATE) 다운로드 실패: {e}")


def fetch_unrate(start: str, end: str) -> pd.Series:
    return safe_read_fred(start, end)


def _normalize_spx_df(spx: pd.DataFrame | pd.Series, sma_window: int) -> pd.DataFrame:
    """
    build_monthly_signals() 진입 전 방어:
      - Series면 DataFrame으로
      - MultiIndex 컬럼이면 1단계 평탄화
      - 'SPX' 컬럼 없으면 첫 컬럼을 'SPX'로 리네임
      - SMA 컬럼 없으면 즉시 재계산 후 부착
    """
    if isinstance(spx, pd.Series):
        spx = spx.to_frame(name="SPX")

    # MultiIndex → 단순 문자열
    spx.columns = [c[0] if isinstance(c, tuple) else c for c in spx.columns]

    if "SPX" not in spx.columns:
        # 첫 컬럼을 SPX로 간주
        spx = spx.rename(columns={spx.columns[0]: "SPX"})

    sma_cols = [c for c in spx.columns if c.startswith("SPX_SMA")]
    if not sma_cols:
        sma = spx["SPX"].rolling(sma_window).mean()
        spx = pd.concat([spx["SPX"], sma.rename(f"SPX_SMA{sma_window}")], axis=1)

    return spx


# -----------------------
# 시그널 생성
# -----------------------
def build_monthly_signals(
    spx: pd.DataFrame,
    unrate: pd.Series,
    fred_lag_months: int = 1,
    sma_window: int = 200,
) -> pd.DataFrame:
    # 입력 정규화 (SPX/SMA 보장)
    spx = _normalize_spx_df(spx, sma_window=sma_window)

    # 월말 집계: 'ME' = MonthEnd
    spx_m = spx.resample("ME").last()
    spx_m.index.name = "DATE"

    # SMA 컬럼명 확보 (동적)
    sma_cols = [c for c in spx_m.columns if c.startswith("SPX_SMA")]
    if not sma_cols:
        sma = spx_m["SPX"].rolling(sma_window).mean()
        spx_m = pd.concat([spx_m, sma.rename(f"SPX_SMA{sma_window}")], axis=1)
        sma_cols = [c for c in spx_m.columns if c.startswith("SPX_SMA")]
        if not sma_cols:
            raise RuntimeError("SMA 컬럼을 찾지 못했습니다 (SPX_SMA###).")
    sma_col = sma_cols[0]

    # 실업률: 발표 랙 보정 후 12M MA
    un_m = unrate.copy()
    un_m.index = pd.to_datetime(un_m.index)
    un_m = un_m.resample("ME").last()
    un_m_lagged = un_m.shift(fred_lag_months)
    un_ma12 = un_m_lagged.rolling(12).mean().rename("UNRATE_MA12")

    df = pd.concat([spx_m, un_m_lagged.rename("UNRATE"), un_ma12], axis=1).dropna()

    # 조건
    df["PRICE_ABOVE_SMA200"] = (df["SPX"] > df[sma_col])
    df["UNEMP_ABOVE_MA12"]   = (df["UNRATE"] > df["UNRATE_MA12"])

    # 타이밍 자산
    df["TIMING_ASSET"] = df.apply(
        lambda r: "SHY (미국 단기국채)"
        if (not r["PRICE_ABOVE_SMA200"] and r["UNEMP_ABOVE_MA12"])
        else "QQQ (나스닥)",
        axis=1,
    )
    return df


# -----------------------
# 저장/출력
# -----------------------
def round_and_save(signals: pd.DataFrame, out_path: Path) -> None:
    csv_out = signals.copy()
    sma_cols = [c for c in csv_out.columns if c.startswith("SPX_SMA")]
    for c in ["SPX", "UNRATE", "UNRATE_MA12"] + sma_cols:
        if c in csv_out.columns:
            csv_out[c] = csv_out[c].round(2)
    csv_out.to_csv(out_path, encoding="utf-8-sig", date_format="%Y-%m")


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="LAA 월별 타이밍 시그널 생성기")
    ap.add_argument("--years", type=int, default=YEARS_DEFAULT, help="가져올 연도 범위")
    ap.add_argument("--sma", type=int, default=SMA_WINDOW_DEFAULT, help="SMA 윈도우(일)")
    ap.add_argument("--fred-lag", type=int, default=FRED_LAG_DEFAULT, help="FRED 발표 랙(개월)")
    ap.add_argument("--ticker", type=str, default=SPX_TICKER_DEFAULT, help="S&P500 티커(^GSPC 또는 SPY)")
    ap.add_argument("--out", type=str, default=str(OUT_DIR / "laa_signals.csv"), help="출력 CSV 경로")
    return ap.parse_args()


def main():
    args = parse_args()
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    today = pd.Timestamp.today().normalize()
    start = (today - pd.DateOffset(years=args.years)).strftime("%Y-%m-%d")
    end = today.strftime("%Y-%m-%d")

    spx = fetch_spx(start, end, ticker=args.ticker, sma_window=args.sma)
    un  = fetch_unrate(start, end)

    signals = build_monthly_signals(spx, un, fred_lag_months=args.fred_lag, sma_window=args.sma)

    out_path = Path(args.out)
    round_and_save(signals, out_path)

    # 최근 월 요약
    last = signals.dropna().iloc[-1]
    sma_cols = [c for c in signals.columns if c.startswith("SPX_SMA")]
    sma_col = sma_cols[0] if sma_cols else f"SPX_SMA{args.sma}"

    print("=== LAA 타이밍 시그널 (최근 월) ===")
    print(f"기준월: {last.name.strftime('%Y-%m')}")
    print(f"S&P500 종가: {last['SPX']:.2f} | {sma_col}: {last[sma_col]:.2f} | Above? {bool(last['PRICE_ABOVE_SMA200'])}")
    print(f"실업률: {last['UNRATE']:.2f}% | 12M MA: {last['UNRATE_MA12']:.2f}% | Unemp>MA12? {bool(last['UNEMP_ABOVE_MA12'])}")
    print(f"▶ 타이밍 자산 선택: {last['TIMING_ASSET']}")
    print(f"\nCSV 저장: {out_path.resolve()}")


if __name__ == "__main__":
    main()