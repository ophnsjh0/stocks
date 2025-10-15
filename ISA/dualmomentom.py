# -*- coding: utf-8 -*-
# dualmomentom_returns_report.py
import os
from datetime import datetime
import pandas as pd
import yfinance as yf

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

OUT_DIR = "dual_momentum_out"
os.makedirs(OUT_DIR, exist_ok=True)

# =========================
# 백데이터(의사결정) 티커
# =========================
US_TICKERS = {
    "SPY": "미국 주식SPY",
    "EFA": "선진국 주식EFA",
    "BIL": "초단기채권BIL",
    "AGG": "미국 혼합채권AGG",  # fallback
}

# =========================
# 실제 투자 매핑 (국내 ETF)
# =========================
KR_MAPPING = {
    "SPY": [
        {"분류": "미국",   "종목명": "KODEX 미국S&P500",                "Code": "379800", "환율": "환해지", "비중(%)": 100.0},
    ],
    "EFA": [
        {"분류": "선진국", "종목명": "TIGER 유로스탁스50(합성 H)",     "Code": "195930", "환율": "환노출", "비중(%)": 50.0},
        {"분류": "선진국", "종목명": "TIGER 일본TOPIX(합성 H)",        "Code": "195920", "환율": "환노출", "비중(%)": 50.0},
    ],
    "AGG": [
        {"분류": "채권",   "종목명": "KODEX 미국종합채권SRI액티브(H)", "Code": "437080", "환율": "환노출", "비중(%)": 100.0},
    ],
}

# 국내 ETF 코드 목록 (Returns 시트 계산용)
KR_CODES = ["379800", "195930", "195920", "437080"]

# =========================
# 유틸: 월말 종가 / 12M 수익률
# =========================
def monthly_close(ticker: str, start="2010-01-01") -> pd.Series:
    df = yf.download(ticker, start=start, progress=False)
    if df.empty or "Close" not in df:
        raise RuntimeError(f"{ticker} 데이터가 비어 있습니다.")
    m = df["Close"].resample("M").last().dropna()
    return m

def trailing_12m_return(monthly: pd.Series) -> float:
    """최근 월말 기준 12개월 수익률 (비율, 0.1234=12.34%)"""
    if len(monthly) < 13:
        raise RuntimeError("12개월 수익률 계산에 필요한 월말 데이터가 부족합니다.")
    p0 = float(monthly.iloc[-1])     # 최근 월말
    p12 = float(monthly.iloc[-13])   # 12개월 전 월말
    return (p0 / p12) - 1.0

# =========================
# 의사결정 로직
# =========================
def decide_allocation():
    # 월말 시계열
    m_spy = monthly_close("SPY")
    m_efa = monthly_close("EFA")
    m_bil = monthly_close("BIL")
    m_agg = monthly_close("AGG")

    # 최근 12M 수익률
    r_spy = trailing_12m_return(m_spy)
    r_efa = trailing_12m_return(m_efa)
    r_bil = trailing_12m_return(m_bil)
    r_agg = trailing_12m_return(m_agg)

    # 룰:
    # 1) SPY 12M > BIL 12M → SPY vs EFA 중 12M 높은 ETF
    # 2) 아니면 AGG
    if r_spy > r_bil:
        chosen_us = "SPY" if r_spy >= r_efa else "EFA"
        rule_text = f"[룰1] SPY(12M={r_spy*100:.2f}%) > BIL(12M={r_bil*100:.2f}%) → SPY vs EFA 중 더 높은 12M → {chosen_us}"
    else:
        chosen_us = "AGG"
        rule_text = f"[룰2] SPY(12M={r_spy*100:.2f}%) ≤ BIL(12M={r_bil*100:.2f}%) → AGG 선택"

    # 실제 투자 배분표
    kr_alloc = pd.DataFrame(KR_MAPPING[chosen_us])

    # 요약표 (미국ETF 12M 수익률)
    summary = pd.DataFrame({
        "US_Ticker": ["SPY", "EFA", "BIL", "AGG"],
        "라벨": [US_TICKERS["SPY"], US_TICKERS["EFA"], US_TICKERS["BIL"], US_TICKERS["AGG"]],
        "12M수익률(%)": [round(r_spy*100, 2), round(r_efa*100, 2), round(r_bil*100, 2), round(r_agg*100, 2)]
    })

    # 배너 문구
    alloc_text = " + ".join([f"{row['종목명']}({row['Code']}) {row['비중(%)']:.0f}%" for _, row in kr_alloc.iterrows()])
    banner = f"이번달 실제 투자 대상: {alloc_text}  |  결정근거: {rule_text}"

    # 기준자산의 12M(%) 값 (Allocation 시트에 참고용으로 넣기)
    chosen_12m_pct = r_spy*100 if chosen_us == "SPY" else (r_efa*100 if chosen_us == "EFA" else r_agg*100)

    return summary, kr_alloc, banner, chosen_us, round(chosen_12m_pct, 2)

# =========================
# Returns 시트 데이터 (미국/국내 모두)
# =========================
def build_returns_sheet_data():
    rows = []

    # 미국 ETF 12M
    for t, label in US_TICKERS.items():
        try:
            r = trailing_12m_return(monthly_close(t)) * 100
            rows.append(["미국", label, t, None, None, round(r, 2)])
        except Exception as e:
            rows.append(["미국", label, t, None, None, None])

    # 국내 ETF 12M (야후 '.KS')
    for code in KR_CODES:
        y_ticker = f"{code}.KS"
        # 간단한 라벨 추출(코드→이름을 알면 더 좋지만, 여기서는 코드로 표기)
        label = f"국내 ETF {code}"
        try:
            r = trailing_12m_return(monthly_close(y_ticker)) * 100
            rows.append(["국내", label, None, code, "KS", round(r, 2)])
        except Exception as e:
            rows.append(["국내", label, None, code, "KS", None])

    return pd.DataFrame(rows, columns=["구분","자산라벨","US_Ticker","KR_Code","시장","12M수익률(%)"])

# =========================
# 엑셀 저장
# =========================
def autosize_columns(ws, max_width=46):
    widths = {}
    for row in ws.iter_rows(values_only=True):
        for i, v in enumerate(row, start=1):
            v = "" if v is None else str(v)
            widths[i] = max(widths.get(i, 0), len(v))
    for i, w in widths.items():
        ws.column_dimensions[get_column_letter(i)].width = min(max(w + 2, 10), max_width)

def save_excel(summary: pd.DataFrame, alloc: pd.DataFrame, banner: str, chosen_us: str, chosen_12m_pct: float, returns_df: pd.DataFrame):
    month_str = datetime.now().strftime("%Y-%m")
    xlsx_path = os.path.join(OUT_DIR, f"dualmo_report_{month_str}.xlsx")

    wb = Workbook()
    wb.remove(wb.active)

    title_fill = PatternFill("solid", fgColor="E6F0FF")
    header_fill = PatternFill("solid", fgColor="F2F2F2")
    thin = Side(style="thin", color="D9D9D9")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    # === Sheet 1: Decision (미국ETF 12M 수익률 요약) ===
    ws1 = wb.create_sheet("Decision")
    ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    c = ws1.cell(row=1, column=1, value=f"SPY/EFA/BIL 12M 모멘텀 의사결정 — {month_str}")
    c.font = Font(size=14, bold=True); c.fill = title_fill
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 24

    ws1.cell(row=3, column=1, value=banner).font = Font(bold=True)

    for col, h in enumerate(summary.columns, start=1):
        cell = ws1.cell(row=5, column=col, value=h)
        cell.font = Font(bold=True); cell.fill = header_fill
        cell.border = border_all; cell.alignment = Alignment(horizontal="center")

    for r_idx, row in enumerate(summary.itertuples(index=False), start=6):
        for c_idx, val in enumerate(row, start=1):
            cell = ws1.cell(row=r_idx, column=c_idx, value=val)
            cell.border = border_all
            if summary.columns[c_idx-1].endswith("(%)") and isinstance(val, (int, float)):
                cell.number_format = "0.00%"; cell.value = val / 100.0

    ws1.freeze_panes = "A6"
    autosize_columns(ws1, max_width=36)

    # === Sheet 2: Allocation (실제 투자) ===
    ws2 = wb.create_sheet("Allocation")
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    c2 = ws2.cell(row=1, column=1, value=f"실제 투자 배분 (국내 ETF) — {month_str}")
    c2.font = Font(size=14, bold=True); c2.fill = title_fill
    c2.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 24

    headers2 = ["분류","종목명","Code","환율","비중(%)","(참고) 기준자산","(참고) 기준자산 12M(%)"]
    for col, h in enumerate(headers2, start=1):
        cell = ws2.cell(row=3, column=col, value=h)
        cell.font = Font(bold=True); cell.fill = header_fill
        cell.border = border_all; cell.alignment = Alignment(horizontal="center")

    r_idx = 4
    for _, row in alloc.iterrows():
        ws2.cell(row=r_idx, column=1, value=row["분류"]).border = border_all
        ws2.cell(row=r_idx, column=2, value=row["종목명"]).border = border_all
        ws2.cell(row=r_idx, column=3, value=row["Code"]).border = border_all
        ws2.cell(row=r_idx, column=4, value=row["환율"]).border = border_all

        pct = float(row["비중(%)"]) / 100.0
        c = ws2.cell(row=r_idx, column=5, value=pct)
        c.border = border_all; c.number_format = "0.00%"

        ws2.cell(row=r_idx, column=6, value=US_TICKERS[chosen_us]).border = border_all

        c12 = ws2.cell(row=r_idx, column=7, value=chosen_12m_pct / 100.0)
        c12.border = border_all; c12.number_format = "0.00%"

        r_idx += 1

    ws2.freeze_panes = "A4"
    autosize_columns(ws2, max_width=46)

    # === Sheet 3: Returns (미국/국내 각 자산 12M 수익률) ===
    ws3 = wb.create_sheet("Returns")
    ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    c3 = ws3.cell(row=1, column=1, value=f"각 자산 12개월 수익률 — {month_str}")
    c3.font = Font(size=14, bold=True); c3.fill = title_fill
    c3.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 24

    for col, h in enumerate(returns_df.columns, start=1):
        cell = ws3.cell(row=3, column=col, value=h)
        cell.font = Font(bold=True); cell.fill = header_fill
        cell.border = border_all; cell.alignment = Alignment(horizontal="center")

    for r_idx, row in enumerate(returns_df.itertuples(index=False), start=4):
        for c_idx, val in enumerate(row, start=1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val)
            cell.border = border_all
            if returns_df.columns[c_idx-1].endswith("(%)") and isinstance(val, (int, float)):
                cell.number_format = "0.00%"; cell.value = val / 100.0

    autosize_columns(ws3, max_width=46)

    # 저장
    wb.save(xlsx_path)
    print(f"✅ 엑셀 저장 완료: {xlsx_path}")

# =========================
# main
# =========================
if __name__ == "__main__":
    summary_df, alloc_df, banner_txt, chosen_us, chosen_12m_pct = decide_allocation()
    returns_df = build_returns_sheet_data()
    save_excel(summary_df, alloc_df, banner_txt, chosen_us, chosen_12m_pct, returns_df)
    print("📌", banner_txt)
