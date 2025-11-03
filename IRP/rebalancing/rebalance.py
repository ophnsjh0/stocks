# -*- coding: utf-8 -*-
"""
K-올웨더(성장형) 리밸런싱 계산기 - 현재가 기준
입력: 작년 보유 데이터(엑셀). '종목코드', '보유수량'(또는 유사명)만 사용하고,
     파일 내 과거 가격/금액/비중은 모두 무시.
처리: FinanceDataReader로 현재가(최근 종가) 조회 → 목표 비중으로 리밸런싱.
출력: 종목별 매수/매도 수량/금액, 요약(총 매수/매도/수수료/순거래금액),
     '입력 총 평가금액' vs '리밸런싱 후 총 평가금액', '수수료 반영 가정 순가치' 비교.
"""

import math
import argparse
from datetime import datetime
from typing import Dict, Optional

import pandas as pd
import FinanceDataReader as fdr


# -----------------------
# 목표 비중 (성장형) - 합계 1.0
# -----------------------
TARGET_WEIGHTS: Dict[str, float] = {
    "379800": 0.24,  # KODEX 미국 S&P500TR
    "294400": 0.08,  # KOSEF 200TR
    "283580": 0.08,  # KODEX 차이나CSI300
    "453810": 0.08,  # KODEX 인도 Nifty50
    "411060": 0.19,  # ACE KRX금현물
    "308620": 0.07,  # KODEX 미국채10년선물
    "453850": 0.07,  # ACE 미국30년국채액티브(H)
    "385560": 0.14,  # KBSTAR KIS 국고채 30년 Enhanced
    "449170": 0.05,  # TIGER KOFR금리액티브(합성)
}

# 수량 컬럼 후보(우선순위)
QTY_CANDIDATES = [
    "보유수량", "구매 수량", "수량", "리밸런싱후수량",
    "Qty", "qty", "quantity", "holdings_qty",
]


# ---------- 유틸 ----------

def _normalize_code(x: Optional[object]) -> Optional[str]:
    """엑셀에서 읽은 종목코드를 안전하게 문자열로 정규화: 379800.0 -> '379800'"""
    if pd.isna(x):
        return None
    s = str(x).strip()
    if not s:
        return None
    try:
        s = str(int(float(s)))  # '379800.0' -> '379800'
    except Exception:
        pass
    return s


def _pick_qty_column(df: pd.DataFrame) -> str:
    """수량 컬럼명을 유연하게 탐지"""
    cols = list(df.columns)
    for name in QTY_CANDIDATES:
        if name in cols:
            return name
    for c in cols:
        if "수량" in str(c):
            return str(c)
    raise ValueError(f"수량 컬럼을 찾지 못했습니다. 후보: {QTY_CANDIDATES} / 실제 컬럼: {cols}")


# ---------- 입력 로드 ----------

def load_holdings(path: str) -> pd.DataFrame:
    """
    작년 보유 데이터(엑셀)를 읽어 '종목명, 종목코드, 보유수량'으로 정규화.
    - 가격/금액/비중 등의 다른 컬럼은 무시
    - 목표에 없는 종목도 포함(이후 0% 비중 처리 → 전량 매도)
    """
    df = pd.read_excel(path)

    # 컬럼 매핑(유연)
    colmap = {}
    for c in df.columns:
        k = str(c).strip()
        if "종목" in k and "코드" not in k: colmap[c] = "종목명"
        elif "코드" in k:                  colmap[c] = "종목코드"
    df = df.rename(columns=colmap)

    if "종목코드" not in df.columns:
        raise ValueError(f"입력 파일에 '종목코드' 컬럼이 없습니다. 실제 컬럼: {list(df.columns)}")

    qty_col = _pick_qty_column(df)

    # 정규화
    df["종목코드"] = df["종목코드"].apply(_normalize_code)
    df = df[df["종목코드"].notna()]
    df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0).astype(int)

    # 종목명 없으면 코드로 대체
    if "종목명" not in df.columns:
        df["종목명"] = df["종목코드"]

    # 같은 코드 합산
    df = df.groupby(["종목코드", "종목명"], as_index=False)[qty_col].sum()

    # 표준 컬럼명
    df = df.rename(columns={qty_col: "보유수량"})
    df = df[df["보유수량"] > 0].reset_index(drop=True)

    if df.empty:
        raise ValueError("유효한 보유 종목(수량>0)이 없습니다. 파일 내용을 확인하세요.")
    return df[["종목명", "종목코드", "보유수량"]]


# ---------- 가격 조회 ----------

def fetch_prices_from_fdr(codes: pd.Series) -> Dict[str, float]:
    """FDR에서 현재가(최근 종가) 조회"""
    prices: Dict[str, float] = {}
    for raw in codes.unique():
        code = _normalize_code(raw)
        if not code:
            raise RuntimeError(f"빈/잘못된 종목코드 발견: {raw!r}")
        hist = fdr.DataReader(code)
        if hist is None or hist.empty:
            raise RuntimeError(f"가격 조회 실패(빈 데이터): {code}")
        prices[code] = float(hist["Close"].iloc[-1])
    return prices


# ---------- 리밸런싱 코어 ----------

def greedy_cash_spend(df: pd.DataFrame, leftover_cash: float) -> pd.DataFrame:
    """
    내림으로 남은 현금을 '목표가치 - 리밸런싱후가치'가 큰 종목부터 1주씩 추가 매수.
    """
    if leftover_cash <= 0:
        return df
    df = df.copy()
    df["gap_per_share"] = (df["목표가치"] - df["리밸런싱후가치"]) / df["현재가"]
    order = df.sort_values("gap_per_share", ascending=False).index.tolist()

    i = 0
    while i < len(order):
        idx = order[i]
        price = df.at[idx, "현재가"]
        if leftover_cash >= price:
            df.at[idx, "거래수량"] += 1
            df.at[idx, "리밸런싱후수량"] += 1
            df.at[idx, "리밸런싱후가치"] = df.at[idx, "리밸런싱후수량"] * price
            leftover_cash -= price
        else:
            i += 1

    # 금액 재계산
    df["매수금액"] = (df["거래수량"].clip(lower=0) * df["현재가"]).round()
    df["매도금액"] = (-df["거래수량"].clip(upper=0) * df["현재가"]).round()
    df["거래금액(순)"] = df["매수금액"] - df["매도금액"]
    return df


def compute_rebalance(holdings: pd.DataFrame,
                      price_map: Dict[str, float],
                      fee_rate: float = 0.0,
                      use_greedy: bool = True) -> (pd.DataFrame, Dict[str, float]):
    """
    - holdings: (종목명, 종목코드, 보유수량)
    - price_map: 종목코드 -> 현재가
    - fee_rate: 매수/매도 수수료율 (예: 0.00015)
    - use_greedy: 잔여현금 1주씩 추가매수 여부
    반환: (결과 DF, 요약 숫자 dict)
    """
    df = holdings.copy()
    df["현재가"] = df["종목코드"].map(price_map).astype(float)
    if df["현재가"].isna().any():
        missing = df.loc[df["현재가"].isna(), "종목코드"].tolist()
        raise ValueError(f"가격을 찾지 못한 코드: {missing}")

    # 현재 평가금액(현재가 × 보유수량)
    df["현재가치"] = df["보유수량"] * df["현재가"]
    total_before = float(df["현재가치"].sum())  # 입력(현재가 기준) 총 평가금액

    # 목표 비중: 목표 리스트에 없으면 0% (전량 매도 대상으로 처리)
    df["목표비중"] = df["종목코드"].map(TARGET_WEIGHTS).fillna(0.0)

    # 목표 합 1.0 확인
    if abs(sum(TARGET_WEIGHTS.values()) - 1.0) > 1e-6:
        raise ValueError("TARGET_WEIGHTS 합계가 1이 아닙니다. 설정을 확인하세요.")

    # 목표가치/목표수량(내림)
    df["목표가치"] = total_before * df["목표비중"]
    df["목표수량(raw)"] = df["목표가치"] / df["현재가"]
    df["목표수량"] = df["목표수량(raw)"].apply(math.floor).astype(int)

    # 거래수량(양수=매수, 음수=매도)
    df["거래수량"] = (df["목표수량"] - df["보유수량"]).astype(int)

    # 리밸런싱 후
    df["리밸런싱후수량"] = df["보유수량"] + df["거래수량"]
    df["리밸런싱후가치"] = df["리밸런싱후수량"] * df["현재가"]
    total_after = float(df["리밸런싱후가치"].sum())  # 이론상 수수료 전에는 total_before와 거의 동일

    # 금액 및 수수료
    df["매수금액"] = (df["거래수량"].clip(lower=0) * df["현재가"])
    df["매도금액"] = (-df["거래수량"].clip(upper=0) * df["현재가"])
    if fee_rate > 0:
        df["매수수수료"] = df["매수금액"] * fee_rate
        df["매도수수료"] = df["매도금액"] * fee_rate
    else:
        df["매수수수료"] = 0.0
        df["매도수수료"] = 0.0

    df["매수금액"] = (df["매수금액"] + df["매수수수료"]).round()
    df["매도금액"] = (df["매도금액"] - df["매도수수료"]).round()
    df["거래금액(순)"] = df["매수금액"] - df["매도금액"]

    # 총액들 계산
    total_buy  = float(df["매수금액"].sum())
    total_sell = float(df["매도금액"].sum())
    total_fee  = float(df["매수수수료"].sum() + df["매도수수료"].sum())
    net_cash   = float(df["거래금액(순)"].sum())  # 매수-매도 (+면 현금 유출)

    # 잔여현금(매도 > 매수) 있으면 부족 종목부터 그리디로 1주씩 추가매수
    leftover_cash = total_sell - total_buy
    if use_greedy and leftover_cash > 0:
        df = greedy_cash_spend(df, leftover_cash)
        # 재계산
        total_buy  = float(df["매수금액"].sum())
        total_sell = float(df["매도금액"].sum())
        net_cash   = float(df["거래금액(순)"].sum())

    # (참고) 수수료 반영 가정 순가치: 리밸런싱 후 총 평가금액 - 총 수수료
    # 실제 계좌 현금/평가 합계는 수수료만큼 줄어드는 효과를 보는 가정
    fee_adjusted_value = total_after - total_fee

    # 보기 좋은 정렬
    cols = [
        "종목명","종목코드","목표비중","현재가",
        "보유수량","현재가치","목표가치","목표수량",
        "거래수량","리밸런싱후수량","리밸런싱후가치",
        "매수금액","매도금액","거래금액(순)"
    ]
    df_out = df[cols].sort_values(["목표비중","종목코드"], ascending=[False, True]).reset_index(drop=True)

    numbers = {
        "입력 총 평가금액(현재가 기준)": total_before,
        "리밸런싱 후 총 평가금액": total_after,
        "총 매수금액": total_buy,
        "총 매도금액": total_sell,
        "총 수수료": total_fee,
        "순거래금액(매수-매도)": net_cash,
        "수수료 반영 가정 순가치": fee_adjusted_value,
        "총 평가금액 변화(후-전)": total_after - total_before,
    }
    return df_out, numbers


# ---------- 출력 ----------

def make_summary(numbers: Dict[str, float]) -> pd.DataFrame:
    rows = [{"항목": k, "금액": int(round(v))} for k, v in numbers.items()]
    return pd.DataFrame(rows)

def write_excel(out_path: str, result: pd.DataFrame, summary: pd.DataFrame, input_preview: pd.DataFrame):
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        result.to_excel(w, index=False, sheet_name="리밸런싱(현재가)")
        summary.to_excel(w, index=False, sheet_name="요약(합계비교)")
        input_preview.to_excel(w, index=False, sheet_name="입력스냅샷(참고)")


# ---------- 메인 ----------

def main():
    ap = argparse.ArgumentParser(description="K-올웨더 성장형 리밸런싱(현재가 기준, 합계 비교 포함)")
    ap.add_argument("--holdings", required=True, help="작년 보유 데이터 엑셀 경로")
    ap.add_argument("--fee", type=float, default=0.0, help="매수/매도 수수료율 (예: 0.00015)")
    ap.add_argument("--no-greedy", action="store_true", help="잔여현금 그리디 추가매수 비활성화")
    args = ap.parse_args()

    holdings = load_holdings(args.holdings)
    input_preview = holdings.copy()

    price_map_now = fetch_prices_from_fdr(holdings["종목코드"])

    result_df, numbers = compute_rebalance(
        holdings,
        price_map_now,
        fee_rate=args.fee,
        use_greedy=(not args.no_greedy),
    )
    summary_df = make_summary(numbers)

    out_name = f"rebalance_now_{datetime.now():%Y%m%d_%H%M}.xlsx"
    write_excel(out_name, result_df, summary_df, input_preview)

    print(f"✅ 저장 완료: {out_name}")
    print("\n[요약]")
    print(summary_df.to_string(index=False))


if __name__ == "__main__":
    main()
