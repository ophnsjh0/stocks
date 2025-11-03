# streamlit_app.py
# -*- coding: utf-8 -*-
import math
from datetime import datetime, timedelta
import pandas as pd
import FinanceDataReader as fdr
import streamlit as st
import altair as alt

# =========================
# 기본 설정
# =========================
st.set_page_config(page_title="K-올웨더(성장형) 배분 계산기", layout="wide")

ASSETS = [
    {"종목명": "KODEX 미국 S&P500TR",           "종목코드": "379800", "비율": 0.24},
    {"종목명": "KOSEF 200TR",                    "종목코드": "294400", "비율": 0.08},
    {"종목명": "KODEX 차이나CSI300",             "종목코드": "283580", "비율": 0.08},
    {"종목명": "KODEX 인도 Nifty50",             "종목코드": "453810", "비율": 0.08},
    {"종목명": "ACE KRX금현물",                  "종목코드": "411060", "비율": 0.19},
    {"종목명": "KODEX 미국채10년선물",           "종목코드": "308620", "비율": 0.07},
    {"종목명": "ACE 미국30년국채액티브(H)",      "종목코드": "453850", "비율": 0.07},
    {"종목명": "KBSTAR KIS 국고채 30년 Enhanced", "종목코드": "385560", "비율": 0.14},
    {"종목명": "TIGER KOFR금리액티브(합성)",      "종목코드": "449170", "비율": 0.05},
]

# 설명 보강
ASSET_DESC = {
    "379800": "S&P 500 총수익(TR) 추종 ETF. 배당 재투자 효과 반영, 미국 대형주 노출.",
    "294400": "KOSPI200 총수익(TR) 추종 ETF. 국내 대형주 대표지수에 배당 재투자 포함.",
    "283580": "중국 CSI300 지수 연동 ETF. 상하이/선전 대형 우량주 중심.",
    "453810": "인도 Nifty50 지수 연동 ETF. 인도 대표 50개 우량주 노출.",
    "411060": "KRX 금 현물 가격 연동 ETF. 원화 기준 금 가격 변동성 반영.",
    "308620": "미국 10년 국채선물 노출 ETF. 중장기 금리 민감도.",
    "453850": "미국 30년 장기국채 액티브 운용, 환헤지(H)로 환율 변동 노출 축소.",
    "385560": "KIS 국고채 30년 듀레이션 강화형 ETF. 초장기 금리 변동에 민감.",
    "449170": "KOFR(무담보콜금리) 연동 단기금리형 ETF(합성). 현금성 대기자금 성격.",
}

KRW_COLS = ["투자금액", "현재가", "실제매수금액", "잔여(목표-실제)"]

# (이미지 스타일에 맞춘) 색상 팔레트 — 선명한 블루/청록/레드/그린/옐로우/보라 계열
PALETTE = [
    "#3B82F6",  # blue
    "#60A5FA",  # light blue
    "#0EA5E9",  # sky
    "#10B981",  # emerald
    "#F59E0B",  # amber
    "#EF4444",  # red
    "#22C55E",  # green
    "#8B5CF6",  # violet
    "#F97316",  # orange
]

# 코드별 색상 매핑
CODE_ORDER = [a["종목코드"] for a in ASSETS]
COLOR_MAP = {code: PALETTE[i % len(PALETTE)] for i, code in enumerate(CODE_ORDER)}

# =========================
# 데이터 함수
# =========================
@st.cache_data(ttl=300, show_spinner=False)
def get_last_price(krx_code: str):
    df = fdr.DataReader(krx_code)
    if df is None or df.empty:
        raise RuntimeError(f"가격 조회 실패: {krx_code}")
    close = float(df["Close"].iloc[-1])
    date = pd.to_datetime(df.index[-1]).to_pydatetime()
    return close, date

@st.cache_data(ttl=900, show_spinner=False)
def get_price_history(krx_code: str, start: datetime | None = None) -> pd.DataFrame:
    if start is None:
        start = datetime.now() - timedelta(days=365 * 20)
    df = fdr.DataReader(krx_code, start)
    if df is None or df.empty:
        raise RuntimeError(f"시세 조회 실패: {krx_code}")
    out = df[["Close"]].copy()
    out.index = pd.to_datetime(out.index)
    out.sort_index(inplace=True)
    return out

def to_index_100(df: pd.DataFrame) -> pd.DataFrame:
    base = df.iloc[0]
    return df / base * 100.0

def build_allocation(total_krw: int):
    rows, dates = [], []
    for a in ASSETS:
        price, d = get_last_price(a["종목코드"])
        dates.append(d)
        target_amt = total_krw * a["비율"]
        qty = math.floor(target_amt / price)
        buy_amt = qty * price
        rows.append({
            "종목명": a["종목명"],
            "종목코드": a["종목코드"],
            "%비율": a["비율"],
            "현재가": price,
            "투자금액": target_amt,          # 목표금액
            "보유수량": qty,
            "실제매수금액": buy_amt,
            "잔여(목표-실제)": target_amt - buy_amt,
        })
    df = pd.DataFrame(rows)
    total_row = {
        "종목명": "합계",
        "종목코드": "",
        "%비율": df["%비율"].sum(),
        "현재가": None,
        "투자금액": df["투자금액"].sum(),
        "보유수량": int(df["보유수량"].sum()),
        "실제매수금액": df["실제매수금액"].sum(),
        "잔여(목표-실제)": df["잔여(목표-실제)"].sum(),
    }
    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    last_updated = max(dates) if dates else None
    return df, last_updated

def format_krw(x):
    try:
        return f"{int(x):,}"
    except Exception:
        return x

# =========================
# UI 상단: 투자금/새로고침
# =========================
st.title("💹 K-올웨더 (성장형) 배분 계산기")
st.caption("※ 실시간/장마감 데이터는 거래소/데이터 제공 상황에 따라 지연될 수 있습니다.")

topL, topR = st.columns([1, 1])
with topL:
    total = st.number_input(
        "총 투자금액 (KRW)",
        min_value=100_000, step=100_000, value=10_000_000, format="%d",
        help="기본값 10,000,000원 (1,000만원)"
    )
with topR:
    if st.button("🔄 가격/시세 캐시 초기화"):
        get_last_price.clear()
        get_price_history.clear()
        st.toast("캐시 초기화 완료. 표/그래프가 곧 갱신됩니다.", icon="🔄")

# 데이터 빌드
try:
    with st.spinner("가격/배분 계산 중..."):
        df_alloc, last_updated = build_allocation(total)
except Exception as e:
    st.error(f"데이터 조회 중 오류가 발생했습니다: {e}")
    st.stop()

# =========================
# 상단 그래프 영역 (Stock peer analysis 스타일)
# =========================
st.markdown("## 📈 Stock peer analysis")

left, right = st.columns([1.1, 5.5])

# ---- 왼쪽: 컨트롤 패널 (Stock tickers, Time horizon, 선택 종목 리스트) ----
with left:
    st.markdown("#### Stock tickers")

    all_options = [f"{a['종목명']} ({a['종목코드']})" for a in ASSETS]
    default_sel = all_options  # 기본 전체 선택
    selected = st.multiselect(
        label="",
        options=all_options,
        default=default_sel,
        help="비교할 종목을 선택/해제하세요.",
    )

    # Time horizon (4단계: 1주, 6주, 1년, 5년)
    st.markdown("#### Time horizon")

    horizon = st.segmented_control(
        "",
        options=["1주", "6주", "1년", "5년"],
        selection_mode="single",
        default="1주",
    )

    # horizon을 날짜 시작점으로 변환
    now = datetime.now()

    if horizon == "1주":
        # 최근 1주일 (7일)
        start_dt = now - timedelta(weeks=1)
    elif horizon == "6주":
        # 최근 6주 (약 42일)
        start_dt = now - timedelta(weeks=6)
    elif horizon == "1년":
        # 최근 1년 (365일)
        start_dt = now - timedelta(days=365)
    elif horizon == "5년":
        # 최근 5년 (1825일)
        start_dt = now - timedelta(days=365 * 5)
    else:
        # 기본값: 1주
        start_dt = now - timedelta(weeks=1)

    st.write("")  # spacing

    # 선택한 종목 — 세로 나열 + 색상 라벨
    # st.markdown("#### Selected")
    # if not selected:
    #     st.info("선택된 종목이 없습니다.")
    # else:
    #     # 종목명 리스트를 코드/색과 함께 세로 표시
    #     for label in selected:
    #         code = label.split("(")[-1].strip(")")
    #         name = label.split("(")[0].strip()
    #         color = COLOR_MAP.get(code, "#888")
    #         # st.markdown(
    #         #     f"""
    #         #     <div style="display:flex;align-items:center;gap:8px;margin:4px 0;">
    #         #         <span style="display:inline-block;width:12px;height:12px;border-radius:3px;background:{color};"></span>
    #         #         <span>{name}</span>
    #         #     </div>
    #         #     """,
    #         #     unsafe_allow_html=True,
    #         # )

# ---- 오른쪽: 라인 차트 (정규화 = 100) ----
with right:
    # 선택된 코드 파싱
    selected_codes = []
    for label in selected:
        code = label.split("(")[-1].strip(")")
        selected_codes.append(code)

    # 데이터 조립
    try:
        series = []
        for code in selected_codes:
            hist = get_price_history(code, start=start_dt)  # Close
            hist = hist.loc[hist.index >= start_dt]
            if hist.empty:
                continue
            hist_100 = to_index_100(hist)
            hist_100["Date"] = hist_100.index
            hist_100["Symbol"] = code
            hist_100.rename(columns={"Close": "Normalized"}, inplace=True)
            series.append(hist_100[["Date", "Symbol", "Normalized"]])

        if series:
            df_hist = pd.concat(series, axis=0, ignore_index=True)
            # 코드 → 종목명 변경 + 고정 색상
            code_to_name = {a["종목코드"]: a["종목명"] for a in ASSETS}
            df_hist["Name"] = df_hist["Symbol"].map(code_to_name)

            # Altair 라인 차트 (고정 컬러 매핑)
            domain = [code_to_name[c] for c in selected_codes]
            range_ = [COLOR_MAP[c] for c in selected_codes]

            chart = (
                alt.Chart(df_hist)
                .mark_line(point=False, strokeWidth=2)
                .encode(
                    x=alt.X("Date:T", title="Date"),
                    y=alt.Y("Normalized:Q", title="Normalized price"),
                    color=alt.Color("Name:N", scale=alt.Scale(domain=domain, range=range_), legend=alt.Legend(title="Stock")),
                    tooltip=[
                        alt.Tooltip("Name:N", title="Stock"),
                        alt.Tooltip("Date:T", title="Date"),
                        alt.Tooltip("Normalized:Q", title="Normalized", format=".2f"),
                    ],
                )
                .interactive()
                .properties(height=420)
            )
            st.altair_chart(chart, use_container_width=True)
        else:
            st.warning("표시할 데이터가 없습니다. 종목/기간을 조정해보세요.")
    except Exception as e:
        st.error(f"가격 변동 그래프 생성 중 오류: {e}")

# =========================
# 배분 상세표
# =========================
st.subheader("📋 배분 상세표")
df_show = df_alloc.copy()
df_show["%비율"] = (df_show["%비율"] * 100).round(2).astype(str) + "%"
for c in KRW_COLS:
    df_show[c] = df_show[c].apply(format_krw)

st.dataframe(df_show, use_container_width=True, hide_index=True)

leftover = df_alloc.loc[df_alloc["종목명"] == "합계", "잔여(목표-실제)"].iloc[0]
m1, m2, m3 = st.columns(3)
m1.metric("총 투자금액(합계)", format_krw(df_alloc.loc[df_alloc["종목명"] == "합계", "투자금액"].iloc[0]) + " 원")
m2.metric("실제매수금액(합계)", format_krw(df_alloc.loc[df_alloc["종목명"] == "합계", "실제매수금액"].iloc[0]) + " 원")
m3.metric("미집행 현금(잔여 합계)", format_krw(leftover) + " 원")

if last_updated:
    st.caption(f"마지막 가격 기준 시점: {last_updated.strftime('%Y-%m-%d %H:%M')}")

st.divider()

# =========================
# 종목별 세부 카드 (카드 내부에 모두 포함 + 작은 글씨)
# =========================
st.subheader("🧾 종목별 세부 카드")
show_cards = st.checkbox("세부 카드 보기", value=True)

# 카드/타이포 스타일 (작게, 컴팩트)
st.markdown("""
<style>
.card-box {
  padding: 0.8rem 0.9rem;
  margin: 0.8rem 0.9rem;
  border: 1px solid #2a2a2a;
  border-radius: 0.75rem;
  background: rgba(255,255,255,0.03);
}
.card-title {
  font-size: 1.0rem; font-weight: 700; margin-bottom: 0.35rem;
}
.card-code {
  font-size: 0.85rem; color: #9aa0a6; margin-bottom: 0.35rem;
}
.card-desc {
  font-size: 0.95rem; line-height: 1.35rem; margin-bottom: 0.5rem;
}
.metric-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 12px 18px;
}
.metric .label {
  font-size: 0.85rem; color: #9aa0a6; margin-bottom: 4px;
}
.metric .value {
  font-size: 1.15rem; font-weight: 700;
}
.metric .value-strong {
  font-size: 1.25rem; font-weight: 800;
}
.metric .suffix { margin-left: 4px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

def _fmt_krw(x: float) -> str:
    try:
        return f"{int(round(float(x))):,}"
    except Exception:
        return str(x)

def _fmt_pct(x: float) -> str:
    try:
        return f"{round(float(x)*100, 2):,.2f}"
    except Exception:
        return str(x)

if show_cards:
    items = df_alloc[df_alloc["종목명"] != "합계"].to_dict(orient="records")

    # 3열 카드 그리드
    for i in range(0, len(items), 3):
        cols = st.columns(3)
        for col, item in zip(cols, items[i:i+3]):
            code = item["종목코드"]
            desc = ASSET_DESC.get(code, "설명 없음")
            name = item["종목명"]

            pct = _fmt_pct(item["%비율"])
            price = _fmt_krw(item["현재가"])
            target_amt = _fmt_krw(item["투자금액"])
            buy_amt = _fmt_krw(item["실제매수금액"])
            qty = f"{int(item['보유수량']):,}"
            leftover = _fmt_krw(item["잔여(목표-실제)"])

            html = f"""
                    <div class="card-box">
                    <div class="card-title">{name}</div>
                    <div class="card-code">종목코드: {code}</div>
                    <div class="card-desc">{desc}</div>
                    <div class="metric-grid">
                        <div class="metric">
                        <div class="label">목표 비중</div>
                        <div class="value-strong">{pct}<span class="suffix">%</span></div>
                        </div>
                        <div class="metric">
                        <div class="label">현재가</div>
                        <div class="value-strong">{price}<span class="suffix">원</span></div>
                        </div>
                        <div class="metric">
                        <div class="label">투자금액(목표)</div>
                        <div class="value">{target_amt}<span class="suffix">원</span></div>
                        </div>
                        <div class="metric">
                        <div class="label">실제매수금액</div>
                        <div class="value">{buy_amt}<span class="suffix">원</span></div>
                        </div>
                        <div class="metric">
                        <div class="label">보유수량(정수주)</div>
                        <div class="value">{qty}<span class="suffix">주</span></div>
                        </div>
                        <div class="metric">
                        <div class="label">잔여(목표-실제)</div>
                        <div class="value">{leftover}<span class="suffix">원</span></div>
                        </div>
                    </div>
                    </div>
                    """
            with col:
                st.markdown(html, unsafe_allow_html=True)


st.markdown(
    """
    **참고**
    - ‘투자금액’은 포트폴리오 목표 비중에 따른 **목표 금액**입니다.
    - ‘실제매수금액’은 정수 주로 환산해 계산하므로 **잔여(목표-실제)**가 발생할 수 있습니다.
    - 상단 그래프는 **처음 시점 = 100 정규화**로 변동률 비교가 용이합니다.
    """
)
