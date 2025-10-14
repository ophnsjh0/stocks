======================================================================
1. Dual Momentum (Korean ETFs)

## 설치
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install --upgrade pip
pip install finance-datareader pandas python-dateutil
```

## 실행
```bash
python dual_momentum_korea.py
```

## 규칙
- 매월 말 기준으로 12개월 수익률(현재가 / 12개월 전 - 1)을 계산
- 주식 3종 중 12개월 수익률이 가장 높은 종목을 후보로 선정
- 후보의 12M 수익률이 채권 12M 수익률(또는 0%)보다 낮으면 채권을 보유

## 출력
- `dual_momentum_out/dm_12m_returns.csv`: 월별 12M 수익률 표 + 선택 결과
- `dual_momentum_out/dm_signals.csv`: 월별 선택(코드/이름)만 정리한 로그

========================================================================


2. LAA 타이밍 시그널 (Monthly)

## 규칙 설명
- **S&P500 200일 이동평균**: 일별 종가의 200거래일 단순이동평균(SMA)
- **미국 실업률 12개월 이동평균**: FRED의 UNRATE(月次) 12개월 단순이동평균
- **타이밍 자산 선택**:
  - (가격 < SMA200) AND (실업률 > 12M MA) → **SHY(미국 단기국채)**
  - 그 외 → **QQQ(나스닥)**

"""
LAA Timing Signals (Monthly, robust version)

타이밍 룰(포트의 25% 슬리브 가정):
  - (S&P500 가격 < 200일 SMA) AND (미국 실업률 > 12개월 이동평균) → SHY(미국 단기국채)
  - 그 외 → QQQ(나스닥)

개선점 반영:
  1) FRED 발표 지연(룩어헤드) 방지: 실업률을 fred_lag개월 만큼 뒤로 시프트한 값으로 12M MA 계산
  2) yfinance 컬럼 변화 방어: 'Adj Close' 없으면 'Close' 폴백 + auto_adjust 명시
  3) 충분한 히스토리 검사: SMA 계산 가능한 최소 길이 체크
  4) CSV 반올림/포맷 정리
  5) 네트워크/서비스 예외 처리
  6) CLI 인자 지원 (연도 범위, SMA 윈도우, FRED 랙 등)

의존성:
  - pandas, yfinance, pandas-datareader, python-dateutil

설치(uv 예시):
  uv venv .venv && source .venv/bin/activate
  uv pip install pandas yfinance pandas-datareader python-dateutil

실행:
  python laa_signals.py --years 25 --sma 200 --fred-lag 1

자동화 팁:
  - 매월 1~5일에 실행(실업률 발표 반영 + 주말/공휴일 보정)
"""

==============================================================================
