import FinanceDataReader as fdr
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

def export_turtle_final_v2(ticker_symbol, total_capital):
    print(f"[{ticker_symbol}] 터틀 트레이딩 분석(매수금액 포함) 생성 중... (자본금: {total_capital:,}원)")
    
    # 1. 데이터 가져오기
    start_date = (datetime.now() - timedelta(days=365)).strftime('%Y-%m-%d')
    try:
        df = fdr.DataReader(ticker_symbol, start=start_date)
        if df.empty:
            print("데이터를 찾을 수 없습니다.")
            return
    except Exception as e:
        print(f"오류 발생: {e}")
        return

    # 2. TR 계산 및 구성요소 분리
    df['Prev Close'] = df['Close'].shift(1)
    df.dropna(inplace=True)

    df['TR1_A'] = abs(df['High'] - df['Prev Close'])
    df['TR2_B'] = abs(df['Prev Close'] - df['Low'])
    df['TR3_C'] = df['High'] - df['Low']
    df['TR'] = df[['TR1_A', 'TR2_B', 'TR3_C']].max(axis=1)

    # 3. 이동평균 계산 (SMA, MMA, EMA)
    tr_values = df['TR'].values
    n_days = len(tr_values)
    period = 20
    
    sma_values = np.zeros(n_days)
    mma_values = np.zeros(n_days)
    ema_values = np.zeros(n_days)
    
    if n_days < period:
        print("데이터 부족")
        return

    # SMA 계산
    sma_series = df['TR'].rolling(window=period).mean()
    sma_values = sma_series.fillna(0).values

    # MMA, EMA 초기값
    first_seed = np.mean(tr_values[:period])
    mma_values[period-1] = first_seed
    ema_values[period-1] = first_seed

    # MMA, EMA 재귀적 계산
    for i in range(period, n_days):
        current_tr = tr_values[i]
        mma_values[i] = (mma_values[i-1] * 19 + current_tr) / 20
        ema_values[i] = (ema_values[i-1] * 19 + current_tr * 2) / 21

    df['ATR_SMA_20'] = sma_values
    df['ATR_MMA_20'] = mma_values
    df['ATR_EMA_20'] = ema_values

    # 앞쪽 NaN 처리
    df.loc[df.index[:period-1], ['ATR_MMA_20', 'ATR_EMA_20']] = np.nan

    # 4. 엑셀 출력용 데이터 정리
    cols = ['Close', 'TR1_A', 'TR2_B', 'TR3_C', 'TR', 'ATR_SMA_20', 'ATR_MMA_20', 'ATR_EMA_20']
    output_df = df[cols].copy()
    output_df = output_df.tail(60)
    
    int_cols = ['TR1_A', 'TR2_B', 'TR3_C', 'TR', 'ATR_SMA_20', 'ATR_MMA_20', 'ATR_EMA_20']
    output_df[int_cols] = output_df[int_cols].fillna(0).round().astype(int)
    output_df.index = output_df.index.strftime('%Y.%m.%d')

    # 5. 매수 수량 및 금액 계산
    current_price = int(output_df['Close'].iloc[-1])
    current_atr = int(output_df['ATR_EMA_20'].iloc[-1])
    if current_atr <= 0: current_atr = 1

    risk_amt_1pct = total_capital * 0.01
    risk_amt_2pct = total_capital * 0.02
    stop_loss = current_price - (2 * current_atr)

    # 계산 함수 (수량, 금액)
    def calc_qty_amt(risk_money, divisor_atr):
        qty = int(risk_money / divisor_atr)
        amt = qty * current_price
        return qty, amt

    # Case 1: 1N 나누기 (정석)
    qty_1n_1pct, amt_1n_1pct = calc_qty_amt(risk_amt_1pct, current_atr)
    qty_1n_2pct, amt_1n_2pct = calc_qty_amt(risk_amt_2pct, current_atr)

    # Case 2: 2N 나누기 (보수적)
    qty_2n_1pct, amt_2n_1pct = calc_qty_amt(risk_amt_1pct, 2 * current_atr)
    qty_2n_2pct, amt_2n_2pct = calc_qty_amt(risk_amt_2pct, 2 * current_atr)

    # -------------------------------------------------------
    # 6. 엑셀 저장
    # -------------------------------------------------------
    file_name = f"{ticker_symbol}_Turtle_Analysis_V2.xlsx"
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    
    # 상단 표가 길어졌으므로 시작 행을 조금 더 아래로 조정 (15행부터 데이터)
    start_row = 14
    output_df.to_excel(writer, sheet_name='Sheet1', startrow=start_row)

    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # 포맷 정의
    fmt_title = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
    fmt_head  = workbook.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    fmt_val   = workbook.add_format({'border': 1, 'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter'})
    
    # 스타일 (초록: 1% 정석 / 노랑: 2% 공격적)
    fmt_std_qty   = workbook.add_format({'bold': True, 'bg_color': '#E2EFDA', 'border': 1, 'num_format': '#,##0', 'align': 'center'})
    fmt_std_amt   = workbook.add_format({'bg_color': '#E2EFDA', 'border': 1, 'num_format': '#,##0', 'align': 'center', 'font_color': '#548235'}) # 금액은 약간 연하게
    
    fmt_agg_qty  = workbook.add_format({'bold': True, 'bg_color': '#FFF2CC', 'border': 1, 'num_format': '#,##0', 'align': 'center'})
    fmt_agg_amt  = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1, 'num_format': '#,##0', 'align': 'center', 'font_color': '#BF8F00'})

    # --- 상단 요약 ---
    worksheet.merge_range('A1:H1', f"🐢 터틀 트레이딩 종합 리포트 ({ticker_symbol})", fmt_title)

    # 기본 정보
    worksheet.write(2, 0, "총 투자금", fmt_head)
    worksheet.write(2, 1, total_capital, fmt_val)
    worksheet.write(2, 2, "현재가", fmt_head)
    worksheet.write(2, 3, current_price, fmt_val)
    worksheet.write(2, 4, "현재 ATR", fmt_head)
    worksheet.write(2, 5, current_atr, fmt_val)
    worksheet.write(2, 6, "손절가", fmt_head)
    worksheet.write(2, 7, stop_loss, fmt_val)

    # --- 핵심 비교 표 (수량 & 금액) ---
    # [수정된 부분] 헤더 행과 내용 행의 위치 충돌 해결
    
    # 1. 헤더 (Row 4 / 엑셀 5행)
    worksheet.write(4, 0, "구분 (공식)", fmt_head) # A5
    worksheet.write(4, 1, "1% 리스크 (정석)", fmt_head) # B5
    worksheet.write(4, 2, "2% 리스크 (공격적)", fmt_head) # C5

    # 2. Row 1: 방식 1 (Row 5~6 / 엑셀 6~7행 병합)
    worksheet.merge_range('A6:A7', "방식 1: 나누기 1N\n(손절 시 2% 타격)", fmt_head)
    
    # 값 채우기 (방식 1)
    worksheet.write(5, 1, f"수량: {qty_1n_1pct:,} 주", fmt_std_qty)
    worksheet.write(6, 1, f"금액: {amt_1n_1pct:,} 원", fmt_std_amt)
    
    worksheet.write(5, 2, f"수량: {qty_1n_2pct:,} 주", fmt_agg_qty)
    worksheet.write(6, 2, f"금액: {amt_1n_2pct:,} 원", fmt_agg_amt)

    # 3. Row 2: 방식 2 (Row 7~8 / 엑셀 8~9행 병합)
    worksheet.merge_range('A8:A9', "방식 2: 나누기 2N\n(손절 시 1% 타격)", fmt_head)
    
    # 값 채우기 (방식 2)
    worksheet.write(7, 1, f"수량: {qty_2n_1pct:,} 주", fmt_std_qty)
    worksheet.write(8, 1, f"금액: {amt_2n_1pct:,} 원", fmt_std_amt)
    
    worksheet.write(7, 2, f"수량: {qty_2n_2pct:,} 주", fmt_agg_qty)
    worksheet.write(8, 2, f"금액: {amt_2n_2pct:,} 원", fmt_agg_amt)

    # 컬럼 너비 조정
    worksheet.set_column('A:A', 20) 
    worksheet.set_column('B:C', 22) # 금액이 길어지므로 넓게
    worksheet.set_column('D:I', 11)

    # -------------------------------------------------------
    # 차트 (위치는 start_row + 데이터 길이 고려)
    # -------------------------------------------------------
    data_start = start_row + 1
    data_end = start_row + len(output_df)

    # 차트 1: 주가 (최소값 적용)
    min_close = output_df['Close'].min()
    y_min = min_close * 0.99 

    price_chart = workbook.add_chart({'type': 'line'})
    price_chart.add_series({
        'name':       'Close',
        'categories': ['Sheet1', data_start, 0, data_end, 0], 
        'values':     ['Sheet1', data_start, 1, data_end, 1],
        'line':       {'color': '#4472C4', 'width': 2.0},
    })
    price_chart.set_title({'name': f'{ticker_symbol} Price Trend'})
    price_chart.set_y_axis({'min': y_min, 'major_gridlines': {'visible': True}})
    price_chart.set_x_axis({'visible': False})
    price_chart.set_size({'width': 800, 'height': 300})
    worksheet.insert_chart('J2', price_chart)

    # 차트 2: ATR (TR 변동량 + SMA, MMA, EMA)
    atr_chart = workbook.add_chart({'type': 'line'})
    
    # 1. TR (Daily Raw) - 회색 얇은 선
    atr_chart.add_series({
        'name':       'Daily TR',
        'categories': ['Sheet1', data_start, 0, data_end, 0], 
        'values':     ['Sheet1', data_start, 5, data_end, 5], 
        'line':       {'color': '#BFBFBF', 'width': 1.0},
    })
    
    # 2. SMA 20 - 녹색 점선
    atr_chart.add_series({
        'name':       'SMA 20',
        'categories': ['Sheet1', data_start, 0, data_end, 0], 
        'values':     ['Sheet1', data_start, 6, data_end, 6], 
        'line':       {'color': '#00B050', 'width': 1.5, 'dash_type': 'dash'},
    })

    # 3. MMA 20 - 파란색 실선
    atr_chart.add_series({
        'name':       'MMA 20',
        'categories': ['Sheet1', data_start, 0, data_end, 0], 
        'values':     ['Sheet1', data_start, 7, data_end, 7], 
        'line':       {'color': '#0070C0', 'width': 1.5},
    })

    # 4. EMA 20 - 빨간색 굵은 실선
    atr_chart.add_series({
        'name':       'EMA 20',
        'categories': ['Sheet1', data_start, 0, data_end, 0], 
        'values':     ['Sheet1', data_start, 8, data_end, 8], 
        'line':       {'color': '#FF0000', 'width': 2.5},
    })

    atr_chart.set_title({'name': 'Volatility Analysis (TR, SMA, MMA, EMA)'})
    atr_chart.set_size({'width': 800, 'height': 350})
    worksheet.insert_chart('J18', atr_chart)

    writer.close()
    print(f"완료! '{file_name}' 생성됨.")

# --- 메인 실행부 ---
if __name__ == "__main__":
    try:
        t_ticker = input("종목 코드 입력 (예: 005930, BTC-USD): ")
        t_capital_str = input("총 투자금액 입력 (예: 10000000): ")
        t_capital = int(t_capital_str.replace(",", ""))
        
        if t_ticker and t_capital:
            export_turtle_final_v2(t_ticker, t_capital)
    except ValueError:
        print("금액은 숫자로만 입력해주세요.")