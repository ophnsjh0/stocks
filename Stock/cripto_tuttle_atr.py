import pyupbit
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import time

def export_turtle_upbit_full_chart(ticker_symbol, total_capital):
    # 1. 티커 변환 및 초기 설정
    if "/" in ticker_symbol:
        coin = ticker_symbol.split("/")[0].upper()
        upbit_ticker = f"KRW-{coin}"
    else:
        upbit_ticker = ticker_symbol.upper() # 이미 KRW-BTC 형식이거나 BTC만 입력한 경우

    print(f"\n>> [{ticker_symbol}] (업비트 기준: {upbit_ticker}) 분석 시작... (자본금: {total_capital:,}원)")
    
    # 2. 데이터 가져오기 (pyupbit)
    try:
        df = pyupbit.get_ohlcv(upbit_ticker, interval="day", count=200)
        
        if df is None or df.empty:
            print(f"❌ 데이터를 찾을 수 없습니다. 티커를 확인해주세요. ({upbit_ticker})")
            return
            
        # 컬럼명 대문자 변환 (Open, High, Low, Close, Volume)
        df.columns = ['Open', 'High', 'Low', 'Close', 'Volume', 'Value']
        
    except Exception as e:
        print(f"❌ 오류 발생: {e}")
        return

    # 3. TR 계산
    df['Prev Close'] = df['Close'].shift(1)
    df.dropna(inplace=True)

    df['TR1_A'] = abs(df['High'] - df['Prev Close'])
    df['TR2_B'] = abs(df['Prev Close'] - df['Low'])
    df['TR3_C'] = df['High'] - df['Low']
    df['TR'] = df[['TR1_A', 'TR2_B', 'TR3_C']].max(axis=1)

    # 4. 이동평균 (SMA, MMA, EMA)
    tr_values = df['TR'].values
    n_days = len(tr_values)
    period = 20
    
    sma_values = np.zeros(n_days)
    mma_values = np.zeros(n_days)
    ema_values = np.zeros(n_days)
    
    if n_days < period:
        print("❌ 데이터 부족 (최소 20일 이상 필요)")
        return

    # SMA
    sma_series = df['TR'].rolling(window=period).mean()
    sma_values = sma_series.fillna(0).values

    # MMA, EMA 초기값
    first_seed = np.mean(tr_values[:period])
    mma_values[period-1] = first_seed
    ema_values[period-1] = first_seed

    # 재귀적 계산
    for i in range(period, n_days):
        current_tr = tr_values[i]
        mma_values[i] = (mma_values[i-1] * 19 + current_tr) / 20
        ema_values[i] = (ema_values[i-1] * 19 + current_tr * 2) / 21

    df['ATR_SMA_20'] = sma_values
    df['ATR_MMA_20'] = mma_values
    df['ATR_EMA_20'] = ema_values

    # 앞쪽 NaN 처리
    df.loc[df.index[:period-1], ['ATR_MMA_20', 'ATR_EMA_20']] = np.nan

    # 5. 엑셀 데이터 정리
    cols = ['Close', 'TR1_A', 'TR2_B', 'TR3_C', 'TR', 'ATR_SMA_20', 'ATR_MMA_20', 'ATR_EMA_20']
    output_df = df[cols].copy()
    output_df = output_df.tail(60)
    
    int_cols = ['TR1_A', 'TR2_B', 'TR3_C', 'TR', 'ATR_SMA_20', 'ATR_MMA_20', 'ATR_EMA_20']
    output_df[int_cols] = output_df[int_cols].fillna(0).round().astype(int)
    output_df['Close'] = output_df['Close'].fillna(0).round().astype(int)
    output_df.index = output_df.index.strftime('%Y.%m.%d')

    # 6. 매수 수량 및 금액 (소수점 지원)
    current_price = int(output_df['Close'].iloc[-1])
    current_atr = int(output_df['ATR_EMA_20'].iloc[-1])
    if current_atr <= 0: current_atr = 1

    risk_amt_1pct = total_capital * 0.01
    risk_amt_2pct = total_capital * 0.02
    stop_loss = current_price - (2 * current_atr)

    def calc_qty_amt(risk_money, divisor_atr):
        qty = risk_money / divisor_atr 
        amt = qty * current_price       
        return qty, amt

    qty_1n_1pct, amt_1n_1pct = calc_qty_amt(risk_amt_1pct, current_atr)
    qty_1n_2pct, amt_1n_2pct = calc_qty_amt(risk_amt_2pct, current_atr)
    qty_2n_1pct, amt_2n_1pct = calc_qty_amt(risk_amt_1pct, 2 * current_atr)
    qty_2n_2pct, amt_2n_2pct = calc_qty_amt(risk_amt_2pct, 2 * current_atr)

    # -------------------------------------------------------
    # 7. 엑셀 저장 및 차트 그리기
    # -------------------------------------------------------
    safe_ticker = upbit_ticker.replace("-", "_")
    file_name = f"[Cripto]{safe_ticker}.xlsx"
    
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    start_row = 14
    output_df.to_excel(writer, sheet_name='Sheet1', startrow=start_row)

    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # 포맷 설정
    fmt_title = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
    fmt_head  = workbook.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    fmt_val   = workbook.add_format({'border': 1, 'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter'})
    
    fmt_std_qty = workbook.add_format({'bold': True, 'bg_color': '#E2EFDA', 'border': 1, 'align': 'center', 'num_format': '0.0000'})
    fmt_std_amt = workbook.add_format({'bg_color': '#E2EFDA', 'border': 1, 'num_format': '#,##0', 'align': 'center', 'font_color': '#548235'})
    fmt_agg_qty = workbook.add_format({'bold': True, 'bg_color': '#FFF2CC', 'border': 1, 'align': 'center', 'num_format': '0.0000'})
    fmt_agg_amt = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1, 'num_format': '#,##0', 'align': 'center', 'font_color': '#BF8F00'})

    # 상단 요약
    worksheet.merge_range('A1:H1', f"🐢 업비트 터틀 리포트 ({upbit_ticker})", fmt_title)
    worksheet.write(2, 0, "총 투자금", fmt_head)
    worksheet.write(2, 1, total_capital, fmt_val)
    worksheet.write(2, 2, "현재가", fmt_head)
    worksheet.write(2, 3, current_price, fmt_val)
    worksheet.write(2, 4, "현재 ATR", fmt_head)
    worksheet.write(2, 5, current_atr, fmt_val)
    worksheet.write(2, 6, "손절가", fmt_head)
    worksheet.write(2, 7, stop_loss, fmt_val)

    # 테이블
    worksheet.write(4, 0, "구분 (공식)", fmt_head)
    worksheet.write(4, 1, "1% 리스크 (정석)", fmt_head)
    worksheet.write(4, 2, "2% 리스크 (공격적)", fmt_head)

    worksheet.merge_range('A6:A7', "방식 1: 나누기 1N\n(손절 시 2% 타격)", fmt_head)
    worksheet.write(5, 1, f"수량: {qty_1n_1pct:.4f} 개", fmt_std_qty)
    worksheet.write(6, 1, f"금액: {int(amt_1n_1pct):,} 원", fmt_std_amt)
    worksheet.write(5, 2, f"수량: {qty_1n_2pct:.4f} 개", fmt_agg_qty)
    worksheet.write(6, 2, f"금액: {int(amt_1n_2pct):,} 원", fmt_agg_amt)

    worksheet.merge_range('A8:A9', "방식 2: 나누기 2N\n(손절 시 1% 타격)", fmt_head)
    worksheet.write(7, 1, f"수량: {qty_2n_1pct:.4f} 개", fmt_std_qty)
    worksheet.write(8, 1, f"금액: {int(amt_2n_1pct):,} 원", fmt_std_amt)
    worksheet.write(7, 2, f"수량: {qty_2n_2pct:.4f} 개", fmt_agg_qty)
    worksheet.write(8, 2, f"금액: {int(amt_2n_2pct):,} 원", fmt_agg_amt)

    worksheet.set_column('A:A', 20) 
    worksheet.set_column('B:C', 24) 
    worksheet.set_column('D:I', 11)

    # --------------------------------------------------------------------------
    # ★ 수정된 차트 부분 (TR, MMA 포함)
    # --------------------------------------------------------------------------
    data_start = start_row + 1
    data_end = start_row + len(output_df)

    # 1. 가격 차트
    min_close = output_df['Close'].min()
    y_min = min_close * 0.99 
    price_chart = workbook.add_chart({'type': 'line'})
    price_chart.add_series({
        'name':       'Close',
        'categories': ['Sheet1', data_start, 0, data_end, 0], 
        'values':     ['Sheet1', data_start, 1, data_end, 1],
        'line':       {'color': '#4472C4', 'width': 2.0},
    })
    price_chart.set_title({'name': f'{upbit_ticker} Price Trend'})
    price_chart.set_y_axis({'min': y_min, 'major_gridlines': {'visible': True}})
    price_chart.set_x_axis({'visible': False})
    price_chart.set_size({'width': 800, 'height': 300})
    worksheet.insert_chart('J2', price_chart)

    # 2. ATR 차트 (TR, SMA, MMA, EMA 모두 추가)
    atr_chart = workbook.add_chart({'type': 'line'})
    
    # [추가됨] (1) Daily TR (회색 얇은 선) - 5번째 컬럼(F열)
    atr_chart.add_series({
        'name':       'Daily TR',
        'categories': ['Sheet1', data_start, 0, data_end, 0], 
        'values':     ['Sheet1', data_start, 5, data_end, 5], 
        'line':       {'color': '#D9D9D9', 'width': 1.0}, # 연한 회색
    })
    
    # (2) SMA 20 (녹색 점선) - 6번째 컬럼
    atr_chart.add_series({
        'name':       'SMA 20',
        'categories': ['Sheet1', data_start, 0, data_end, 0], 
        'values':     ['Sheet1', data_start, 6, data_end, 6], 
        'line':       {'color': '#00B050', 'width': 1.5, 'dash_type': 'dash'},
    })

    # [추가됨] (3) MMA 20 (파란색 실선) - 7번째 컬럼
    atr_chart.add_series({
        'name':       'MMA 20',
        'categories': ['Sheet1', data_start, 0, data_end, 0], 
        'values':     ['Sheet1', data_start, 7, data_end, 7], 
        'line':       {'color': '#0070C0', 'width': 1.5},
    })

    # (4) EMA 20 (빨간색 굵은 선) - 8번째 컬럼
    atr_chart.add_series({
        'name':       'EMA 20',
        'categories': ['Sheet1', data_start, 0, data_end, 0], 
        'values':     ['Sheet1', data_start, 8, data_end, 8], 
        'line':       {'color': '#FF0000', 'width': 2.5},
    })

    atr_chart.set_title({'name': 'Volatility (Daily TR vs SMA, MMA, EMA)'})
    atr_chart.set_size({'width': 800, 'height': 350})
    worksheet.insert_chart('J18', atr_chart)

    writer.close()
    print(f"✅ 완료! '{file_name}' 생성됨.")

# --- 메인 실행부 ---
if __name__ == "__main__":
    print("==================================================")
    print("🐢 업비트 터틀 리포트 (TR/MMA 차트 포함 버전)")
    print("==================================================")
    
    user_capital = 0
    while True:
        cap_input = input("\n💰 총 투자금액 입력 (예: 4000000) [종료: q]: ").strip().replace(",", "")
        if cap_input.lower() == 'q': exit()
        if cap_input.isdigit():
            user_capital = int(cap_input)
            break
        else:
            print("⚠️ 숫자로만 입력해주세요.")

    while True:
        print(f"\n--------------------------------------------------")
        print(f"현재 설정된 투자금: {user_capital:,}원")
        ticker = input("📈 코인 심볼 입력 (예: BTC/KRW 또는 BTC) [종료: q]: ").strip()
        
        if ticker.lower() in ['q', 'quit', 'exit']:
            print("종료합니다.")
            break
        
        if not ticker: continue
        
        # BTC 입력시 자동 변환
        if "/" not in ticker and "-" not in ticker:
            ticker = f"{ticker}/KRW"

        export_turtle_upbit_full_chart(ticker, user_capital)