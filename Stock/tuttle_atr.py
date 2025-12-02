import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta

def export_turtle_atr_ma_analysis(ticker_symbol):
    print(f"[{ticker_symbol}] ATR 이동평균 분석 데이터를 생성 중입니다...")
    
    # 1. 데이터 가져오기 (이동평균 계산을 위해 1년치 확보)
    start_date = (datetime.now() - timedelta(days=365)).strftime('%Y-%m-%d')
    
    try:
        df = fdr.DataReader(ticker_symbol, start=start_date)
        if df.empty:
            print("데이터를 찾을 수 없습니다.")
            return
    except Exception as e:
        print(f"오류 발생: {e}")
        return

    # 2. TR(True Range) 계산
    df['Prev Close'] = df['Close'].shift(1)
    df.dropna(inplace=True)

    df['TR1'] = df['High'] - df['Low']
    df['TR2'] = abs(df['Prev Close'] - df['High'])
    df['TR3'] = abs(df['Prev Close'] - df['Low'])
    
    # 일별 TR (ATR Raw Value)
    df['ATR'] = df[['TR1', 'TR2', 'TR3']].max(axis=1)

    # -------------------------------------------------------
    # [핵심] ATR의 이동평균(Moving Average) 계산
    # -------------------------------------------------------
    # 주가(Close)가 아닌 ATR 컬럼을 기준으로 이동평균을 구합니다.
    df['ATR_MA15'] = df['ATR'].rolling(window=15).mean()
    df['ATR_MA20'] = df['ATR'].rolling(window=20).mean() # 통상적인 터틀의 N값과 유사
    df['ATR_MA55'] = df['ATR'].rolling(window=55).mean()

    # 3. 데이터 자르기 (최근 55일치)
    # 엑셀 출력 순서: 날짜 | 종가 | TR1~3 | ATR | ATR이평들
    cols = ['Close', 'TR1', 'TR2', 'TR3', 'ATR', 'ATR_MA15', 'ATR_MA20', 'ATR_MA55']
    output_df = df[cols].copy()
    
    output_df = output_df.tail(55)
    output_df.index = output_df.index.strftime('%Y.%m.%d')

    # 주가 차트 Y축 설정을 위한 최소값 (그래프 디테일용)
    min_close = output_df['Close'].min()

    # 4. 엑셀 저장
    file_name = f"Ticker-{ticker_symbol}_Tuttle_ATR.xlsx"
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    output_df.to_excel(writer, sheet_name='Sheet1')

    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    row_count = len(output_df)

    # -------------------------------------------------------
    # 차트 1: 주가 (위쪽) - 깔끔하게 주가만 표시
    # -------------------------------------------------------
    price_chart = workbook.add_chart({'type': 'line'})
    price_chart.add_series({
        'name':       'Close Price',
        'categories': ['Sheet1', 1, 0, row_count, 0], 
        'values':     ['Sheet1', 1, 1, row_count, 1], # Close
        'line':       {'color': '#4472C4', 'width': 2.0},
    })
    
    price_chart.set_title({'name': f'{ticker_symbol} Price Trend'})
    price_chart.set_y_axis({'name': 'Price', 'min': min_close, 'major_gridlines': {'visible': True}})
    price_chart.set_x_axis({'visible': False})
    price_chart.set_size({'width': 800, 'height': 300})
    worksheet.insert_chart('J2', price_chart)

    # -------------------------------------------------------
    # 차트 2: ATR + ATR 이동평균선 (아래쪽)
    # -------------------------------------------------------
    atr_chart = workbook.add_chart({'type': 'line'})
    
    # 데이터 컬럼 인덱스 (엑셀 기준 A=0)
    # 0:Date, 1:Close, 2:TR1, 3:TR2, 4:TR3, 5:ATR, 6:MA15, 7:MA20, 8:MA55

    # (1) 일일 ATR (실선, 굵게) - 변동성 그 자체
    atr_chart.add_series({
        'name':       'Daily ATR',
        'categories': ['Sheet1', 1, 0, row_count, 0], 
        'values':     ['Sheet1', 1, 5, row_count, 5], 
        'line':       {'color': '#595959', 'width': 2.0}, # 진한 회색
    })

    # (2) ATR MA 15 (보라색)
    atr_chart.add_series({
        'name':       'ATR MA15',
        'categories': ['Sheet1', 1, 0, row_count, 0], 
        'values':     ['Sheet1', 1, 6, row_count, 6], 
        'line':       {'color': '#7030A0', 'width': 1.0, 'dash_type': 'solid'},
    })

    # (3) ATR MA 20 (녹색) - 터틀의 N값 추세
    atr_chart.add_series({
        'name':       'ATR MA20 (N Trend)',
        'categories': ['Sheet1', 1, 0, row_count, 0], 
        'values':     ['Sheet1', 1, 7, row_count, 7], 
        'line':       {'color': '#00B050', 'width': 1.2, 'dash_type': 'solid'},
    })

    # (4) ATR MA 55 (빨간색) - 장기 변동성 추세
    atr_chart.add_series({
        'name':       'ATR MA55',
        'categories': ['Sheet1', 1, 0, row_count, 0], 
        'values':     ['Sheet1', 1, 8, row_count, 8], 
        'line':       {'color': '#FF0000', 'width': 1.0, 'dash_type': 'solid'},
    })

    atr_chart.set_title({'name': 'Volatility Analysis (ATR & Moving Averages)'})
    atr_chart.set_y_axis({'name': 'True Range'})
    atr_chart.set_x_axis({'name': 'Date'})
    atr_chart.set_size({'width': 800, 'height': 250}) # 높이를 조금 키움
    
    worksheet.insert_chart('J18', atr_chart)

    writer.close()
    print(f"완료! '{file_name}' 파일이 생성되었습니다.")

# --- 메인 실행부 ---
if __name__ == "__main__":
    target_ticker = input("종목 코드 입력 (예: GC=F, 005930): ")
    if target_ticker:
        export_turtle_atr_ma_analysis(target_ticker)