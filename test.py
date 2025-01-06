import twstock
from datetime import datetime, timedelta
import time
from FinMind.data import DataLoader

def get_stock_data(code, days=30):
    print(f"開始獲取股票 {code} 的數據...")
    
    # 檢查股票類型
    if code in twstock.twse:
        print(f"{code} 是上市股票")
        # 使用 twstock 獲取上市股票數據
        stock = twstock.Stock(code)
        
        # 獲取當前日期
        today = datetime.now()
        data = []
        
        # 從當前月份開始往前查詢，直到收集到足夠的數據
        for i in range((days // 20) + 2):
            month = today.month - i
            year = today.year
            if month <= 0:
                month += 12
                year -= 1
            try:
                print(f"正在獲取 {year}/{month} 的數據...")
                month_data = stock.fetch(year, month)
                if month_data:
                    data.extend(month_data)
                    print(f"成功獲取 {year}/{month} 的數據，共 {len(month_data)} 筆")
                else:
                    print(f"{year}/{month} 沒有數據")
                time.sleep(3)
            except Exception as e:
                print(f"獲取 {year}/{month} 數據失敗: {e}")
                continue
            
            if len(data) >= days:
                break
                
    elif code in twstock.tpex:
        print(f"{code} 是上櫃股票")
        # 使用 FinMind 獲取上櫃股票數據
        dl = DataLoader()
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days)
        
        try:
            df = dl.taiwan_stock_daily(
                stock_id=code,
                start_date=start_date.strftime('%Y-%m-%d'),
                end_date=end_date.strftime('%Y-%m-%d')
            )
            
            # 轉換 FinMind 數據格式為與 twstock 相容的格式
            data = []
            for _, row in df.iterrows():
                data.append(type('Data', (), {
                    'date': datetime.strptime(row['date'], '%Y-%m-%d'),
                    'open': float(row['open']),
                    'high': float(row['max']),
                    'low': float(row['min']),
                    'close': float(row['close']),
                    'transaction': int(row['Trading_Volume'])
                }))
            print(f"成功獲取上櫃股票數據，共 {len(data)} 筆")
            
        except Exception as e:
            print(f"獲取上櫃股票數據失敗: {e}")
            return
            
    else:
        print(f"{code} 不是有效的股票代碼")
        return
    
    if not data:
        print("未能獲取任何數據")
        return
        
    # 排序並只取最近的指定天數
    data = sorted(data, key=lambda x: x.date)[-days:]
    
    # 輸出數據
    print(f"\n{code} 最近 {days} 天的交易數據：")
    print("日期\t\t開盤\t收盤\t最高\t最低\t成交量")
    print("-" * 70)
    for d in data:
        print(f"{d.date}\t{d.open}\t{d.close}\t{d.high}\t{d.low}\t{d.transaction}")

if __name__ == "__main__":
    get_stock_data('1294', 365)  # 測試上櫃股票
