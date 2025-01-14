import requests
import mysql.connector
from datetime import datetime, timedelta
import time
from dotenv import load_dotenv
import os

# 載入環境變數
load_dotenv()

class StockDailyUpdater:
    def __init__(self):
        # 從環境變數獲取資料庫設定
        self.db_config = {
            'host': os.getenv('DB_HOST'),
            'user': os.getenv('DB_USER'),
            'password': os.getenv('DB_PASSWORD'),
            'database': os.getenv('DB_DATABASE')
        }
        
    def update_daily_data(self, start_year=2010):
        """更新所有上市股票的每日交易資料"""
        try:
            connection = mysql.connector.connect(**self.db_config)
            cursor = connection.cursor()
            
            # 取得所有上市股票代碼
            cursor.execute("SELECT code FROM stock_code WHERE market_type = 'twse'")
            stock_codes = {row[0] for row in cursor.fetchall()}
            print(f"需要更新的股票數量: {len(stock_codes)}")
            
            # 設定起始日期和結束日期
            start_date = datetime(start_year, 1, 1)
            current_date = datetime.now()
            
            # 如果現在是交易時間（17:00前），則只處理到昨天
            if current_date.hour < 17:
                end_date = (current_date - timedelta(days=1)).date()
                print(f"目前時間未到下午5點，僅更新至：{end_date}")
            else:
                end_date = current_date.date()
            
            # 逐日處理
            check_date = start_date
            while check_date.date() <= end_date:
                date_str = check_date.strftime('%Y-%m-%d')
                
                try:
                    # 檢查是否為交易日
                    cursor.execute("""
                        SELECT has_trading FROM stock_days 
                        WHERE date = %s
                    """, (date_str,))
                    result = cursor.fetchone()
                    
                    if result is None or result[0] == 0:  # 非交易日或無記錄
                        check_date += timedelta(days=1)
                        continue
                    
                    # 使用 TWSE API 獲取當日所有股票資料
                    query_date = check_date.strftime('%Y%m%d')
                    url = f"https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&date={query_date}&type=ALL"
                    print(f"\n{date_str} URL: {url}")
                    
                    response = requests.get(url)
                    data = response.json()
                    
                    if data.get('stat') == 'OK':
                        # 找到個股交易資料的表格
                        stock_data = None
                        for table in data.get('tables', []):
                            if '個股' in table.get('title', ''):
                                stock_data = table.get('data', [])
                                break
                        
                        if stock_data:
                            print(f"\n找到個股資料，第一筆資料範例:")
                            print(f"欄位內容: {stock_data[0]}")
                            print(f"欄位數量: {len(stock_data[0])}")
                            
                            success_count = 0
                            skip_count = 0
                            error_count = 0
                            
                            # 批次處理所有股票資料
                            for row in stock_data:
                                try:
                                    stock_id = row[0].strip()
                                    
                                    # 只處理小於 5 碼的股票
                                    if len(stock_id) >= 5 or stock_id not in stock_codes:
                                        continue
                                    
                                    print(f"\n處理股票: {stock_id}")
                                    print(f"原始資料: {row}")
                                    
                                    # 檢查是否有效的價格資料
                                    if '--' in [row[5], row[6], row[7], row[8]]:
                                        print(f"股票 {stock_id} 有無效價格，跳過")
                                        continue
                                    
                                    # 準備資料
                                    table_name = f"stock_{stock_id}"
                                    values = (
                                        date_str,
                                        float(row[5]),      # 開盤價
                                        float(row[6]),      # 最高價
                                        float(row[7]),      # 最低價
                                        float(row[8]),      # 收盤價
                                        int(row[2].replace(',', ''))  # 成交量
                                    )
                                    
                                    # 顯示實際的 SQL 指令
                                    formatted_sql = f"""
INSERT INTO {table_name} 
(date, open_price, high_price, low_price, close_price, volume)
VALUES 
('{date_str}', {values[1]}, {values[2]}, {values[3]}, {values[4]}, {values[5]});
"""
                                    print("\n實際執行的 SQL 指令:")
                                    print(formatted_sql)
                                    
                                    cursor.execute(f"""
                                        INSERT IGNORE INTO {table_name}
                                        (date, open_price, high_price, low_price, close_price, volume)
                                        VALUES (%s, %s, %s, %s, %s, %s)
                                    """, values)
                                    if cursor.rowcount > 0:
                                        success_count += 1
                                    
                                except Exception as e:
                                    error_count += 1
                                    print(f"處理股票 {stock_id} 發生錯誤: {e}")
                                    continue
                            
                            connection.commit()
                            print(f"日期 {date_str} 處理完成 - 成功: {success_count}, 跳過: {skip_count}, 錯誤: {error_count}")
                    
                    time.sleep(3)  # 避免請求太快
                    check_date += timedelta(days=1)
                    
                except Exception as e:
                    print(f"處理日期 {date_str} 時發生錯誤: {e}")
                    check_date += timedelta(days=1)
            
        except Exception as e:
            print(f"更新資料時發生錯誤: {e}")
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

def main():
    updater = StockDailyUpdater()
    updater.update_daily_data()

if __name__ == "__main__":
    main() 