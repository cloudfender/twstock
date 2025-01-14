import yfinance as yf
import mysql.connector
from datetime import datetime, timedelta
import time
from dotenv import load_dotenv
import os
import pandas as pd
from datetime import timezone, timedelta
import argparse

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
    
    def update_daily_data(self, start_year=2010, market_type='twse'):
        """更新股票的每日交易資料
        market_type: 'twse' 或 'tpex'
        """
        try:
            connection = mysql.connector.connect(**self.db_config)
            cursor = connection.cursor()
            
            # 取得股票代碼
            cursor.execute("SELECT code FROM stock_code WHERE market_type = %s", (market_type,))
            stock_codes = {row[0] for row in cursor.fetchall()}
            print(f"需要更新的{market_type}股票數量: {len(stock_codes)}")
            
            # 定時區為台北時間
            tz = timezone(timedelta(hours=8))
            
            # 取得所有交易日
            cursor.execute("""
                SELECT date 
                FROM stock_days 
                WHERE has_trading = 1 
                AND date >= %s 
                ORDER BY date
            """, (f"{start_year}-01-01",))
            
            trading_dates = [row[0] for row in cursor.fetchall()]
            print(f"需要處理的交易日數量: {len(trading_dates)}")
            
            # 處理每個股票
            for stock_id in sorted(stock_codes):
                try:
                    print(f"\n處理股票: {stock_id}")
                    
                    # 檢查是否為有效股票代碼（小於5碼）
                    if len(stock_id) >= 5:
                        continue
                    
                    # 準備 Yahoo Finance 股票代碼
                    if market_type == 'twse':
                        yf_stock_id = f"{stock_id.zfill(4)}.TW"
                    else:  # tpex
                        yf_stock_id = f"{stock_id.zfill(4)}.TWO"
                    
                    print(f"嘗試取得 {yf_stock_id} 的歷史資料")
                    
                    # 定日期範圍
                    start_date = pd.Timestamp(f"{start_year}-01-01", tz=tz)
                    end_date = pd.Timestamp.now(tz=tz)
                    
                    # 取得歷史資料
                    hist = yf.download(
                        yf_stock_id,
                        start=start_date,
                        end=end_date,
                        progress=False
                    )
                    
                    if hist.empty:
                        print(f"股票 {stock_id} 無歷史資料")
                        continue
                        
                    print(f"取得資料筆數: {len(hist)}")
                    print(f"資料日期範圍: {hist.index[0]} 到 {hist.index[-1]}")
                    
                    # 準備資料表名稱
                    table_name = f"stock_{stock_id}"
                    success_count = 0
                    skip_count = 0
                    error_count = 0
                    
                    # 處理每個交易日的資料
                    for date in trading_dates:
                        try:
                            # 轉換日期格式 (設定為日 UTC 00:00)
                            date_str = date.strftime('%Y-%m-%d')
                            date_utc = pd.Timestamp(date_str).tz_localize('UTC')
                            
                            # 檢查是否有該日期的資料
                            if date_utc not in hist.index:
                                continue
                            
                            # 取得當日資料
                            daily_data = hist.loc[date_utc]
                            
                            # 在插入資料前，先檢查是否已存在
                            check_sql = f"SELECT 1 FROM {table_name} WHERE date = %s"
                            cursor.execute(check_sql, (date_str,))
                            exists = cursor.fetchone() is not None

                            if exists:
                                print(f"資料已存在於 {table_name} {date_str}")
                                skip_count += 1
                            else:
                                # 插入資料
                                sql = """
                                    INSERT IGNORE INTO {} 
                                    (date, open_price, high_price, low_price, close_price, volume)
                                    VALUES (%s, %s, %s, %s, %s, %s)
                                """.format(table_name)
                                
                                values = (
                                    date_str,
                                    float(daily_data['Open'].iloc[0]),
                                    float(daily_data['High'].iloc[0]),
                                    float(daily_data['Low'].iloc[0]),
                                    float(daily_data['Close'].iloc[0]),
                                    int(daily_data['Volume'].iloc[0])
                                )
                                
                                print(f"\n準備插入資料到 {table_name}:")
                                print(f"日期: {date_str}")
                                print(f"開盤價: {values[1]}")
                                print(f"最高價: {values[2]}")
                                print(f"最低價: {values[3]}")
                                print(f"收盤價: {values[4]}")
                                print(f"成交量: {values[5]}")
                                
                                try:
                                    cursor.execute(sql, values)
                                    connection.commit()
                                    if cursor.rowcount > 0:
                                        success_count += 1
                                        print(f"成功插入資料")
                                    else:
                                        print(f"插入失敗，但沒有報錯")
                                except Exception as e:
                                    print(f"插入資料時發生錯誤: {e}")
                                    error_count += 1
                            
                        except Exception as e:
                            error_count += 1
                            print(f"處理日期 {date} 時發生錯誤: {e}")
                            continue
                    
                    connection.commit()
                    print(f"股票 {stock_id} 處理完成 - 成功: {success_count}, 跳過: {skip_count}, 錯誤: {error_count}")
                    time.sleep(1)
                    
                except Exception as e:
                    print(f"處理股票 {stock_id} 時發生錯誤: {e}")
                    continue
            
        except Exception as e:
            print(f"更新資料時發生錯誤: {e}")
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

def main():
    # 設定命令列參數
    parser = argparse.ArgumentParser(description='更新股票歷史資料')
    parser.add_argument('--market', type=str, choices=['twse', 'tpex', 'all'],
                      default='all', help='要更新的市場: twse(上市)、tpex(上櫃)或all(全部)')
    parser.add_argument('--year', type=int, default=2010,
                      help='起始年份 (預設: 2010)')
    
    args = parser.parse_args()
    
    updater = StockDailyUpdater()
    
    if args.market in ['twse', 'all']:
        print("開始更新上市股票...")
        updater.update_daily_data(start_year=args.year, market_type='twse')
    
    if args.market in ['tpex', 'all']:
        print("\n開始更新上櫃股票...")
        updater.update_daily_data(start_year=args.year, market_type='tpex')

if __name__ == "__main__":
    main() 