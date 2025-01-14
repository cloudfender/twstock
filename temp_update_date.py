import requests
import mysql.connector
from datetime import datetime, timedelta
import time
from dotenv import load_dotenv
import os

# 載入環境變數
load_dotenv()

class StockDaysUpdater:
    def __init__(self):
        # 從環境變數獲取資料庫設定
        self.db_config = {
            'host': os.getenv('DB_HOST'),
            'user': os.getenv('DB_USER'),
            'password': os.getenv('DB_PASSWORD'),
            'database': os.getenv('DB_DATABASE')
        }
        
        # 確保資料表存在
        self.init_tables()
        
    def init_tables(self):
        """初始化必要的資料表"""
        try:
            connection = mysql.connector.connect(**self.db_config)
            cursor = connection.cursor()
            
            # 建立股票日期記錄表
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS stock_days (
                    date DATE PRIMARY KEY,
                    has_trading TINYINT(1) DEFAULT 0,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # 建立0050的股價資料表（如果不存在）
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS stock_0050 (
                    date DATE PRIMARY KEY,
                    open_price DECIMAL(10,2),
                    high_price DECIMAL(10,2),
                    low_price DECIMAL(10,2),
                    close_price DECIMAL(10,2),
                    volume BIGINT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            connection.commit()
            print("資料表初始化完成")
            
        except Exception as e:
            print(f"初始化資料表時發生錯誤: {e}")
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()
    
    def update_0050_data(self, start_year=2010):
        """更新0050的歷史資料，按天檢查並更新"""
        try:
            connection = mysql.connector.connect(**self.db_config)
            cursor = connection.cursor()
            
            # 設定起始日期和結束日期
            start_date = datetime(start_year, 1, 1)
            current_date = datetime.now()
            
            # 如果現在是交易時間（17:00前），則只處理到昨天
            if current_date.hour < 17:
                end_date = (current_date - timedelta(days=1)).date()
                print(f"目前時間未到下午5點，僅更新至：{end_date}")
            else:
                end_date = current_date.date()
                
            # 逐天檢查
            check_date = start_date
            while check_date.date() <= end_date:
                date_str = check_date.strftime('%Y-%m-%d')
                
                try:
                    # 檢查是否已有0050的交易資料
                    cursor.execute("""
                        SELECT 1 FROM stock_0050 
                        WHERE date = %s
                    """, (date_str,))
                    has_stock_data = cursor.fetchone() is not None
                    
                    # 檢查是否已標記為非交易日
                    cursor.execute("""
                        SELECT has_trading FROM stock_days 
                        WHERE date = %s
                    """, (date_str,))
                    day_record = cursor.fetchone()
                    
                    # 如果已有股票資料或已標記為非交易日，則跳過
                    if has_stock_data or (day_record and day_record[0] == 0):
                        print(f"跳過 {date_str}: 已有��料或已標記為非交易日")
                        check_date += timedelta(days=1)
                        continue
                    
                    # 需要查詢該日資料
                    year = check_date.year
                    month = check_date.month
                    query_date = f"{year}{month:02d}01"
                    
                    # 使用TWSE API獲取當月資料
                    url = f"https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date={query_date}&stockNo=0050"
                    print(f"請求 {year}年{month}月 資料")
                    
                    response = requests.get(url)
                    data = response.json()
                    
                    if data.get('data'):
                        # 找到該日期的資料
                        found_data = False
                        for row in data['data']:
                            # 轉換日期格式（民國轉西元）
                            row_date_str = row[0].replace('/', '-')
                            year = int(row_date_str.split('-')[0]) + 1911
                            row_date = f"{year}-{row_date_str.split('-')[1]}-{row_date_str.split('-')[2]}"
                            
                            if row_date == date_str:
                                found_data = True
                                # 更新0050資料
                                sql = """
                                    INSERT IGNORE INTO stock_0050 
                                    (date, open_price, high_price, low_price, close_price, volume)
                                    VALUES (%s, %s, %s, %s, %s, %s)
                                """
                                values = (
                                    date_str,
                                    float(row[3].replace(',', '')),
                                    float(row[4].replace(',', '')),
                                    float(row[5].replace(',', '')),
                                    float(row[6].replace(',', '')),
                                    int(row[1].replace(',', ''))
                                )
                                cursor.execute(sql, values)
                                
                                # 標記為交易日
                                sql = """
                                    INSERT IGNORE INTO stock_days (date, has_trading)
                                    VALUES (%s, 1)
                                """
                                cursor.execute(sql, (date_str,))
                                print(f"更新 {date_str} 的交易資料")
                                break
                        
                        # 如果當月資料中找不到該日期，且不是當天，則標記為非交易日
                        if not found_data and check_date.date() < current_date.date():
                            sql = """
                                INSERT IGNORE INTO stock_days (date, has_trading)
                                VALUES (%s, 0)
                            """
                            cursor.execute(sql, (date_str,))
                            print(f"標記 {date_str} 為非交易日")
                        
                    connection.commit()
                    time.sleep(3)  # 避免請求過快
                    
                except Exception as e:
                    print(f"處理 {date_str} 時發生錯誤: {e}")
                    
                check_date += timedelta(days=1)
                
        except Exception as e:
            print(f"更新資料時發生錯誤: {e}")
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()
    
    def fill_non_trading_days(self, start_year=2010):
        """填充非交易日期記錄"""
        try:
            connection = mysql.connector.connect(**self.db_config)
            cursor = connection.cursor()
            
            start_date = datetime(start_year, 1, 1)
            # 只處理到昨天
            end_date = datetime.now() - timedelta(days=1)
            
            current_date = start_date
            while current_date.date() <= end_date.date():
                date_str = current_date.strftime('%Y-%m-%d')
                
                # 檢查是否已經有記錄
                cursor.execute("SELECT 1 FROM stock_days WHERE date = %s", (date_str,))
                if not cursor.fetchone():
                    # 如果沒有記錄，新增一筆非交易日記錄
                    cursor.execute("""
                        INSERT IGNORE INTO stock_days (date, has_trading)
                        VALUES (%s, 0)
                    """, (date_str,))
                
                current_date += timedelta(days=1)
            
            connection.commit()
            print(f"非交易日期填充完成 (更新至 {end_date.strftime('%Y-%m-%d')})")
            
        except Exception as e:
            print(f"填充非交易日期時發生錯誤: {e}")
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

def main():
    updater = StockDaysUpdater()
    updater.update_0050_data()  # 更新0050資料
    updater.fill_non_trading_days()  # 填充非交易日記錄

if __name__ == "__main__":
    main()