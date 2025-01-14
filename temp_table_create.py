import mysql.connector
from datetime import datetime
from dotenv import load_dotenv
import os

# 載入環境變數
load_dotenv()

class TableCreator:
    def __init__(self):
        self.db_config = {
            'host': os.getenv('DB_HOST'),
            'user': os.getenv('DB_USER'),
            'password': os.getenv('DB_PASSWORD'),
            'database': os.getenv('DB_DATABASE')
        }

    def create_tables(self):
        """建立所有股票的每日交易資料表"""
        try:
            # 檢查環境變數是否都存在
            if not all(self.db_config.values()):
                raise ValueError("缺少必要的資料庫連線資訊，請檢查 .env 檔案")

            # 連接資料庫
            connection = mysql.connector.connect(**self.db_config)
            cursor = connection.cursor()

            # 獲取所有股票代碼
            cursor.execute("SELECT code FROM stock_code ORDER BY code")
            stock_codes = cursor.fetchall()

            # 記錄開始時間
            start_time = datetime.now()
            print(f"開始建立資料表... {start_time}")

            # 計數器
            total = len(stock_codes)
            success = 0
            failed = 0

            # 為每個股票建立資料表
            for i, (code,) in enumerate(stock_codes, 1):
                try:
                    table_name = f"stock_{code}"
                    
                    # 建立資料表 SQL
                    create_table_sql = f"""
                    CREATE TABLE IF NOT EXISTS {table_name} (
                        `date` date NOT NULL comment '日期',
                        `open_price` decimal(10,2) DEFAULT NULL comment '開盤價',
                        `high_price` decimal(10,2) DEFAULT NULL comment '最高價',
                        `low_price` decimal(10,2) DEFAULT NULL comment '最低價',
                        `close_price` decimal(10,2) DEFAULT NULL comment '收盤價',
                        `volume` bigint(20) DEFAULT NULL comment '成交量',
                        PRIMARY KEY (`date`)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;
                    """
                    
                    # 執行建立資料表
                    cursor.execute(create_table_sql)
                    success += 1
                    print(f"進度: {i}/{total} - 成功建立資料表 {table_name}")

                except Exception as e:
                    failed += 1
                    print(f"建立資料表 stock_{code} 時發生錯誤: {e}")

            # 記錄結束時間
            end_time = datetime.now()
            duration = end_time - start_time

            # 輸出統計資訊
            print("\n=== 建立資料表完成 ===")
            print(f"開始時間: {start_time}")
            print(f"結束時間: {end_time}")
            print(f"耗時: {duration}")
            print(f"總計: {total} 個資料表")
            print(f"成功: {success} 個")
            print(f"失敗: {failed} 個")

        except mysql.connector.Error as err:
            print(f"資料庫錯誤: {err}")
        except ValueError as ve:
            print(f"設定錯誤: {ve}")
        finally:
            if 'connection' in locals() and connection.is_connected():
                cursor.close()
                connection.close()
                print("\n資料庫連線已關閉")

def main():
    creator = TableCreator()
    creator.create_tables()

if __name__ == "__main__":
    main()