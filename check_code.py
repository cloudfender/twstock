import mysql.connector
from twstock import Stock
import time
import twstock

class DatabaseConnection:
    def __init__(self):
        self.connection = None

    def connect(self):
        if not self.connection:
            self.connection = mysql.connector.connect(
                host='192.168.1.145',
                user='root',
                password='fender1106',
                database='stock_code'
            )
        return self.connection

    def close(self):
        if self.connection:
            self.connection.close()
            self.connection = None

# 創建全局連接對象
db = DatabaseConnection()

# 定义保存到数据库的函数
def save_to_database(code, name, market_type):
    try:
        connection = db.connect()
        cursor = connection.cursor()
        cursor.execute("""
            INSERT INTO stock_code (code, name, market_type) 
            VALUES (%s, %s, %s)
        """, (code, name, market_type))
        connection.commit()
    except mysql.connector.Error as err:
        print(f"數據庫錯誤: 無法連接到數據庫: {err}")
    finally:
        cursor.close()

# 检查代码是否已存在于数据库
def code_exists_in_database(code):
    try:
        connection = db.connect()
        cursor = connection.cursor()
        cursor.execute("SELECT 1 FROM stock_code WHERE code = %s", (code,))
        return cursor.fetchone() is not None
    except mysql.connector.Error as err:
        print(f"數據庫錯誤: 無法連接到數據庫: {err}")
        return False
    finally:
        cursor.close()

# 检查代码是否在无效代码表中
def code_exists_in_invalid_database(code):
    try:
        connection = db.connect()
        cursor = connection.cursor(buffered=True)
        cursor.execute("SELECT 1 FROM stack_Ncode WHERE code = %s", (code,))
        return cursor.fetchone() is not None
    except mysql.connector.Error as err:
        print(f"數據庫錯誤: 無法連接到數據庫: {err}")
        return False
    finally:
        cursor.close()

# 保存无效代码到另一个表
def save_invalid_code(code):
    try:
        connection = db.connect()
        cursor = connection.cursor()
        cursor.execute("INSERT INTO stack_Ncode (code) VALUES (%s)", (code,))
        connection.commit()
    except mysql.connector.Error as err:
        print(f"數據庫錯誤: 無法連接到數據庫: {err}")
    finally:
        cursor.close()

# 自动检查所有股票代
def check_all_codes():
    try:
        for code in range(10000):
            code_str = f"{code:04d}"
            # 检查是否在有效或无效表中
            if code_exists_in_database(code_str):
                print(f"股票代號 {code_str} 已存在，跳過")
                continue
            if code_exists_in_invalid_database(code_str):
                print(f"股票代號 {code_str} 已標記為無效，跳過")
                continue
                
            try:
                # 檢查是否為上市股票
                if code_str in twstock.twse:
                    try:
                        stock = Stock(code_str)
                        stock_name = twstock.codes[code_str].name
                        save_to_database(code_str, stock_name, 'twse')
                        print(f"股票代號 {code_str} {stock_name} (上市) 已記錄到數據庫")
                    except Exception as e:
                        print(f"處理上市股票 {code_str} 時發生錯誤: {e}")
                        save_invalid_code(code_str)
                    continue
                
                # 檢查是否為上櫃股票
                if code_str in twstock.tpex:
                    try:
                        # 對於上櫃股票，直接從 twstock.codes 獲取名稱
                        stock_name = twstock.codes[code_str].name
                        save_to_database(code_str, stock_name, 'tpex')
                        print(f"股票代號 {code_str} {stock_name} (上櫃) 已記錄到數據庫")
                    except Exception as e:
                        print(f"處理上櫃股票 {code_str} 時發生錯誤: {e}")
                        save_invalid_code(code_str)
                    continue
                
                # 如果都不是，標記為無效
                print(f"股票代號 {code_str} 無效，跳過")
                save_invalid_code(code_str)
                
            except Exception as e:
                print(f"處理股票代號 {code_str} 時發生錯誤: {e}")
                save_invalid_code(code_str)
                
            time.sleep(0.05)  # 每0.05秒檢查一次
    finally:
        db.close()

if __name__ == "__main__":
    check_all_codes() 