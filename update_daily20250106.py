import tkinter as tk
from tkinter import ttk, messagebox
import mysql.connector
import requests
from datetime import datetime, timedelta
import time
from dotenv import load_dotenv
import os

# 載入環境變數
load_dotenv()

class StockUpdaterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("股票資料回補系統 2025")
        self.root.geometry("1000x800")
        
        # 從環境變數獲取資料庫設定
        self.db_config = {
            'host': os.getenv('DB_HOST'),
            'user': os.getenv('DB_USER'),
            'password': os.getenv('DB_PASSWORD'),
            'database': os.getenv('DB_DATABASE')
        }
        
        # 檢查資料庫設定
        if not all(self.db_config.values()):
            messagebox.showerror("錯誤", "缺少必要的資料庫連線資訊，請檢查 .env 檔案")
            root.destroy()
            return
        
        # 創建主框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 建立上市股票區域
        self.create_twse_frame()
        
        # 建立上櫃股票區域
        self.create_tpex_frame()
        
        # 建立日誌區域
        self.create_log_frame()
        
        # 暫停狀態
        self.is_paused = False
        
    def create_twse_frame(self):
        """建立上市股票區域"""
        twse_frame = ttk.LabelFrame(self.main_frame, text="上市股票", padding="5")
        twse_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # 進度條
        self.twse_progress = ttk.Progressbar(twse_frame, mode='determinate')
        self.twse_progress.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # 進度標籤
        self.twse_label = ttk.Label(twse_frame, text="準備就緒")
        self.twse_label.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=5)
        
        # 控制按鈕
        control_frame = ttk.Frame(twse_frame)
        control_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(control_frame, text="開始更新", 
                  command=self.update_twse).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="暫停/繼續", 
                  command=self.toggle_pause).pack(side=tk.LEFT, padx=5)
        
    def create_tpex_frame(self):
        """建立上櫃股票區域"""
        tpex_frame = ttk.LabelFrame(self.main_frame, text="上櫃股票", padding="5")
        tpex_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # 進度條
        self.tpex_progress = ttk.Progressbar(tpex_frame, mode='determinate')
        self.tpex_progress.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # 進度標籤
        self.tpex_label = ttk.Label(tpex_frame, text="準備就緒")
        self.tpex_label.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=5)
        
        # 控制按鈕
        control_frame = ttk.Frame(tpex_frame)
        control_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(control_frame, text="開始更新", 
                  command=self.update_tpex).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="暫停/繼續", 
                  command=self.toggle_pause).pack(side=tk.LEFT, padx=5)
        
    def create_log_frame(self):
        """建立日誌區域"""
        log_frame = ttk.LabelFrame(self.main_frame, text="執行日誌", padding="5")
        log_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 日誌文本框
        self.log_text = tk.Text(log_frame, height=20, width=100)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 捲軸
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
    def log(self, message):
        """添加日誌訊息"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert(tk.END, f"[{current_time}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def toggle_pause(self):
        """切換暫停狀態"""
        self.is_paused = not self.is_paused
        status = "暫停" if self.is_paused else "繼續"
        self.log(f"更新作業已{status}")
        
    def update_twse(self):
        """更新上市股票資料"""
        try:
            connection = mysql.connector.connect(**self.db_config)
            cursor = connection.cursor()
            
            # 獲取上市股票列表
            cursor.execute("SELECT code, name FROM stock_code WHERE market_type = 'twse' ORDER BY code")
            stocks = cursor.fetchall()
            
            total = len(stocks)
            self.twse_progress['maximum'] = total
            self.twse_progress['value'] = 0
            
            # 設定起始日期和結束日期
            start_year = 2010
            current_date = datetime.now()
            
            # 如果現在是交易時間（17:00前），則只處理到昨天
            if current_date.hour < 17:
                end_date = (current_date - timedelta(days=1)).date()
                self.log(f"目前時間未到下午5點，僅更新至：{end_date}")
            else:
                end_date = current_date.date()
            
            for i, (stock_id, name) in enumerate(stocks):
                try:
                    # 檢查暫停狀態
                    while self.is_paused:
                        self.root.update()
                        time.sleep(0.1)
                        continue
                    
                    self.log(f"正在更新上市股票 {stock_id} {name} ({i+1}/{total})")
                    
                    # 檢查資料表是否存在
                    table_name = f"stock_{stock_id}"
                    
                    # 設定起始日期
                    check_date = datetime(start_year, 1, 1)
                    
                    # 逐日檢查
                    while check_date.date() <= end_date:
                        date_str = check_date.strftime('%Y-%m-%d')
                        
                        try:
                            # 檢查是否已有交易資料
                            cursor.execute(f"""
                                SELECT 1 FROM {table_name} 
                                WHERE date = %s
                            """, (date_str,))
                            has_stock_data = cursor.fetchone() is not None
                            
                            # 如果已有資料，跳過
                            if has_stock_data:
                                self.log(f"跳過 {date_str}: 已有資料")
                                check_date += timedelta(days=1)
                                continue
                            
                            # 需要查詢該日資料
                            year = check_date.year
                            month = check_date.month
                            query_date = f"{year}{month:02d}01"
                            
                            # 使用 TWSE API 獲取當月�料
                            url = f"https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date={query_date}&stockNo={stock_id}"
                            self.log(f"請求 {year}年{month}月 資料")
                            
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
                                        # 準備 SQL 參數
                                        sql_params = (
                                            date_str,
                                            float(row[3].replace(',', '')),  # 開盤價
                                            float(row[4].replace(',', '')),  # 最高價
                                            float(row[5].replace(',', '')),  # 最低價
                                            float(row[6].replace(',', '')),  # 收盤價
                                            int(row[1].replace(',', ''))     # 成交量
                                        )
                                        
                                        # 建立 SQL 指令
                                        sql = f"""
                                            INSERT IGNORE INTO {table_name}
                                            (date, open_price, high_price, low_price, close_price, volume)
                                            VALUES (%s, %s, %s, %s, %s, %s)
                                        """
                                        
                                        # 執行 SQL
                                        cursor.execute(sql, sql_params)
                                        self.log(f"更新 {date_str} 的交易��料")
                                        connection.commit()
                                        break
                                
                            time.sleep(3)  # 避免請求過快
                            
                        except Exception as e:
                            self.log(f"處理 {date_str} 時發生錯誤: {e}")
                        
                        check_date += timedelta(days=1)
                        self.root.update()  # 更新 GUI
                    
                    self.twse_progress['value'] = i + 1
                    self.twse_label.config(text=f"進度: {i+1}/{total}")
                    
                except Exception as e:
                    self.log(f"更新 {stock_id} 時發生錯誤: {e}")
                    continue
                
            self.log("上市股票資料更新完成")
            
        except Exception as e:
            self.log(f"更新上市股票資料時發生錯誤: {e}")
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()
                
    def update_tpex(self):
        """更新上櫃股票資料"""
        try:
            connection = mysql.connector.connect(**self.db_config)
            cursor = connection.cursor()
            
            # 獲取上櫃股票列表
            cursor.execute("SELECT code, name FROM stock_code WHERE market_type = 'tpex' ORDER BY code")
            stocks = cursor.fetchall()
            
            total = len(stocks)
            self.tpex_progress['maximum'] = total
            self.tpex_progress['value'] = 0
            
            for i, (stock_id, name) in enumerate(stocks):
                try:
                    # 檢查暫停狀態
                    while self.is_paused:
                        self.root.update()
                        time.sleep(0.1)
                        continue
                    
                    self.log(f"正在更新上櫃股票 {stock_id} {name} ({i+1}/{total})")
                    
                    # 取得最近一年的資料
                    current_date = datetime.now()
                    for month in range(12):
                        target_date = current_date - timedelta(days=30 * month)
                        # 轉換日期格式為民國年
                        tw_year = target_date.year - 1911
                        tw_date = f"{tw_year}/{target_date.month:02d}/01"
                        
                        url = f"https://www.tpex.org.tw/web/stock/aftertrading/daily_trading_info/st43_result.php?l=zh-tw&d={tw_date}&stkno={stock_id}"
                        
                        self.log(f"請求 URL: {url}")  # 顯示請求的 URL
                        
                        response = requests.get(url)
                        self.log(f"API 回應狀態碼: {response.status_code}")  # 顯示回應狀態
                        
                        try:
                            data = response.json()
                            self.log(f"API 回應內容: {data}")  # 顯示完整回應
                            
                            if data.get('aaData'):
                                self.log(f"找到 {len(data['aaData'])} 筆資料")
                                self.save_stock_data(stock_id, data['aaData'], 'tpex')
                            else:
                                self.log(f"未找到資料: {data}")
                                
                        except Exception as e:
                            self.log(f"解析 JSON 時發生錯誤: {e}")
                            self.log(f"原始回應: {response.text}")
                        
                        time.sleep(3)  # 避免請求過快
                    
                    self.tpex_progress['value'] = i + 1
                    self.tpex_label.config(text=f"進度: {i+1}/{total}")
                    
                except Exception as e:
                    self.log(f"更新 {stock_id} 時發生錯誤: {e}")
                    continue
                    
            self.log("上櫃股票資料更新完成")
            
        except Exception as e:
            self.log(f"更新上櫃股票資料時發生錯誤: {e}")
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()
                
    def save_stock_data(self, stock_id, data, market_type):
        """保存股票資料到資料庫"""
        try:
            connection = mysql.connector.connect(**self.db_config)
            cursor = connection.cursor()
            
            table_name = f"stock_{stock_id}"
            
            for row in data:
                try:
                    if market_type == 'twse':
                        # 處理上市股票資料
                        date_str = row[0].replace('/', '-')
                        year = int(date_str.split('-')[0]) + 1911
                        date = f"{year}-{date_str.split('-')[1]}-{date_str.split('-')[2]}"
                        
                        # 準備 SQL 參數
                        sql_params = (
                            date,
                            float(row[3].replace(',', '')),  # 開盤價
                            float(row[4].replace(',', '')),  # 最高價
                            float(row[5].replace(',', '')),  # 最低價
                            float(row[6].replace(',', '')),  # 收盤價
                            int(row[1].replace(',', ''))     # 成交量
                        )
                    else:
                        # 處理上櫃股票資料
                        date_parts = row[0].split('/')
                        year = int(date_parts[0]) + 1911
                        date = f"{year}-{date_parts[1]}-{date_parts[2]}"
                        
                        # 準備 SQL 參數
                        sql_params = (
                            date,
                            float(row[3].replace(',', '')),  # 開盤價
                            float(row[4].replace(',', '')),  # 最高價
                            float(row[5].replace(',', '')),  # 最低價
                            float(row[6].replace(',', '')),  # 收盤價
                            int(row[1].replace(',', ''))     # 成交量（股數）
                        )

                    # 建立 SQL 指令
                    sql = f"""
                        INSERT IGNORE INTO {table_name}
                        (date, open_price, high_price, low_price, close_price, volume)
                        VALUES (%s, %s, %s, %s, %s, %s)
                    """
                    
                    # 顯示實際的 SQL 指令
                    actual_sql = sql % sql_params
                    self.log(f"執行 SQL: {actual_sql}")
                    
                    # 執行 SQL
                    cursor.execute(sql, sql_params)
                    
                    # 顯示影響的行數
                    self.log(f"影響的行數: {cursor.rowcount}")
                    
                except Exception as e:
                    self.log(f"處理 {stock_id} 的資料時發生錯誤: {e}")
                    self.log(f"錯誤的資料行: {row}")
                    continue
            
            connection.commit()
            self.log(f"提交完成: {stock_id}")
            
        except Exception as e:
            self.log(f"保存 {stock_id} 資料時發生錯誤: {e}")
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

def main():
    root = tk.Tk()
    app = StockUpdaterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
