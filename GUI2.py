import tkinter as tk
from tkinter import ttk, messagebox
import mysql.connector
import yfinance as yf
import twstock
import requests
from datetime import datetime, timedelta
import time
import pandas as pd

class StockDataUpdater:
    def __init__(self, root):
        self.root = root
        self.root.title("股票資料回補系統")
        self.root.geometry("800x600")
        
        # 創建主框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 創建進度顯示區
        self.progress_frame = ttk.LabelFrame(self.main_frame, text="進度", padding="5")
        self.progress_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # 進度條
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate')
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # 進度標籤
        self.progress_label = ttk.Label(self.progress_frame, text="準備就緒")
        self.progress_label.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=5)
        
        # 創建按鈕框架
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=10)
        
        # 創建資料來源按鈕
        ttk.Button(button_frame, text="使用 twstock 更新", 
                  command=self.update_with_twstock).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="使用 Yahoo 更新", 
                  command=self.update_with_yahoo).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="更新上市股票", 
                  command=self.update_with_official_api).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="更新上櫃股票", 
                  command=self.update_with_tpex_api).pack(side=tk.LEFT, padx=5)
        
        # 添加暫停按鈕
        self.pause_button = ttk.Button(button_frame, text="暫停", 
                                      command=self.toggle_pause)
        self.pause_button.pack(side=tk.LEFT, padx=5)
        
        # 添加暫停狀態標記
        self.is_paused = False
        
        # 創建日誌框架
        log_frame = ttk.LabelFrame(self.main_frame, text="執行日誌", padding="5")
        log_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 日誌文本框
        self.log_text = tk.Text(log_frame, height=20, width=80)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 日誌滾動條
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
    def log(self, message):
        """添加日誌訊息"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert(tk.END, f"[{current_time}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def get_stock_list(self):
        """從資料庫獲取股票列表"""
        try:
            connection = mysql.connector.connect(
                host='192.168.1.145',
                user='root',
                password='fender1106',
                database='stock_code'
            )
            cursor = connection.cursor()
            cursor.execute("SELECT code FROM stock_code ORDER BY code")
            return [row[0] for row in cursor.fetchall()]
        except Exception as e:
            self.log(f"獲取股票列表時發生錯誤: {e}")
            return []
        finally:
            if 'connection' in locals() and connection.is_connected():
                cursor.close()
                connection.close()
                
    def update_with_twstock(self):
        """使用 twstock 更新資料"""
        stocks = self.get_stock_list()
        total = len(stocks)
        self.progress_bar['maximum'] = total
        self.progress_bar['value'] = 0
        
        for i, stock_id in enumerate(stocks):
            try:
                self.log(f"正在更新 {stock_id} 的資料 ({i+1}/{total})")
                stock = twstock.Stock(stock_id)
                data = stock.fetch_from(2010, 1)  # 從2010年開始抓取
                
                if data:
                    self.save_stock_data(stock_id, data, 'twstock')
                
                self.progress_bar['value'] = i + 1
                self.progress_label.config(text=f"進度: {i+1}/{total}")
                time.sleep(3)  # 避免請求過快
                
            except Exception as e:
                self.log(f"更新 {stock_id} 時發生錯誤: {e}")
                
        self.log("使用 twstock 更新完成")
        
    def update_with_yahoo(self):
        """使用 Yahoo Finance 更新資料"""
        stocks = self.get_stock_list()
        total = len(stocks)
        self.progress_bar['maximum'] = total
        self.progress_bar['value'] = 0
        
        for i, stock_id in enumerate(stocks):
            try:
                self.log(f"正在更新 {stock_id} 的資料 ({i+1}/{total})")
                stock = yf.Ticker(f"{stock_id}.TW")
                data = stock.history(period="max")
                
                if not data.empty:
                    self.save_stock_data(stock_id, data, 'yahoo')
                
                self.progress_bar['value'] = i + 1
                self.progress_label.config(text=f"進度: {i+1}/{total}")
                time.sleep(1)  # 避免請求過快
                
            except Exception as e:
                self.log(f"更新 {stock_id} 時發生錯誤: {e}")
                
        self.log("使用 Yahoo Finance 更新完成")
        
    def update_with_official_api(self):
        """使用官方 API 更新資料"""
        stocks = self.get_stock_list()
        total = len(stocks)
        self.progress_bar['maximum'] = total
        self.progress_bar['value'] = 0
        
        for i, stock_id in enumerate(stocks):
            try:
                # 檢查暫停狀態
                while self.is_paused:
                    self.root.update()
                    time.sleep(0.1)
                    if not self.is_paused:  # 如果恢復運行，更新日誌
                        self.log("繼續更新...")
                
                self.log(f"正在更新 {stock_id} 的資料 ({i+1}/{total})")
                # 這裡使用 TWSE API
                url = f"https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date={datetime.now().strftime('%Y%m%d')}&stockNo={stock_id}"
                response = requests.get(url)
                data = response.json()
                
                if data.get('data'):
                    self.save_stock_data(stock_id, data['data'], 'official')
                
                self.progress_bar['value'] = i + 1
                self.progress_label.config(text=f"進度: {i+1}/{total}")
                time.sleep(3)  # 避免請求過快
                
            except Exception as e:
                self.log(f"更新 {stock_id} 時發生錯誤: {e}")
                
        self.log("使用官方 API 更新完成")
        
    def update_with_tpex_api(self):
        """使用櫃買中心 API 更新上櫃股票資料"""
        try:
            connection = mysql.connector.connect(
                host='192.168.1.145',
                user='root',
                password='fender1106',
                database='stock_code'
            )
            cursor = connection.cursor()
            
            # 只獲取上櫃股票
            cursor.execute("SELECT code, name FROM stock_code WHERE market_type = 'tpex' ORDER BY code")
            stocks = cursor.fetchall()
            
            total = len(stocks)
            self.progress_bar['maximum'] = total
            self.progress_bar['value'] = 0
            
            for i, (stock_id, name) in enumerate(stocks):
                try:
                    self.log(f"正在更新上櫃股票 {stock_id} {name} ({i+1}/{total})")
                    
                    # 取得最近一年的資料
                    current_date = datetime.now()
                    for month in range(12):  # 只處理最近12個月
                        # 計算目標日期
                        target_date = current_date - timedelta(days=30 * month)
                        date_str = target_date.strftime('%Y%m%d')
                        
                        # 使用櫃買中心每日收盤資訊API
                        url = f"https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php?l=zh-tw&d={date_str}&stkno={stock_id}"
                        
                        try:
                            response = requests.get(url)
                            data = response.json()
                            
                            if data.get('aaData'):
                                self.save_tpex_daily_data(stock_id, data['aaData'])
                                self.log(f"成功獲取 {stock_id} 在 {date_str} 的資料")
                            
                            # 檢查暫停狀態
                            while self.is_paused:
                                self.root.update()
                                time.sleep(0.1)
                                if not self.is_paused:
                                    self.log("繼續更新...")
                            
                            time.sleep(1)  # 降低請求頻率
                            
                        except Exception as e:
                            self.log(f"獲取 {stock_id} 在 {date_str} 的資料時發生錯誤: {e}")
                            continue
                    
                    self.progress_bar['value'] = i + 1
                    self.progress_label.config(text=f"進度: {i+1}/{total}")
                    self.root.update()
                    
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

    def save_stock_data(self, stock_id, data, source):
        """保存股票資料到資料庫"""
        try:
            connection = mysql.connector.connect(
                host='192.168.1.145',
                user='root',
                password='fender1106',
                database='stock_code'
            )
            cursor = connection.cursor()
            
            # 確保表格存在
            table_name = f"stock_{stock_id}"
            cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS {table_name} (
                    date DATE PRIMARY KEY,
                    open_price FLOAT,
                    close_price FLOAT,
                    high_price FLOAT,
                    low_price FLOAT,
                    volume INT
                )
            """)
            
            # 根據不同來源處理數據格式
            if source == 'twstock':
                for record in data:
                    cursor.execute(f"""
                        INSERT IGNORE INTO {table_name}
                        (date, open_price, close_price, high_price, low_price, volume)
                        VALUES (%s, %s, %s, %s, %s, %s)
                    """, (
                        record.date,
                        record.open,
                        record.close,
                        record.high,
                        record.low,
                        record.volume
                    ))
                    
            elif source == 'yahoo':
                for date, row in data.iterrows():
                    cursor.execute(f"""
                        INSERT IGNORE INTO {table_name}
                        (date, open_price, close_price, high_price, low_price, volume)
                        VALUES (%s, %s, %s, %s, %s, %s)
                    """, (
                        date.date(),
                        row['Open'],
                        row['Close'],
                        row['High'],
                        row['Low'],
                        row['Volume']
                    ))
                    
            elif source == 'official':
                for row in data:
                    # 轉換民國年到西元年
                    date_str = row[0].replace('/', '-')
                    year = int(date_str.split('-')[0]) + 1911
                    date = f"{year}-{date_str.split('-')[1]}-{date_str.split('-')[2]}"
                    
                    cursor.execute(f"""
                        INSERT IGNORE INTO {table_name}
                        (date, open_price, close_price, high_price, low_price, volume)
                        VALUES (%s, %s, %s, %s, %s, %s)
                    """, (
                        date,
                        float(row[3].replace(',', '')),
                        float(row[6].replace(',', '')),
                        float(row[4].replace(',', '')),
                        float(row[5].replace(',', '')),
                        int(row[1].replace(',', ''))
                    ))
            
            connection.commit()
            self.log(f"成功更新 {stock_id} 的資料")
            
        except Exception as e:
            self.log(f"保存 {stock_id} 資料時發生錯誤: {e}")
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def save_tpex_daily_data(self, stock_id, data):
        """保存上櫃股票每日資料"""
        try:
            connection = mysql.connector.connect(
                host='192.168.1.145',
                user='root',
                password='fender1106',
                database='stock_code'
            )
            cursor = connection.cursor()
            
            table_name = f"stock_{stock_id}"
            
            # 確保表格存在
            cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS {table_name} (
                    date DATE PRIMARY KEY,
                    open_price FLOAT,
                    close_price FLOAT,
                    high_price FLOAT,
                    low_price FLOAT,
                    volume INT
                )
            """)
            
            # 處理每一筆資料
            for row in data:
                try:
                    # 轉換日期格式 (格式為 111/01/01)
                    date_parts = row[0].split('/')
                    year = int(date_parts[0]) + 1911  # 民國年轉西元年
                    date = f"{year}-{date_parts[1]}-{date_parts[2]}"
                    
                    # 插入資料
                    cursor.execute(f"""
                        INSERT IGNORE INTO {table_name}
                        (date, open_price, close_price, high_price, low_price, volume)
                        VALUES (%s, %s, %s, %s, %s, %s)
                    """, (
                        date,
                        float(row[1].replace(',', '')),  # 開盤價
                        float(row[2].replace(',', '')),  # 收盤價
                        float(row[3].replace(',', '')),  # 最高價
                        float(row[4].replace(',', '')),  # 最低價
                        int(row[7].replace(',', ''))     # 成交量
                    ))
                    
                except Exception as e:
                    self.log(f"處理資料時發生錯誤: {e}")
                    continue
            
            connection.commit()
            
        except Exception as e:
            self.log(f"保存資料時發生錯誤: {e}")
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def toggle_pause(self):
        """切換暫停狀態"""
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.pause_button.configure(text="繼續")
            self.log("已暫停更新")
        else:
            self.pause_button.configure(text="暫停")
            self.log("繼續更新")

def main():
    root = tk.Tk()
    app = StockDataUpdater(root)
    root.mainloop()

if __name__ == "__main__":
    main()