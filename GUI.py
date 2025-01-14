import tkinter as tk
from tkinter import ttk, messagebox
import mysql.connector
import yfinance as yf
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import matplotlib
matplotlib.use('TkAgg')
# 設置中文字體
plt.rcParams['font.sans-serif'] = ['DFKai-SB']  # 標楷體
plt.rcParams['axes.unicode_minus'] = False  # 解決負號顯示問題
import twstock
from twstock import Stock
from FinMind.data import DataLoader
import time
import webbrowser
import requests
from dotenv import load_dotenv
import os

# 載入環境變數
load_dotenv()

# 資料庫設定
DB_CONFIG = {
    'host': os.getenv('DB_HOST'),
    'user': os.getenv('DB_USER'),
    'password': os.getenv('DB_PASSWORD'),
    'database': os.getenv('DB_DATABASE')
}

def get_week_first_day_indices(dates):
    """獲取每週第一個交易日的索引"""
    indices = []
    current_week = -1
    for i, date in enumerate(dates):
        week_number = date.isocalendar()[1]  # 獲取週數
        if week_number != current_week:
            indices.append(i)
            current_week = week_number
    return indices

class StockApp:
    def __init__(self, root):
        self.root = root
        self.root.title("股票查詢系統")
        
        # 獲取螢幕尺寸
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        window_width = int(screen_width * 0.8)
        window_height = int(screen_height * 0.8)
        
        # 設置主視窗大小
        self.root.geometry(f"{window_width}x{window_height}")
        
        # 建立主框架 (只建立一次)
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        root.grid_rowconfigure(0, weight=1)
        root.grid_columnconfigure(0, weight=1)
        
        # 建立左側選單框架
        self.menu_frame = ttk.Frame(self.main_frame)
        self.menu_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        # 建立功能選擇下拉選單
        self.function_var = tk.StringVar()
        self.function_choices = [
            "所有股票",
            "上市股票",
            "上櫃股票",
            "波動30%以上股票",
            "布林帶排序",
            "追蹤自選股"
        ]
        self.function_var.set(self.function_choices[0])
        function_menu = ttk.OptionMenu(
            self.menu_frame, 
            self.function_var, 
            self.function_choices[0], 
            *self.function_choices,
            command=self.on_function_change
        )
        function_menu.pack(fill=tk.X, pady=5)
        
        # 建立股票列表
        columns = ('代碼', '名稱', '市場', '資訊')
        self.stock_tree = ttk.Treeview(
            self.menu_frame, 
            columns=columns,
            show='headings',
            height=20
        )
        
        # 設置列標題
        for col in columns:
            self.stock_tree.heading(col, text=col)
            self.stock_tree.column(col, width=100)
        
        # 加入垂直捲動條
        scrollbar = ttk.Scrollbar(
            self.menu_frame, 
            orient=tk.VERTICAL, 
            command=self.stock_tree.yview
        )
        self.stock_tree.configure(yscrollcommand=scrollbar.set)
        
        # 放置 Treeview 和捲動條
        self.stock_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 建立右側圖表框架
        self.chart_frame = ttk.Frame(self.main_frame)
        self.chart_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        
        # 設置列和行的權重
        self.main_frame.grid_columnconfigure(1, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)
        
        # 載入股票列表
        self.load_stock_list()
        
        # 添加窗關事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 創建主框架並設置權重
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        root.grid_rowconfigure(0, weight=1)
        root.grid_columnconfigure(0, weight=1)
        
        # 創建輸入框架
        input_frame = ttk.Frame(self.main_frame)
        input_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # 股票代碼輸入
        ttk.Label(input_frame, text="請輸入股票代碼：").pack(side=tk.LEFT, padx=5)
        self.code_entry = ttk.Entry(input_frame, width=10)
        self.code_entry.pack(side=tk.LEFT, padx=5)
        self.code_entry.bind('<Return>', lambda event: self.search_stock())
        
        # 天數輸入
        ttk.Label(input_frame, text="查詢天數：").pack(side=tk.LEFT, padx=5)
        self.days_entry = ttk.Entry(input_frame, width=10)
        self.days_entry.pack(side=tk.LEFT, padx=5)
        self.days_entry.insert(0, "300")  # 修改預設值為 300
        self.days_entry.bind('<Return>', lambda event: self.search_stock())
        
        # 查詢按鈕
        ttk.Button(input_frame, text="查詢", command=self.search_stock).pack(side=tk.LEFT, padx=5)
        ttk.Button(input_frame, text="加入追蹤", command=self.add_to_track).pack(side=tk.LEFT, padx=5)
        
        # 建框架
        self.chart_frame = ttk.Frame(self.main_frame)
        self.chart_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        # 創建圖表，使用固定的高度比例
        window_height = self.root.winfo_height()
        chart_height = window_height * 0.85  # 使用視窗高度的85%
        self.fig, (self.ax1, self.ax2, self.ax3, self.ax4) = plt.subplots(4, 1, 
            figsize=(self.root.winfo_screenwidth()/100, chart_height/100),
            height_ratios=[3, 1, 1, 1])  # 添加第四個子圖，調整比例
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.chart_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 綁定視窗大小改變事件
        self.root.bind('<Configure>', self.on_resize)
        
        # 數據顯示標籤
        self.info_label = ttk.Label(self.main_frame, text="", justify=tk.LEFT, font=('Arial', 12))
        self.info_label.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # 添加用於存當前數據的屬性
        self.current_data = None
        # 添加用於高亮顯示的屬性
        self.highlight_bar = None
        self.vertical_line = None
        
        # 連接動
        self.canvas.mpl_connect('motion_notify_event', self.on_mouse_move)
        
        # 添加的屬性來儲背景圖
        self.background = None
        # 添加新的屬性來存儲滑鼠位置標記
        self.cursor_annotation = None
        
        # 綁定鍵盤事件
        self.root.bind('<Prior>', self.previous_stock)  # Page Up
        self.root.bind('<Next>', self.next_stock)      # Page Down
        
        # 創建輪詢按鈕框架（放在右下角）
        self.polling_frame = ttk.Frame(self.main_frame)
        self.polling_frame.grid(row=4, column=0, sticky=tk.E, pady=5)
        
        # 創建輪詢按鈕
        self.polling_button = ttk.Button(self.polling_frame, text="開始輪詢", command=self.toggle_polling)
        self.polling_button.pack(side=tk.RIGHT, padx=5)
        
        # 添加輪詢狀態標記
        self.is_polling = False
        self.current_polling_index = 0
        
        # 建進度條框架（在股票列表下方）
        self.progress_frame = ttk.Frame(self.menu_frame)
        self.progress_frame.pack(fill=tk.X, pady=5)
        
        # 綁定點��
        self.stock_tree.bind('<Double-1>', self.on_stock_select)
        
    def check_stock_code(self, code):
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            # 先查股票是否存在
            cursor.execute("SELECT code, name FROM stock_code WHERE code = %s", (code,))
            result = cursor.fetchone()
            
            if result:
                # 如果存在但沒有名，則嘗試更新
                if result[1] is None or result[1] == '':
                    try:
                        stock = yf.Ticker(f"{code}.TW")
                        stock_name = stock.info['longName']  # 使用個方法取股票名稱
                        cursor.execute("UPDATE stock_code SET name = %s WHERE code = %s", 
                                     (stock_name, code))
                        connection.commit()
                        return {'code': code, 'name': stock_name}
                    except Exception as e:
                        print(f"更新股票名稱時出錯: {e}")
                        return {'code': code, 'name': ''}
                return {'code': result[0], 'name': result[1]}
            else:
                # 如果股票不存在，則嘗試獲取名稱並新增
                try:
                    stock = yf.Ticker(f"{code}.TW")
                    stock_name = stock.info['longName']  # 使用這個方法獲取股票名稱
                    cursor.execute("INSERT INTO stock_code (code, name) VALUES (%s, %s)", 
                                 (code, stock_name))
                    connection.commit()
                    return {'code': code, 'name': stock_name}
                except Exception as e:
                    print(f"獲取股票名稱時出錯: {e}")
                    return {'code': code, 'name': ''}
            
        except mysql.connector.Error as err:
            messagebox.showerror("錯誤", f"數據庫連接錯誤: {err}")
            return None
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def save_stock_code(self, code, name):
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            cursor.execute("INSERT INTO stock_code (code, name) VALUES (%s, %s)", 
                          (code, name))
            connection.commit()
        except mysql.connector.Error as err:
            messagebox.showerror("錯誤", f"數據庫連接錯誤: {err}")
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def draw_candlestick(self, ax, dates, opens, highs, lows, closes):
        width = 0.6
        
        for i in range(len(dates)):
            # 判斷是上漲還是下跌
            if closes[i] >= opens[i]:
                # 上漲，畫紅色蠟
                color = 'red'
                body_bottom = opens[i]
                body_height = closes[i] - opens[i]
            else:
                # 下跌，畫綠色蠟燭
                color = 'green'
                body_bottom = closes[i]
                body_height = opens[i] - closes[i]
            
            # 繪製蠟燭主體
            ax.bar(dates[i], body_height, width, bottom=body_bottom, color=color, alpha=0.7)
            
            # 繪製上下影線
            ax.plot([dates[i], dates[i]], [lows[i], highs[i]], color=color, linewidth=1)

    def on_mouse_move(self, event):
        if event.inaxes != self.ax1 or self.current_data is None:
            return
            
        # 獲接近的數據點
        x_index = int(round(event.xdata))
        if 0 <= x_index < len(self.current_data):
            # 清除之前的高亮和垂直線
            if self.highlight_bar is not None:
                try:
                    for bar in self.highlight_bar:
                        if isinstance(bar, matplotlib.container.BarContainer):
                            for b in bar:
                                b.remove()
                        else:
                            bar.remove()
                except:
                    pass
                self.highlight_bar = None
            
            if self.vertical_line is not None:
                try:
                    for line in self.vertical_line:
                        line.remove()
                except:
                    pass
                self.vertical_line = None
            
            # 獲取前K線數據
            data = self.current_data[x_index]
            date = data['date']
            open_price = data['open']
            close_price = data['close']
            high_price = data['high']
            low_price = data['low']
            volume = data['volume']
            
            # 更新數據顯示標籤
            info = f'日期: {date.strftime("%Y-%m-%d")}  開盤: {open_price:.2f}  收盤: {close_price:.2f}  最高: {high_price:.2f}  最低: {low_price:.2f}  成交量: {volume}'
            self.info_label.config(text=info)
            
            # 繪製高亮K線
            if close_price >= open_price:
                color = 'pink'
            else:
                color = 'lightgreen'
            
            # 繪製高亮K線
            body_bottom = min(open_price, close_price)
            body_height = abs(close_price - open_price)
            bar = self.ax1.bar(x_index, body_height, width=0.6, bottom=body_bottom, color=color, alpha=0.7)
            line = self.ax1.vlines(x_index, low_price, high_price, color=color, linewidth=2)
            
            # 製直線
            vline1 = self.ax1.axvline(x=x_index, color='gray', linestyle='--', alpha=0.5)
            vline2 = self.ax2.axvline(x=x_index, color='gray', linestyle='--', alpha=0.5)
            vline3 = self.ax3.axvline(x=x_index, color='gray', linestyle='--', alpha=0.5)
            
            self.highlight_bar = [bar, line]
            self.vertical_line = [vline1, vline2, vline3]
            
            # 更新表
            self.canvas.draw()

    def save_daily_data(self, code, data):
        """保存每日資料到資料庫"""
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            # 檢查並創建表格（如果不存在）
            table_name = f"stock_{code}"
            cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS {table_name} (
                    date date NOT NULL COMMENT '日期',
                    open_price decimal(10,2) DEFAULT NULL COMMENT '開盤價',
                    high_price decimal(10,2) DEFAULT NULL COMMENT '最高價',
                    low_price decimal(10,2) DEFAULT NULL COMMENT '最低價',
                    close_price decimal(10,2) DEFAULT NULL COMMENT '收盤價',
                    volume bigint(20) DEFAULT NULL COMMENT '成交量',
                    PRIMARY KEY (date)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci
            """)
            
            # 插入資料，如果日期已存在則跳過
            for d in data:
                cursor.execute(f"""
                    INSERT IGNORE INTO {table_name} 
                    (date, open_price, high_price, low_price, close_price, volume)
                    VALUES (%s, %s, %s, %s, %s, %s)
                """, (
                    d.date,
                    float(d.open),
                    float(d.high),
                    float(d.low),
                    float(d.close),
                    int(d.transaction)  # 成交量
                ))
            
            connection.commit()
            print(f"已保存 {code} 的交易數據")
            
        except mysql.connector.Error as err:
            print(f"數據庫錯誤: {err}")
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def search_stock(self):
        code = self.code_entry.get().strip()
        if not code.isdigit() or len(code) != 4:
            messagebox.showerror("錯誤", "請輸入4位數字的股票代碼")
            return
            
        try:
            # 先從資料獲取股票市場類型
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            cursor.execute("SELECT market_type FROM stock_code WHERE code = %s", (code,))
            result = cursor.fetchone()
            
            if result:
                market_type = result[0]
            else:
                market_type = 'unknown'
            
            cursor.close()
            connection.close()

            # 使用 FinMind 獲取股票數據
            dl = DataLoader()
            days = int(self.days_entry.get().strip())
            end_date = datetime.now()
            start_date = end_date - timedelta(days=days * 1.5)
            
            print(f"開始獲取股票 {code} 的數據...")
            df = dl.taiwan_stock_daily(
                stock_id=code,
                start_date=start_date.strftime('%Y-%m-%d'),
                end_date=end_date.strftime('%Y-%m-%d')
            )
            
            if df is None or df.empty:
                messagebox.showerror("誤", "無法獲取股票數據")
                return
            
            # 轉換數據格式
            data = []
            for _, row in df.iterrows():
                try:
                    data.append(type('Data', (), {
                        'date': datetime.strptime(row['date'], '%Y-%m-%d'),
                        'open': float(row['open']),
                        'high': float(row['max']),
                        'low': float(row['min']),
                        'close': float(row['close']),
                        'transaction': int(row['Trading_Volume']) // 1000  # 成交除以1000
                    }))
                except Exception as e:
                    print(f"處理數據時發生錯誤: {e}")
                    continue
            
            # 檢查是否成功獲取數據
            if not data:
                messagebox.showerror("錯誤", "無法獲取股票數據")
                return
            
            # 排序並只取最近的指定天數
            data = sorted(data, key=lambda x: x.date)[-days:]
            
            # 保存每日交易數據
            self.save_daily_data(code, data)
            
            # 準備數據
            dates = [d.date for d in data]
            opens = [d.open for d in data]
            closes = [d.close for d in data]
            highs = [d.high for d in data]
            lows = [d.low for d in data]
            volumes = [d.transaction for d in data]
            
            # 轉換為 pandas Series 進計算
            closes_series = pd.Series(closes)
            
            # 計算移動平線
            ma5 = closes_series.rolling(window=5).mean()
            ma10 = closes_series.rolling(window=10).mean()
            ma20 = closes_series.rolling(window=20).mean()
            ma21 = closes_series.rolling(window=21).mean()
            
            # 計算布林通道
            std20 = closes_series.rolling(window=20).std()
            upper_band = ma20 + 2 * std20
            lower_band = ma20 - 2 * std20
            
            # 計算乖離率
            bias = ((closes_series - ma20) / ma20) * 100
            
            # 獲取票名稱
            stock_name = twstock.codes[code].name if code in twstock.codes else "未知"
            
            # 獲取當前實際的視窗大小
            self.root.update_idletasks()  # 確保視窗尺寸已更新
            window_width = self.root.winfo_width()
            window_height = self.root.winfo_height()
            chart_height = window_height * 0.85
            
            # 計算圖表實際可用寬度（減去右側功能表的寬度）
            chart_width = window_width - 750 - 30  # 減去功能表寬度(750)和一些邊距(30)
            
            # 重新設置表大小
            self.fig.set_size_inches(chart_width/100, chart_height/100)
            
            # 清除現有圖表
            self.ax1.clear()
            self.ax2.clear()
            self.ax3.clear()
            self.ax4.clear()
            
            # 設置圖表背景顏色
            self.fig.patch.set_facecolor('black')
            self.ax1.set_facecolor('black')
            self.ax2.set_facecolor('black')
            self.ax3.set_facecolor('black')
            self.ax4.set_facecolor('black')
            
            # 設置軸線顏色
            for ax in [self.ax1, self.ax2, self.ax3, self.ax4]:
                ax.spines['bottom'].set_color('white')
                ax.spines['top'].set_color('white')
                ax.spines['left'].set_color('white')
                ax.spines['right'].set_color('white')
                ax.tick_params(axis='both', colors='white')  # 設置度標籤顏色
            
            # 設置網格色
            self.ax1.grid(True, color='gray', alpha=0.2)
            self.ax2.grid(True, color='gray', alpha=0.2)
            self.ax3.grid(True, color='gray', alpha=0.2)
            self.ax4.grid(True, color='gray', alpha=0.2)
            
            # 置標題
            self.ax1.set_title(f'{code} {stock_name} {market_type}', color='white', fontsize=12, pad=20)
            self.ax2.set_title('成交量', color='white', fontsize=10, pad=20)
            self.ax3.set_title('乖離率%', color='white', fontsize=10, pad=20)
            
            # 設週期性的 x 軸刻度
            week_indices = get_week_first_day_indices(dates)
            
            # 繪製 K 線圖
            x_range = range(len(dates))
            for i in x_range:
                try:
                    # 檢查數據是否有效
                    if opens[i] is None or closes[i] is None or highs[i] is None or lows[i] is None:
                        continue  # 跳無效數據
                        
                    if closes[i] >= opens[i]:
                        color = 'red'
                    else:
                        color = 'green'
                    
                    # 繪製 K 線實體
                    self.ax1.bar(i, abs(closes[i] - opens[i]), 0.6,
                                 bottom=min(opens[i], closes[i]),
                                 color=color, alpha=0.7)
                    # 繪製上下影線
                    self.ax1.plot([i, i], [lows[i], highs[i]], color=color, linewidth=1)
                except Exception as e:
                    print(f"繪製第 {i} 根 K 時發生錯誤: {e}")
                    continue
            
            # 繪製移動平均線
            self.ax1.plot(x_range, ma5, color='yellow', label='MA5', linewidth=1)
            self.ax1.plot(x_range, ma10, color='purple', label='MA10', linewidth=1)
            self.ax1.plot(x_range, ma20, color='orange', label='MA20', linewidth=1)
            self.ax1.plot(x_range, ma21, color='blue', label='MA21', linewidth=1)
            
            # 繪製布林通道
            self.ax1.plot(x_range, upper_band, '--', color='gray', label='Upper Band', alpha=0.5)
            self.ax1.plot(x_range, lower_band, '--', color='gray', label='Lower Band', alpha=0.5)
            
            # 繪製成交量圖
            self.ax2.bar(x_range, volumes, color='gray', alpha=0.7)
            
            # 計算乖離率（股價與移動平均線的偏離度）
            ma20 = closes_series.rolling(window=20).mean()  # 20日移動平線
            bias = ((closes_series - ma20) / ma20) * 100  # 乖離率計算公式
            
            # 設置所有子圖的 x 軸範圍
            x_range = range(len(dates))
            x_lim = (-1, len(dates))  # 統一 x 範圍

            # 設每個子圖的範圍
            self.ax1.set_xlim(x_lim)
            self.ax2.set_xlim(x_lim)
            self.ax3.set_xlim(x_lim)
            self.ax4.set_xlim(x_lim)

            # 繪製乖離率圖
            self.ax3.plot(x_range, bias, color='blue', label='乖離率%')
            self.ax3.fill_between(x_range, bias, 0, 
                                 where=(bias >= 0), color='red', alpha=0.3)  # 正乖離填充紅色
            self.ax3.fill_between(x_range, bias, 0, 
                                 where=(bias <= 0), color='green', alpha=0.3)  # 負乖離填充綠色
            
            # 添加水平參考線（保在設置 x 軸範圍後添加）
            for y in [10, -10, 0]:
                line = self.ax3.axhline(y=y, color='gray', 
                                       linestyle='--' if y != 0 else '-', 
                                       alpha=0.5)
                line.set_xdata([x_lim[0], x_lim[1]])  # 確水平線跨越整個圖表
            
            # 設置圖例
            self.ax1.legend(loc='best')
            self.ax3.legend(loc='best')
            
            # 設置 x 軸標籤
            self.ax1.set_xticks(week_indices)
            self.ax1.set_xticklabels([dates[i].strftime('%Y-%m-%d') for i in week_indices], 
                            rotation=45, color='white')
            
            # 設置其他子圖的 x 軸（不顯示標籤）
            self.ax2.set_xticks(week_indices)
            self.ax2.set_xticklabels([''] * len(week_indices))
            self.ax3.set_xticks(week_indices)
            self.ax3.set_xticklabels([''] * len(week_indices))
            self.ax4.set_xticks(week_indices)
            self.ax4.set_xticklabels([''] * len(week_indices))
            
            # 計算 MACD
            exp1 = closes_series.ewm(span=12, adjust=False).mean()
            exp2 = closes_series.ewm(span=26, adjust=False).mean()
            macd = exp1 - exp2
            signal = macd.ewm(span=9, adjust=False).mean()
            histogram = macd - signal
            
            # 繪製 MACD 圖
            self.ax4.plot(x_range, macd, label='MACD', color='white', linewidth=1)
            self.ax4.plot(x_range, signal, label='Signal', color='yellow', linewidth=1)
            
            # 繪製 MACD 柱狀圖
            for i in range(len(histogram)):
                if histogram[i] >= 0:
                    color = 'red'
                else:
                    color = 'green'
                self.ax4.bar(i, histogram[i], color=color, alpha=0.7)
            
            # 添加零線
            self.ax4.axhline(y=0, color='gray', linestyle='-', alpha=0.5)
            
            # 設置 MACD 圖的標題和圖例
            self.ax4.set_title('MACD', color='white', fontsize=10, pad=20)
            self.ax4.legend(loc='best')
            
            # 設置 MACD 圖的背景和網格
            self.ax4.set_facecolor('black')
            self.ax4.grid(True, color='gray', alpha=0.2)
            
            # 設置 MACD 圖的軸線和刻度顏色
            self.ax4.spines['bottom'].set_color('white')
            self.ax4.spines['top'].set_color('white')
            self.ax4.spines['left'].set_color('white')
            self.ax4.spines['right'].set_color('white')
            self.ax4.tick_params(axis='both', colors='white')
            
            # 設置 x 軸
            self.ax4.set_xlim(x_lim)
            self.ax4.set_xticks(week_indices)
            self.ax4.set_xticklabels([''] * len(week_indices))
            
            # 調布局
            self.fig.tight_layout()
            
            # 強制更新畫布大小
            self.canvas.get_tk_widget().configure(width=chart_width, height=chart_height)
            
            # 更新圖表
            self.canvas.draw()
            
            # 在完成所有圖繪製後，保存當前數據和背景圖
            self.current_data = [
                {
                    'date': dates[i],
                    'open': opens[i],
                    'close': closes[i],
                    'high': highs[i],
                    'low': lows[i],
                    'volume': volumes[i]
                }
                for i in range(len(dates))
            ]
            
            self.background = self.fig.canvas.copy_from_bbox(self.fig.bbox)
            
        except Exception as e:
            messagebox.showerror("錯誤", f"獲取股票數據時發生錯誤: {str(e)}")

    def on_closing(self):
        """處理視窗關閉事件"""
        plt.close('all')
        self.root.quit()
        self.root.destroy()

    def get_next_stock_code(self, current_code, direction='next'):
        """獲取一個或上一股票代碼"""
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            if direction == 'next':
                # 獲取大於當前代碼的最小代碼
                cursor.execute("""
                    SELECT code FROM stock_code 
                    WHERE code > %s 
                    ORDER BY code ASC 
                    LIMIT 1
                """, (current_code,))
            else:
                # 獲取小於當前代碼的最大代碼
                cursor.execute("""
                    SELECT code FROM stock_code 
                    WHERE code < %s 
                    ORDER BY code DESC 
                    LIMIT 1
                """, (current_code,))
            
            result = cursor.fetchone()
            return result[0] if result else None
            
        except mysql.connector.Error as err:
            messagebox.showerror("錯誤", f"數據庫連接錯誤: {err}")
            return None
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def previous_stock(self, event):
        """切換到上一個股票"""
        # 獲取當前股票列表中的所有項目
        all_items = self.stock_tree.get_children()
        if not all_items:
            return
        
        # 獲取當前股票代碼
        current_code = self.code_entry.get().strip()
        
        # 在列表中找到當前股票的位置
        current_index = -1
        for i, item in enumerate(all_items):
            values = self.stock_tree.item(item)['values']
            if str(values[0]) == current_code:
                current_index = i
                break
        
        # 如果找到當前股票，移動到上一個
        if current_index > 0:
            prev_item = all_items[current_index - 1]
            prev_values = self.stock_tree.item(prev_item)['values']
            prev_code = str(prev_values[0])
            
            # 設置新的股票代碼並執行查詢
            self.code_entry.delete(0, tk.END)
            self.code_entry.insert(0, prev_code)
            self.search_stock()
            
            # 選中並滾動到新的項目
            self.stock_tree.selection_set(prev_item)
            self.stock_tree.see(prev_item)

    def next_stock(self, event):
        """切換到下一個股票"""
        # 獲取當前股票列表中的所有項目
        all_items = self.stock_tree.get_children()
        if not all_items:
            return
        
        # 獲取當前股票代碼
        current_code = self.code_entry.get().strip()
        
        # 在列表中找到當前股票的位置
        current_index = -1
        for i, item in enumerate(all_items):
            values = self.stock_tree.item(item)['values']
            if str(values[0]) == current_code:
                current_index = i
                break
        
        # 如果找到當前股票，移動到下一個
        if current_index >= 0 and current_index < len(all_items) - 1:
            next_item = all_items[current_index + 1]
            next_values = self.stock_tree.item(next_item)['values']
            next_code = str(next_values[0])
            
            # 設置新的股票代碼並執行查詢
            self.code_entry.delete(0, tk.END)
            self.code_entry.insert(0, next_code)
            self.search_stock()
            
            # 選中並滾動到新的項目
            self.stock_tree.selection_set(next_item)
            self.stock_tree.see(next_item)

    def on_resize(self, event):
        """處理視窗大小改變件"""
        # 處理視窗的大小改變
        if event.widget == self.root:
            # 計算新的圖表高度
            window_height = event.height
            chart_height = window_height * 0.85  # 使用視窗高度的85%
            
            # 更新圖表大小
            width = event.width / 100
            height = chart_height / 100
            self.fig.set_size_inches(width, height)
            
            # 重新調整布
            self.fig.tight_layout()
            self.canvas.draw()

    def toggle_polling(self):
        """切換輪詢狀態"""
        if not self.is_polling:
            # 開始輪詢
            self.is_polling = True
            self.polling_button.configure(text="停止詢")
            self.start_polling()
        else:
            # 停止輪詢
            self.is_polling = False
            self.polling_button.configure(text="開始輪詢")

    def start_polling(self):
        """始詢所有股票"""
        if not self.is_polling:
            return
        
        try:
            # 連接數據庫獲取所有股票代碼
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            cursor.execute("SELECT code FROM stock_code ORDER BY code")
            all_codes = [row[0] for row in cursor.fetchall()]
            
            if self.current_polling_index >= len(all_codes):
                # 完成一輪輪詢，重置索引
                self.current_polling_index = 0
                print("完成一輪輪詢，開始一輪...")
            
            # 獲取當前要詢的股票代碼
            current_code = all_codes[self.current_polling_index]
            print(f"正在輪詢股票: {current_code}")
            
            # 設置股票代碼並執行查詢
            self.code_entry.delete(0, tk.END)
            self.code_entry.insert(0, current_code)
            self.search_stock()
            
            # 更新索引並設置下一次查詢
            self.current_polling_index += 1
            self.root.after(3000, self.start_polling)  # 3秒後查下一個
            
        except mysql.connector.Error as err:
            print(f"數據庫錯誤: {err}")
            self.is_polling = False
            self.polling_button.configure(text="開始輪詢")
        finally:
            if 'connection' in locals() and connection.is_connected():
                cursor.close()
                connection.close()

    def load_stock_list(self):
        """從數據庫載入股票列表"""
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            # 清空現有項目
            for item in self.stock_tree.get_children():
                self.stock_tree.delete(item)
            
            # 獲取所有股票資訊
            cursor.execute("""
                SELECT code, name, market_type 
                FROM stock_code 
                WHERE market_type IS NOT NULL 
                ORDER BY code
            """)
            stocks = cursor.fetchall()
            
            # 插入數據
            for stock in stocks:
                code, name, market = stock
                market_display = '上市' if market == 'twse' else '上櫃' if market == 'tpex' else market
                self.stock_tree.insert('', tk.END, values=(code, name, market_display, ''))
                
        except mysql.connector.Error as err:
            print(f"數據庫錯誤: {err}")
        finally:
            if 'connection' in locals() and connection.is_connected():
                cursor.close()
                connection.close()

    def on_stock_select(self, event):
        """處理股票選事"""
        selection = self.stock_tree.selection()
        if selection:
            item = self.stock_tree.item(selection[0])
            code = str(item['values'][0])  # 轉換為字符串
            # 確保股票代碼為4位數
            code = code.zfill(4)  # 補零到4位
            # 設置股票碼並執行查
            self.code_entry.delete(0, tk.END)
            self.code_entry.insert(0, code)
            self.search_stock()

    def calculate_volatility(self, stock_code):
        """計算股票波動率"""
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            # 獲取最近30天的收盤價日期
            table_name = f"stock_{stock_code}"
            cursor.execute(f"""
                SELECT date, close_price, volume 
                FROM {table_name} 
                ORDER BY date DESC 
                LIMIT 30
            """)
            
            rows = cursor.fetchall()
            if len(rows) < 2:  # 確保有足夠的數據
                return 0
            
            # 檢查最近一天的成交量
            latest_date = rows[0][0]
            latest_volume = rows[0][2]
            if latest_volume is None or latest_volume < 100:  # 成量過小的跳過
                return 0
            
            # 取出收盤價和最早的日期
            prices = [row[1] for row in rows if row[1] is not None]
            earliest_date = rows[-1][0]
            
            # 計算波動率 (最高價 - 最低價) / 最低價 * 100
            max_price = max(prices)
            min_price = min(prices)
            
            # 防止除以零的情況
            if min_price == 0 or min_price is None:
                return 0
            
            try:
                volatility = ((max_price - min_price) / min_price) * 100
                return {
                    'volatility': volatility,
                    'start_date': earliest_date,
                    'end_date': latest_date,
                    'max_price': max_price,
                    'min_price': min_price,
                    'volume': latest_volume
                }
            except ZeroDivisionError:
                return 0
            
        except mysql.connector.Error as err:
            print(f"計算波動率時發生錯誤 {stock_code}: {err}")
            return 0
        except Exception as e:
            print(f"計算 {stock_code} 波動率時發生預期的錯誤: {e}")
            return 0
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def calculate_bollinger(self, stock_code):
        """計算布林帶指標"""
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            # 獲取最近20天的收盤價
            table_name = f"stock_{stock_code}"
            cursor.execute(f"""
                SELECT close_price 
                FROM {table_name} 
                ORDER BY date DESC 
                LIMIT 20
            """)
            
            prices = [row[0] for row in cursor.fetchall() if row[0] is not None]
            if len(prices) < 20:  # 確保有足夠的數據
                return None
            
            # 計算布林帶指標
            prices = prices[::-1]  # 反轉順序以便計算
            ma20 = sum(prices) / len(prices)
            std = (sum((x - ma20) ** 2 for x in prices) / len(prices)) ** 0.5
            
            upper_band = ma20 + 2 * std
            lower_band = ma20 - 2 * std
            current_price = prices[-1]
            
            # 計算股價在布林帶中的位置 (百分)
            band_width = upper_band - lower_band
            if band_width == 0:
                return None
            
            position = ((current_price - lower_band) / band_width) * 100
            
            return {
                'position': position,
                'current_price': current_price,
                'upper_band': upper_band,
                'lower_band': lower_band
            }
            
        except mysql.connector.Error as err:
            print(f"計算布林帶時發生錯誤 {stock_code}: {err}")
            return None
        except Exception as e:
            print(f"計算 {stock_code} 布林帶時生未預期的錯誤: {e}")
            return None
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def calculate_bias(self, stock_code):
        """計算乖離率"""
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            # 獲取最近一個交易日的成交量和最近20天的收盤價
            table_name = f"stock_{stock_code}"
            cursor.execute(f"""
                SELECT close_price, volume 
                FROM {table_name} 
                ORDER BY date DESC 
                LIMIT 20
            """)
            
            rows = cursor.fetchall()
            if len(rows) < 20:  # 確保有足夠的數據
                return None
            
            # 檢查最近一天的成交量
            latest_volume = rows[0][1]  # 第一筆資料的成交量
            if latest_volume is None or latest_volume < 100:  # 成交量小於100的股票跳過
                return None
            
            # 取出收盤價
            prices = [row[0] for row in rows if row[0] is not None]
            if len(prices) < 20:
                return None
            
            # 計算20日均線
            ma20 = sum(prices) / len(prices)
            current_price = prices[0]  # 最新格
            
            # 計算乖離率
            if ma20 == 0:
                return None
            
            bias = ((current_price - ma20) / ma20) * 100
            
            return {
                'bias': bias,
                'current_price': current_price,
                'ma20': ma20,
                'volume': latest_volume
            }
            
        except mysql.connector.Error as err:
            print(f"計算乖離率時發生錯誤 {stock_code}: {err}")
            return None
        except Exception as e:
            print(f"計算 {stock_code} 乖離率時發生未預期的錯誤: {e}")
            return None
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def get_convertible_bonds(self):
        """獲取可轉債股票"""
        try:
            # 使用 FinMind 獲取可轉債資料
            dl = DataLoader()
            # 獲取台灣可轉債資訊
            df = dl.taiwan_stock_convertible_bond_info()
            
            # 檢查數據是否為空
            if df is None or df.empty:
                print("未獲取到可轉債資料")
                return []
            
            # 轉換成需要的格式
            bonds = []
            for _, row in df.iterrows():
                try:
                    cb_code = row['cb_id']  # 可轉債代碼
                    cb_name = row['cb_name']  # 可轉債名稱
                    start_date = row['InitialDateOfConversion']  # 上市日期
                    due_date = row['DueDateOfConversion']  # 到期日
                    
                    # 從可轉債名稱中提取標的股票代碼
                    stock_code = cb_code[:4]  # 取4碼作為標的票代碼
                    
                    # 從資料庫獲取標的股票資訊
                    try:
                        connection = mysql.connector.connect(**DB_CONFIG)
                        cursor = connection.cursor()
                        
                        cursor.execute("SELECT name, market_type FROM stock_code WHERE code = %s", (stock_code,))
                        result = cursor.fetchone()
                        
                        if result:
                            stock_name = result[0]
                            market_type = result[1]
                            market_display = '上市' if market_type == 'twse' else '上櫃'
                            
                            # 添加未到期的可轉債
                            due_date_obj = datetime.strptime(due_date, '%Y-%m-%d')
                            if due_date_obj > datetime.now():
                                bonds.append((
                                    stock_code,  # 標的股票代碼
                                    f"{stock_name} - {cb_name}",  # 標的股票稱 + 可轉債名稱
                                    market_display,  # 市場別
                                    f"上市:{start_date} 到期:{due_date}"  # 上市日期和到期日
                                ))
                                
                    except mysql.connector.Error as err:
                        print(f"查詢標的股票資訊發生錯誤 {stock_code}: {err}")
                    finally:
                        if 'connection' in locals() and connection.is_connected():
                            cursor.close()
                            connection.close()
                        
                except Exception as e:
                    print(f"處理可轉債資料時發生錯誤: {e}")
                    continue
            
            return bonds
            
        except Exception as e:
            print(f"獲取可轉債時發生錯誤: {e}")
            return []

    def on_function_change(self, event):
        try:
            selected_function = self.function_var.get()
            
            # 清空現有的樹形視圖
            for item in self.stock_tree.get_children():
                self.stock_tree.delete(item)
            
            # 連接數據庫
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            try:
                if selected_function == "上市股票":
                    cursor.execute("""
                        SELECT code, name, market_type 
                        FROM stock_code 
                        WHERE market_type = 'twse' 
                        ORDER BY code
                    """)
                    stocks = cursor.fetchall()
                    for stock in stocks:
                        code, name, market_type = stock
                        self.stock_tree.insert('', tk.END, values=(code, name, '上市', ''))
                        
                elif selected_function == "上櫃股票":
                    cursor.execute("""
                        SELECT code, name, market_type 
                        FROM stock_code 
                        WHERE market_type = 'tpex' 
                        ORDER BY code
                    """)
                    stocks = cursor.fetchall()
                    for stock in stocks:
                        code, name, market_type = stock
                        self.stock_tree.insert('', tk.END, values=(code, name, '上櫃', ''))
                        
                elif selected_function == "波動30%以上股票":
                    cursor.execute("SELECT code, name, market_type FROM stock_code ORDER BY code")
                    stocks = cursor.fetchall()
                    
                    # 在進度條框架中創建進度條
                    progress = ttk.Progressbar(self.progress_frame, mode='determinate')
                    progress.pack(fill=tk.X, padx=5)
                    progress['maximum'] = len(stocks)
                    progress['value'] = 0
                    
                    # 添加進度標籤
                    progress_label = ttk.Label(self.progress_frame, text="計算中: 0%")
                    progress_label.pack()
                    
                    self.root.update()
                    
                    volatile_stocks = []
                    for i, stock in enumerate(stocks):
                        code, name, market_type = stock
                        result = self.calculate_volatility(code)
                        if isinstance(result, dict) and result.get('volatility', 0) >= 30:
                            market_display = '上市' if market_type == 'twse' else '上櫃'
                            volatile_stocks.append((
                                code, 
                                name, 
                                market_display, 
                                result  # 保存完整的結果字典
                            ))
                        
                        # 更新進度條和標籤
                        progress['value'] = i + 1
                        progress_percent = int((i + 1) / len(stocks) * 100)
                        progress_label.config(text=f"計算中: {progress_percent:.1f}%")
                        self.root.update()
                    
                    # 計算完成後移除進度條和標籤
                    progress.destroy()
                    progress_label.destroy()
                    
                    # 按波動率降序排序
                    volatile_stocks.sort(key=lambda x: x[3]['volatility'], reverse=True)
                    
                    # 添加到樹形視圖
                    for code, name, market_display, result in volatile_stocks:
                        display_text = (
                            f"{result['volatility']:.2f}% "
                            f"({result['start_date']}~{result['end_date']}) "
                            f"最高:{result['max_price']:.2f} "
                            f"最低:{result['min_price']:.2f} "
                            f"量:{result['volume']}"
                        )
                        self.stock_tree.insert('', tk.END, values=(code, name, market_display, display_text))
                        
                elif selected_function == "布林帶排序":
                    cursor.execute("SELECT code, name, market_type FROM stock_code ORDER BY code")
                    stocks = cursor.fetchall()
                    
                    # 創建進度條
                    progress = ttk.Progressbar(self.progress_frame, mode='determinate')
                    progress.pack(fill=tk.X, padx=5)
                    progress['maximum'] = len(stocks)
                    progress['value'] = 0
                    
                    # 添加進度標籤
                    progress_label = ttk.Label(self.progress_frame, text="計算中: 0%")
                    progress_label.pack()
                    
                    self.root.update()
                    
                    bollinger_stocks = []
                    for i, stock in enumerate(stocks):
                        code, name, market_type = stock
                        result = self.calculate_bollinger(code)
                        if result is not None:
                            market_display = '上市' if market_type == 'twse' else '上櫃'
                            bollinger_stocks.append((
                                code,
                                name,
                                market_display,
                                result
                            ))
                        
                        # 更新進度
                        progress['value'] = i + 1
                        progress_percent = int((i + 1) / len(stocks) * 100)
                        progress_label.config(text=f"計算中: {progress_percent:.1f}%")
                        self.root.update()
                    
                    progress.destroy()
                    progress_label.destroy()
                    
                    # 按布林帶位置排序
                    bollinger_stocks.sort(key=lambda x: x[3]['position'])
                    
                    # 添加到樹形視圖
                    for code, name, market_display, result in bollinger_stocks:
                        display_text = (
                            f"位置: {result['position']:.1f}% "
                            f"價格: {result['current_price']:.2f} "
                            f"上軌: {result['upper_band']:.2f} "
                            f"下軌: {result['lower_band']:.2f}"
                        )
                        self.stock_tree.insert('', tk.END, values=(code, name, market_display, display_text))
                        
                elif selected_function == "追蹤自選股":
                    # 獲取追蹤清單
                    cursor.execute("""
                        SELECT s.code, s.name, s.market_type, t.track_date 
                        FROM tracked_stocks t
                        JOIN stock_code s ON t.code = s.code 
                        ORDER BY t.track_date DESC
                    """)
                    
                    tracked_stocks = cursor.fetchall()
                    
                    # 添加到樹形視圖
                    for stock in tracked_stocks:
                        code, name, market_type, track_date = stock
                        market_display = '上市' if market_type == 'twse' else '上櫃'
                        self.stock_tree.insert('', tk.END, values=(
                            code,
                            name,
                            market_display,
                            f"追蹤日期: {track_date}"
                        ))
                        
                else:  # 所有股票
                    cursor.execute("""
                        SELECT code, name, market_type 
                        FROM stock_code 
                        WHERE market_type IS NOT NULL 
                        ORDER BY code
                    """)
                    stocks = cursor.fetchall()
                    for stock in stocks:
                        code, name, market_type = stock
                        market_display = '上市' if market_type == 'twse' else '上櫃'
                        self.stock_tree.insert('', tk.END, values=(code, name, market_display, ''))
                        
            except mysql.connector.Error as err:
                messagebox.showerror("錯誤", f"查詢數據庫時發生錯誤: {err}")
                
            finally:
                cursor.close()
                connection.close()
                
        except Exception as e:
            messagebox.showerror("錯誤", f"發生錯誤: {str(e)}")

    def calculate_bollinger_width(self, stock_code):
        """計算布林通道寬度"""
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            # 獲取最近一個易日的數據和前20天的收盤價
            table_name = f"stock_{stock_code}"
            cursor.execute(f"""
                SELECT date, close_price, volume 
                FROM {table_name} 
                ORDER BY date DESC 
                LIMIT 20
            """)
            
            rows = cursor.fetchall()
            if len(rows) < 20:  # 確保有足夠的數據
                return None
            
            # 檢查最近一天的成交量
            latest_date = rows[0][0]
            latest_volume = rows[0][2]
            if latest_volume is None or latest_volume < 100:  # 成交量過小的跳過
                return None
            
            # 取出收盤價
            prices = [row[1] for row in rows if row[1] is not None]
            if len(prices) < 20:
                return None
            
            # 計算布林通道
            prices = prices[::-1]  # 反轉順序以便計算
            ma20 = sum(prices) / len(prices)
            std = (sum((x - ma20) ** 2 for x in prices) / len(prices)) ** 0.5
            
            upper_band = ma20 + 2 * std
            lower_band = ma20 - 2 * std
            current_price = prices[-1]
            
            # 計算布林通道寬度百分比
            width_percent = ((upper_band - lower_band) / ma20) * 100
            
            return {
                'width_percent': width_percent,
                'current_price': current_price,
                'latest_date': latest_date,
                'volume': latest_volume
            }
            
        except mysql.connector.Error as err:
            print(f"計算布林通道寬度時發生錯誤 {stock_code}: {err}")
            return None
        except Exception as e:
            print(f"計算 {stock_code} 布林通道寬度時發生未預期的錯誤: {e}")
            return None
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def calculate_ma_trend(self, stock_code):
        """計算均線多頭排列"""
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            table_name = f"stock_{stock_code}"
            cursor.execute(f"""
                SELECT date, close_price, volume 
                FROM {table_name} 
                ORDER BY date DESC 
                LIMIT 30
            """)
            
            rows = cursor.fetchall()
            if len(rows) < 30:  # 確保有足夠的數據
                return None
            
            # 檢查最近一天的成交量
            latest_date = rows[0][0]
            latest_volume = rows[0][2]
            if latest_volume is None or latest_volume < 100:  # 成交量過小的跳過
                return None
            
            # 取出收盤價
            prices = [row[1] for row in rows if row[1] is not None]
            if len(prices) < 30:
                return None
            
            # 計算均線
            ma5 = sum(prices[:5]) / 5
            ma10 = sum(prices[:10]) / 10
            ma20 = sum(prices[:20]) / 20
            ma30 = sum(prices[:30]) / 30
            current_price = prices[0]
            
            # 判斷多頭排列 (MA5 > MA10 > MA20 > MA30)
            if ma5 > ma10 > ma20 > ma30:
                return {
                    'current_price': current_price,
                    'ma5': ma5,
                    'ma10': ma10,
                    'ma20': ma20,
                    'ma30': ma30,
                    'latest_date': latest_date,
                    'volume': latest_volume
                }
            
            return None
            
        except Exception as e:
            print(f"計算均線趨勢時發生錯誤 {stock_code}: {e}")
            return None
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def calculate_ma_breakthrough(self, stock_code):
        """計算均線突破"""
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            table_name = f"stock_{stock_code}"
            cursor.execute(f"""
                SELECT date, close_price, volume 
                FROM {table_name} 
                ORDER BY date DESC 
                LIMIT 21
            """)
            
            rows = cursor.fetchall()
            if len(rows) < 21:
                return None
            
            # 檢查最近一天的成交量
            latest_date = rows[0][0]
            latest_volume = rows[0][2]
            if latest_volume is None or latest_volume < 100:
                return None
            
            # 取出收盤價
            prices = [row[1] for row in rows if row[1] is not None]
            if len(prices) < 21:
                return None
            
            # 計算20日均線
            ma20 = sum(prices[1:21]) / 20  # 不包含最新價
            current_price = prices[0]
            prev_price = prices[1]
            
            # 判斷是否突破均線
            if prev_price < ma20 and current_price > ma20:
                return {
                    'current_price': current_price,
                    'ma20': ma20,
                    'breakthrough_percent': ((current_price - ma20) / ma20) * 100,
                    'latest_date': latest_date,
                    'volume': latest_volume
                }
            
            return None
            
        except Exception as e:
            print(f"計算均線突破時發生錯誤 {stock_code}: {e}")
            return None
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def calculate_volume_surge(self, stock_code):
        """計算量能放大"""
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            table_name = f"stock_{stock_code}"
            cursor.execute(f"""
                SELECT date, close_price, volume 
                FROM {table_name} 
                ORDER BY date DESC 
                LIMIT 6
            """)
            
            rows = cursor.fetchall()
            if len(rows) < 6:
                return None
                
            # 最近一天的成交量和價格
            latest_date = rows[0][0]
            latest_volume = rows[0][2]
            latest_price = rows[0][1]
            
            # 計算前5天平均成交量
            prev_volumes = [row[2] for row in rows[1:] if row[2] is not None]
            if not prev_volumes:
                return None
                
            avg_volume = sum(prev_volumes) / len(prev_volumes)
            
            # 計算量能放大倍數
            if avg_volume > 0 and latest_volume > 100:  # 確保有基本成交量
                volume_ratio = latest_volume / avg_volume
                
                if volume_ratio > 2:  # 成交量是前5天平均的2倍以上
                    return {
                        'volume_ratio': volume_ratio,
                        'current_volume': latest_volume,
                        'avg_volume': int(avg_volume),
                        'price': latest_price,
                        'date': latest_date
                    }
            
            return None
            
        except Exception as e:
            print(f"計算量能放大時發生錯誤 {stock_code}: {e}")
            return None
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def calculate_price_gap(self, stock_code):
        """計算跳空缺口"""
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            table_name = f"stock_{stock_code}"
            cursor.execute(f"""
                SELECT date, open_price, close_price, high_price, low_price, volume
                FROM {table_name} 
                ORDER BY date DESC 
                LIMIT 2
            """)
            
            rows = cursor.fetchall()
            if len(rows) < 2:
                return None
            
            # 今天的數據
            today_date = rows[0][0]
            today_open = rows[0][1]
            today_close = rows[0][2]
            today_volume = rows[0][5]
            
            # 昨天的數據
            yesterday_close = rows[1][2]
            yesterday_high = rows[1][3]
            yesterday_low = rows[1][4]
            
            # 確保基本成交量
            if today_volume < 100:
                return None
            
            # 計算向上跳空
            if today_open > yesterday_high:
                gap_percent = ((today_open - yesterday_high) / yesterday_high) * 100
                return {
                    'type': '向上跳空',
                    'gap_percent': gap_percent,
                    'current_price': today_close,
                    'volume': today_volume,
                    'date': today_date
                }
            
            # 計算向下跳空
            if today_open < yesterday_low:
                gap_percent = ((yesterday_low - today_open) / yesterday_low) * 100
                return {
                    'type': '向下跳空',
                    'gap_percent': gap_percent,
                    'current_price': today_close,
                    'volume': today_volume,
                    'date': today_date
                }
            
            return None
            
        except Exception as e:
            print(f"計算跳空缺口時發生錯誤 {stock_code}: {e}")
            return None
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def calculate_ma_cross(self, stock_code):
        """計算均線交叉"""
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            table_name = f"stock_{stock_code}"
            cursor.execute(f"""
                SELECT date, close_price, volume 
                FROM {table_name} 
                ORDER BY date DESC 
                LIMIT 22
            """)
            
            rows = cursor.fetchall()
            if len(rows) < 22:  # 確保有足夠的數據
                return None
                
            # 檢查最近一天的成交量
            latest_date = rows[0][0]
            latest_volume = rows[0][2]
            if latest_volume is None or latest_volume < 100:  # 成交量過小的過
                return None
                
            # 取出收盤價
            prices = [row[1] for row in rows if row[1] is not None]
            if len(prices) < 22:
                return None
                
            # 計算今天的均線
            ma5_today = sum(prices[:5]) / 5
            ma20_today = sum(prices[:20]) / 20
            
            # 計算昨天的均線
            ma5_yesterday = sum(prices[1:6]) / 5
            ma20_yesterday = sum(prices[1:21]) / 20
            
            # 判斷交叉
            if ma5_yesterday < ma20_yesterday and ma5_today > ma20_today:
                return {
                    'type': '黃金交叉',
                    'ma5': ma5_today,
                    'ma20': ma20_today,
                    'price': prices[0],
                    'volume': latest_volume,
                    'date': latest_date
                }
            elif ma5_yesterday > ma20_yesterday and ma5_today < ma20_today:
                return {
                    'type': '死亡交叉',
                    'ma5': ma5_today,
                    'ma20': ma20_today,
                    'price': prices[0],
                    'volume': latest_volume,
                    'date': latest_date
                }
            
            return None
            
        except Exception as e:
            print(f"計算均線交叉時發生錯誤 {stock_code}: {e}")
            return None
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def calculate_macd(self, stock_code):
        """計算 MACD 指標"""
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            table_name = f"stock_{stock_code}"
            cursor.execute(f"""
                SELECT date, close_price, volume 
                FROM {table_name} 
                ORDER BY date DESC 
                LIMIT 35
            """)
            
            rows = cursor.fetchall()
            if len(rows) < 35:  # 確保有足夠的數據
                return None
                
            # 檢查最近一天的成交量
            latest_date = rows[0][0]
            latest_volume = rows[0][2]
            if latest_volume is None or latest_volume < 100:  # 成交量過小的跳過
                return None
                
            # 取出收盤價
            prices = [row[1] for row in rows if row[1] is not None]
            if len(prices) < 35:
                return None
                
            # 計算 EMA
            def calculate_ema(data, period):
                multiplier = 2 / (period + 1)
                ema = data[0]  # 一個值作為始點
                for price in data[1:]:
                    ema = (price * multiplier) + (ema * (1 - multiplier))
                return ema
                
            # 計算 12 日 EMA
            ema12 = calculate_ema(prices[:12], 12)
            # 計算 26 日 EMA
            ema26 = calculate_ema(prices[:26], 26)
            
            # 計算 MACD 線
            macd_line = ema12 - ema26
            
            # 計算前一天的 MACD 值（用於判斷趨勢變化）
            prev_ema12 = calculate_ema(prices[1:13], 12)
            prev_ema26 = calculate_ema(prices[1:27], 26)
            prev_macd = prev_ema12 - prev_ema26
            
            # 計算 9 日 MACD 訊號線
            signal_line = calculate_ema([macd_line], 9)
            
            # 判斷 MACD 趨勢
            if macd_line > 0 and prev_macd < 0:
                trend = "黃金交叉"
            elif macd_line < 0 and prev_macd > 0:
                trend = "死亡交叉"
            else:
                return None  # 只返回交叉點
            
            return {
                'trend': trend,
                'macd': macd_line,
                'signal': signal_line,
                'price': prices[0],
                'volume': latest_volume,
                'date': latest_date
            }
            
        except Exception as e:
            print(f"計算 MACD 時發生錯誤 {stock_code}: {e}")
            return None
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def calculate_sideways(self, stock_code, days=20, threshold=0.05):
        """計算票是否處於橫向整理"""
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            table_name = f"stock_{stock_code}"
            cursor.execute(f"""
                SELECT date, close_price, volume 
                FROM {table_name} 
                ORDER BY date DESC 
                LIMIT {days}
            """)
            
            rows = cursor.fetchall()
            if len(rows) < days:  # 確保有足夠的數據
                return None
                
            # 檢查最近一天的成交量
            latest_date = rows[0][0]
            latest_volume = rows[0][2]
            if latest_volume is None or latest_volume < 100:  # 成交量過小的跳過
                return None
                
            # 取出收盤價
            prices = [row[1] for row in rows if row[1] is not None]
            if len(prices) < days:
                return None
                
            # 計算價格區間
            max_price = max(prices)
            min_price = min(prices)
            avg_price = sum(prices) / len(prices)
            
            # 計算波動範圍百分比
            range_percent = (max_price - min_price) / avg_price
            
            # 如果動範圍小於閾值，判定為橫向整理
            if range_percent <= threshold:
                return {
                    'range_percent': range_percent * 100,  # 轉換為百分比
                    'max_price': max_price,
                    'min_price': min_price,
                    'avg_price': avg_price,
                    'days': days,
                    'volume': latest_volume,
                    'date': latest_date
                }
            
            return None
            
        except Exception as e:
            print(f"計算橫向整理時發生錯誤 {stock_code}: {e}")
            return None
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def add_to_track(self):
        """加入追蹤清單"""
        code = self.code_entry.get().strip()
        if not code.isdigit() or len(code) != 4:
            messagebox.showerror("錯誤", "請輸入4位數字的股票代碼")
            return
            
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            # 檢查股票是否存在
            cursor.execute("SELECT name FROM stock_code WHERE code = %s", (code,))
            result = cursor.fetchone()
            
            if not result:
                messagebox.showerror("錯誤", "找不到此股票代碼")
                return
                
            # 使用 REPLACE INTO 來處理更新情況
            cursor.execute("""
                REPLACE INTO tracked_stocks (code, track_date) 
                VALUES (%s, CURDATE())
            """, (code,))
            
            connection.commit()
            messagebox.showinfo("成功", f"已將股票 {code} 加入追蹤清單")
            
        except mysql.connector.Error as err:
            messagebox.showerror("錯誤", f"資料庫錯誤: {err}")
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def get_tracked_stocks(self):
        """獲取追蹤清單"""
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            # 獲取追蹤清單
            cursor.execute("""
                SELECT t.code, s.name, t.track_date 
                FROM tracked_stocks t 
                JOIN stock_code s ON t.code = s.code 
                ORDER BY t.track_date DESC
            """)
            
            return cursor.fetchall()
            
        except mysql.connector.Error as err:
            messagebox.showerror("錯誤", f"資料庫錯誤: {err}")
            return []
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def calculate_daily_rebound(self, stock_code):
        """計算過去200天內每天的反彈幅度，並檢查5日均線趨勢"""
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            table_name = f"stock_{stock_code}"
            cursor.execute(f"""
                SELECT date, close_price, low_price, volume 
                FROM {table_name} 
                ORDER BY date DESC 
                LIMIT 200
            """)
            
            rows = cursor.fetchall()
            if len(rows) < 6:  # 確保有足夠的數據計算5日均線趨勢
                return None
                
            # 檢查最近一天的成交量
            latest_volume = rows[0][3]
            if latest_volume is None or latest_volume < 100:
                return None
                
            # 檢查每一天的反彈幅度和5日均線趨勢
            for i in range(len(rows)-5):  # 留出5天計算均線
                date = rows[i][0]
                close_price = rows[i][1]
                low_price = rows[i][2]
                volume = rows[i][3]
                
                # 檢查數據是否有效
                if None in [close_price, low_price, volume] or low_price <= 0:
                    continue
                
                # 計算反彈幅度
                rebound_percent = ((close_price - low_price) / low_price) * 100
                
                # 如果反彈幅度大於4%，再檢查5日均線趨勢
                if rebound_percent >= 4:
                    # 計算今天和昨天的5日均線
                    today_prices = [rows[j][1] for j in range(i, i+5)]
                    yesterday_prices = [rows[j][1] for j in range(i+1, i+6)]
                    
                    if None in today_prices or None in yesterday_prices:
                        continue
                        
                    today_ma5 = sum(today_prices) / 5
                    yesterday_ma5 = sum(yesterday_prices) / 5
                    
                    # 檢查5日均線是否下降
                    if today_ma5 < yesterday_ma5:
                        return {
                            'date': date,
                            'rebound_percent': rebound_percent,
                            'volume': volume,
                            'ma5': today_ma5  # 確保返回 ma5 值
                        }
            
            return None
            
        except Exception as e:
            print(f"計算反彈幅度時發生錯誤 {stock_code}: {e}")
            return None
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    def calculate_december_volatility(self, stock_code):
        """計算去年12月的價格波動"""
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            # 獲取去年12月的數據
            last_year = datetime.now().year - 1
            table_name = f"stock_{stock_code}"
            cursor.execute(f"""
                SELECT date, high_price, low_price, close_price, volume
                FROM {table_name}
                WHERE YEAR(date) = {last_year} AND MONTH(date) = 12
                ORDER BY date
            """)
            
            rows = cursor.fetchall()
            if len(rows) < 1:  # 確保有數據
                return None
                
            # 檢查成交量
            total_volume = sum(row[4] for row in rows if row[4] is not None)
            if total_volume < 1000:  # 成交量過小的跳過
                return None
                
            # 找出最高價和最低價
            high_prices = [row[1] for row in rows if row[1] is not None]
            low_prices = [row[2] for row in rows if row[2] is not None]
            
            if not high_prices or not low_prices:
                return None
                
            month_high = max(high_prices)
            month_low = min(low_prices)
            
            # 計算波動幅度
            if month_low > 0:  # 避免除以零
                volatility = ((month_high - month_low) / month_low) * 100
                
                # 如果波動超過20%
                if volatility >= 20:
                    # 取得月初和月底收盤價
                    first_close = rows[0][3]
                    last_close = rows[-1][3]
                    
                    return {
                        'volatility': volatility,
                        'month_high': month_high,
                        'month_low': month_low,
                        'first_close': first_close,
                        'last_close': last_close,
                        'trading_days': len(rows),
                        'volume': total_volume
                    }
            
            return None
            
        except Exception as e:
            print(f"計算12月波動率時發生錯誤 {stock_code}: {e}")
            return None
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = StockApp(root)
    root.mainloop()
