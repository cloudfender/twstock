import tkinter as tk
from tkinter import ttk, messagebox
import mysql.connector
from dotenv import load_dotenv
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import matplotlib
matplotlib.use('TkAgg')
plt.style.use('dark_background')  # 使用黑底主題
# 設置中文字體
plt.rcParams['font.sans-serif'] = ['DFKai-SB']  # 標楷體
plt.rcParams['axes.unicode_minus'] = False  # 解決負號顯示問題

# 載入環境變數
load_dotenv()

# 資料庫設定
DB_CONFIG = {
    'host': os.getenv('DB_HOST'),
    'user': os.getenv('DB_USER'),
    'password': os.getenv('DB_PASSWORD'),
    'database': os.getenv('DB_DATABASE')
}

class StockApp:
    def __init__(self, root):
        self.root = root
        self.root.title("股票查詢系統")
        
        # 建立主框架
        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 建立左側框架
        left_frame = ttk.Frame(self.main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        # 建立功能選擇下拉選單
        self.function_var = tk.StringVar()
        self.function_choices = [
            "所有股票",
            "上市股票",
            "上櫃股票",
            "波動30%以上股票",
            "布林帶排序",
            "追蹤自選股",
            "當日波動5%以上"
        ]
        self.function_var.set(self.function_choices[0])
        
        function_menu = ttk.OptionMenu(
            left_frame, 
            self.function_var, 
            self.function_choices[0], 
            *self.function_choices,
            command=self.on_function_change
        )
        function_menu.pack(fill=tk.X, pady=5)
        
        # 建立股票列表
        columns = ('代碼', '名稱', '市場', '資訊')
        self.stock_tree = ttk.Treeview(
            left_frame, 
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
            left_frame, 
            orient=tk.VERTICAL, 
            command=self.stock_tree.yview
        )
        self.stock_tree.configure(yscrollcommand=scrollbar.set)
        
        # 放置 Treeview 和捲動條
        self.stock_tree.pack(side=tk.LEFT, fill=tk.BOTH)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 建立右側框架
        right_frame = ttk.Frame(self.main_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)
        
        # 建立輸入框架
        input_frame = ttk.Frame(right_frame)
        input_frame.pack(fill=tk.X)
        
        # 股票代碼輸入
        ttk.Label(input_frame, text="請輸入股票代碼：").pack(side=tk.LEFT, padx=5)
        self.code_entry = ttk.Entry(input_frame, width=10)
        self.code_entry.pack(side=tk.LEFT, padx=5)
        
        # 天數輸入
        ttk.Label(input_frame, text="查詢天數：").pack(side=tk.LEFT, padx=5)
        self.days_entry = ttk.Entry(input_frame, width=10)
        self.days_entry.pack(side=tk.LEFT, padx=5)
        self.days_entry.insert(0, "300")
        
        # 在輸入框架中添加日期輸入和查詢按鈕
        ttk.Label(input_frame, text="日期：").pack(side=tk.LEFT, padx=5)
        self.date_entry = ttk.Entry(input_frame, width=10)
        self.date_entry.pack(side=tk.LEFT, padx=5)
        # 設置預設值為今天
        today = datetime.now().strftime('%Y-%m-%d')
        self.date_entry.insert(0, today)
        
        # 添加波動率查詢按鈕
        self.volatility_button = ttk.Button(
            input_frame, 
            text="查詢波動率", 
            command=self.search_volatility
        )
        self.volatility_button.pack(side=tk.LEFT, padx=5)
        
        # 查詢按鈕
        ttk.Button(input_frame, text="查詢", command=self.search_stock).pack(side=tk.LEFT, padx=5)
        ttk.Button(input_frame, text="加入追蹤", command=self.add_to_track).pack(side=tk.LEFT, padx=5)
        
        # 修改圖表創建部分
        self.fig, (self.ax1, self.ax2) = plt.subplots(2, 1,
            figsize=(12, 8),  # 加大圖表尺寸
            height_ratios=[3, 1],
            gridspec_kw={'hspace': 0.1}  # 減少子圖之間的間距
        )
        
        # 調整圖表邊距
        self.fig.subplots_adjust(
            left=0.05,    # 左邊距
            right=0.95,   # 右邊距
            top=0.95,     # 上邊距
            bottom=0.05   # 下邊距
        )
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=right_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 數據顯示標籤
        self.info_label = ttk.Label(right_frame, text="", justify=tk.LEFT)
        self.info_label.pack(fill=tk.X, pady=5)
        
        # 載入股票列表
        self.load_stock_list()
        
        # 綁定事件
        self.code_entry.bind('<Return>', lambda e: self.search_stock())
        self.days_entry.bind('<Return>', lambda e: self.search_stock())
        self.stock_tree.bind('<ButtonRelease-1>', self.on_stock_select)
        self.stock_tree.bind('<Up>', self.on_key_up)
        self.stock_tree.bind('<Down>', self.on_key_down)
        self.date_entry.bind('<Return>', lambda e: self.search_volatility())
        
        # 確保 stock_tree 可以接收鍵盤事件
        self.stock_tree.focus_set()
    
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
    
    def on_function_change(self, event):
        """處理功能選擇變更"""
        selected = self.function_var.get()
        
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            # 清空現有項目
            for item in self.stock_tree.get_children():
                self.stock_tree.delete(item)
            
            if selected == "當日波動5%以上":
                # 獲取指定日期
                target_date = self.date_entry.get().strip()
                if not target_date:
                    messagebox.showerror("錯誤", "請輸入日期")
                    return
                
                try:
                    # 驗證日期格式
                    datetime.strptime(target_date, '%Y-%m-%d')
                except ValueError:
                    messagebox.showerror("錯誤", "日期格式錯誤，請使用 YYYY-MM-DD 格式")
                    return
                
                # 2. 獲取所有股票代碼
                cursor.execute("SELECT code, name, market_type FROM stock_code WHERE market_type IS NOT NULL")
                stocks = cursor.fetchall()
                
                # 3. 對每支股票查詢指定日期數據
                volatile_stocks = []
                for stock in stocks:
                    code, name, market = stock
                    table_name = f"stock_{code}"
                    
                    try:
                        # 檢查表是否存在並查詢數據
                        cursor.execute(f"""
                            SELECT high_price, low_price, volume, close_price  # 添加 close_price
                            FROM {table_name} 
                            WHERE date = %s
                        """, (target_date,))
                        
                        result = cursor.fetchone()
                        if result:
                            high_price, low_price, volume, close_price = result  # 獲取收盤價
                            # 只有當成交量大於 100000 時才計算波動率
                            if volume >= 100000:
                                # 計算波動率
                                volatility = ((high_price - low_price) / low_price * 100)
                                
                                if volatility >= 5:
                                    # 計算成交金額 (成交量 * 收盤價)
                                    trading_value = volume * close_price
                                    
                                    # 計算收盤價在最高價和最低價之間的位置
                                    price_range = high_price - low_price
                                    close_position = (close_price - low_price) / price_range
                                    
                                    volatile_stocks.append({
                                        'code': code,
                                        'name': name,
                                        'market': market,
                                        'volatility': round(volatility, 2),
                                        'volume': volume,
                                        'close_position': close_position,
                                        'trading_value': trading_value  # 添加成交金額
                                    })
                    except mysql.connector.Error:
                        continue  # 跳過不存在的表
                
                # 4. 依成交金額排序（而不是波動率）
                volatile_stocks.sort(key=lambda x: x['trading_value'], reverse=True)
                
                # 輸出 CSV 檔案
                try:
                    # 準備 CSV 檔案名稱
                    csv_filename = f"當日波動5%以上_{target_date}.csv"
                    
                    # 寫入 CSV
                    with open(csv_filename, 'w', encoding='utf-8') as f:
                        # 直接寫入股票代碼，不寫入標題行
                        for stock in volatile_stocks:
                            f.write(f"{stock['code']}.TW\n")
                    
                    # 顯示成功訊息
                    self.info_label.config(text=f"日期: {target_date}, 共找到 {len(volatile_stocks)} 檔股票，已輸出到 {csv_filename}")
                except Exception as e:
                    messagebox.showerror("錯誤", f"輸出CSV時發生錯誤: {str(e)}")
                
                # 5. 顯示結果
                for stock in volatile_stocks:
                    market_display = '上市' if stock['market'] == 'twse' else '上櫃'
                    
                    # 根據收盤位置決定顯示顏色
                    if stock['close_position'] <= 0.3:  # 靠近最低點
                        color = 'green'
                    elif stock['close_position'] >= 0.7:  # 靠近最高點
                        color = 'red'
                    else:  # 在中間
                        color = '#808080'
                    
                    # 格式化成交金額（以億為單位）
                    trading_value_billions = stock['trading_value'] / 100000000
                    
                    item = self.stock_tree.insert('', tk.END, values=(
                        stock['code'],
                        stock['name'],
                        market_display,
                        f"波動: {stock['volatility']}%, 量: {stock['volume']:,}, 金額: {trading_value_billions:.2f}億"
                    ))
                    
                    # 設置項目的文字顏色
                    self.stock_tree.tag_configure(color, foreground=color)
                    self.stock_tree.item(item, tags=(color,))
                    
                # 顯示查詢日期和結果數量
                self.info_label.config(text=f"日期: {target_date}, 共找到 {len(volatile_stocks)} 檔股票")
                
            elif selected == "追蹤自選股":
                # 獲取追蹤股票列表，按追蹤日期降序排序
                cursor.execute("""
                    SELECT s.code, s.name, s.market_type, t.track_date 
                    FROM tracked_stocks t
                    JOIN stock_code s ON t.code = s.code 
                    ORDER BY t.track_date DESC
                """)
                stocks = cursor.fetchall()
                
                # 插入數據
                for stock in stocks:
                    code, name, market, track_date = stock
                    market_display = '上市' if market == 'twse' else '上櫃'
                    self.stock_tree.insert('', tk.END, values=(
                        code,
                        name,
                        market_display,
                        track_date.strftime('%Y-%m-%d')  # 只顯示日期
                    ))
            elif selected == "上市股票":
                # 獲取上市股票
                cursor.execute("""
                    SELECT code, name, market_type 
                    FROM stock_code 
                    WHERE market_type = 'twse'
                    ORDER BY code
                """)
            elif selected == "上櫃股票":
                # 獲取上櫃股票
                cursor.execute("""
                    SELECT code, name, market_type 
                    FROM stock_code 
                    WHERE market_type = 'tpex'
                    ORDER BY code
                """)
            else:  # "所有股票" 或其他選項
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
            messagebox.showerror("錯誤", f"數據庫錯誤: {err}")
        finally:
            if 'connection' in locals() and connection.is_connected():
                cursor.close()
                connection.close()

    def search_stock(self):
        """搜尋股票並顯示K線圖"""
        code = self.code_entry.get().strip()
        if not code:
            return
        
        try:
            days = int(self.days_entry.get().strip())
            
            # 從資料庫獲取數據
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            # 檢查股票代碼格式並補0
            code = code.zfill(4)  # 補0到4位數
            
            # 檢查股票表是否存在
            table_name = f"stock_{code}"
            
            # 先檢查表是否存在
            cursor.execute("""
                SELECT COUNT(*) 
                FROM information_schema.tables 
                WHERE table_schema = %s 
                AND table_name = %s
            """, (DB_CONFIG['database'], table_name))
            
            if cursor.fetchone()[0] == 0:
                messagebox.showerror("錯誤", f"找不到股票 {code} 的資料表")
                return
            
            cursor.execute(f"""
                SELECT date, open_price, high_price, low_price, close_price, volume 
                FROM {table_name} 
                ORDER BY date DESC 
                LIMIT {days}
            """)
            
            data = cursor.fetchall()
            if not data:
                messagebox.showerror("錯誤", "無法獲取股票數據")
                return
            
            # 轉換數據
            dates = []
            opens = []
            highs = []
            lows = []
            closes = []
            volumes = []
            
            for row in reversed(data):  # 反轉數據以按時間順序排列
                dates.append(row[0])  # 保存完整的日期對象
                opens.append(float(row[1]))
                highs.append(float(row[2]))
                lows.append(float(row[3]))
                closes.append(float(row[4]))
                volumes.append(int(row[5]))
            
            # 清除舊圖並設置背景
            for ax in [self.ax1, self.ax2]:
                ax.clear()
                ax.set_facecolor('black')
            self.fig.set_facecolor('black')
            
            # 創建日期索引
            date_index = range(len(dates))
            
            # 繪製K線圖
            self.draw_candlestick(self.ax1, date_index, opens, highs, lows, closes)
            
            # 計算布林通道
            closes_series = pd.Series(closes)
            sma = closes_series.rolling(window=20).mean()
            std = closes_series.rolling(window=20).std()
            upper_band = sma + (std * 2)
            lower_band = sma - (std * 2)
            
            # 繪製布林通道
            self.ax1.plot(date_index, sma, label='中軌', color='yellow', linewidth=1)
            self.ax1.plot(date_index, upper_band, label='上軌', color='magenta', linewidth=1)
            self.ax1.plot(date_index, lower_band, label='下軌', color='cyan', linewidth=1)
            
            # 設置圖表格式
            self.ax1.set_title(f"{code} 日K線圖", color='white', pad=10)
            self.ax1.grid(True, color='gray', alpha=0.2)
            self.ax1.legend(facecolor='black', edgecolor='white', loc='upper right')
            
            # 繪製成交量
            volume_colors = ['red' if closes[i] >= opens[i] else 'green' for i in range(len(dates))]
            self.ax2.bar(date_index, volumes, color=volume_colors)
            self.ax2.set_title("成交量", color='white', pad=5)
            self.ax2.grid(True, color='gray', alpha=0.2)
            
            # 設置日期刻度
            # 找出每個月份的第一天的索引
            month_first_days = []
            current_month = None
            for i, date in enumerate(dates):
                if current_month != date.strftime('%Y-%m'):
                    current_month = date.strftime('%Y-%m')
                    month_first_days.append(i)
            
            # 設置 x 軸刻度
            self.ax2.set_xticks(month_first_days)
            self.ax2.set_xticklabels([dates[i].strftime('%Y-%m') for i in month_first_days], 
                                    rotation=45, ha='right')
            
            # K線圖不顯示 x 軸標籤
            self.ax1.set_xticks(month_first_days)
            self.ax1.set_xticklabels([])
            
            # 設置軸線顏色
            for ax in [self.ax1, self.ax2]:
                ax.tick_params(colors='white')
                ax.spines['bottom'].set_color('white')
                ax.spines['top'].set_color('white')
                ax.spines['left'].set_color('white')
                ax.spines['right'].set_color('white')
                # 添加垂直網格線在月份分隔處
                ax.grid(True, which='major', axis='x', color='gray', alpha=0.2)
            
            # 在繪製完K線圖後，檢查是否需要添加垂直線
            if self.function_var.get() == "當日波動5%以上":
                target_date = self.date_entry.get().strip()
                try:
                    target_date = datetime.strptime(target_date, '%Y-%m-%d')  # 不需要 .date()
                    # 找到目標日期的索引
                    for i, date in enumerate(dates):
                        if date.strftime('%Y-%m-%d') == target_date.strftime('%Y-%m-%d'):  # 比較日期字符串
                            break  # 添加垂直線后，跳出循环
                except ValueError:
                    pass  # 如果日期格式錯誤，就不添加垂直線
            
            # 更新圖表
            self.canvas.draw()
            
        except Exception as e:
            messagebox.showerror("錯誤", f"查詢時發生錯誤: {str(e)}")
        finally:
            if 'connection' in locals() and connection.is_connected():
                cursor.close()
                connection.close()

    def draw_candlestick(self, ax, dates, opens, highs, lows, closes):
        """繪製K線"""
        width = 0.6
        for i in range(len(dates)):
            # 判斷漲跌
            if closes[i] >= opens[i]:
                color = 'red'
                body_bottom = opens[i]
                body_height = closes[i] - opens[i]
            else:
                color = 'green'
                body_bottom = closes[i]
                body_height = opens[i] - closes[i]
            
            # 繪製K線實體
            ax.bar(dates[i], body_height, width, bottom=body_bottom, color=color)
            # 繪製上下影線
            ax.plot([dates[i], dates[i]], [lows[i], highs[i]], color=color, linewidth=1)

    def add_to_track(self):
        """添加股票到追蹤清單"""
        code = self.code_entry.get().strip()
        if not code:
            messagebox.showerror("錯誤", "請輸入股票代碼")
            return
        
        try:
            connection = mysql.connector.connect(**DB_CONFIG)
            cursor = connection.cursor()
            
            # 檢查股票是否存在
            cursor.execute("SELECT name FROM stock_code WHERE code = %s", (code,))
            stock = cursor.fetchone()
            if not stock:
                messagebox.showerror("錯誤", "找不到此股票")
                return
            
            # 添加到追蹤清單
            cursor.execute("""
                INSERT INTO tracked_stocks (code, track_date) 
                VALUES (%s, CURRENT_DATE())
                ON DUPLICATE KEY UPDATE track_date = CURRENT_DATE()
            """)
            connection.commit()
            
            messagebox.showinfo("成功", f"已將股票 {code} 加入追蹤清單")
            
            # 如果當前是在追蹤自選股頁面，則更新顯示
            if self.function_var.get() == "追蹤自選股":
                self.on_function_change(None)
            
        except mysql.connector.Error as err:
            messagebox.showerror("錯誤", f"數據庫錯誤: {err}")
        finally:
            if 'connection' in locals() and connection.is_connected():
                cursor.close()
                connection.close()

    def on_stock_select(self, event):
        """處理股票選擇事件"""
        selection = self.stock_tree.selection()
        if selection:
            item = self.stock_tree.item(selection[0])
            code = str(item['values'][0])
            self.code_entry.delete(0, tk.END)
            self.code_entry.insert(0, code)
            self.search_stock()

    def on_key_up(self, event):
        """處理向上鍵事件"""
        selection = self.stock_tree.selection()
        if not selection:
            return
        
        # 獲取當前選中項的索引
        current_index = self.stock_tree.index(selection[0])
        if current_index > 0:
            # 選擇上一個項目
            prev_item = self.stock_tree.get_children()[current_index - 1]
            self.stock_tree.selection_set(prev_item)
            self.stock_tree.see(prev_item)  # 確保項目可見
            self.on_stock_select(None)  # 觸發選擇事件

    def on_key_down(self, event):
        """處理向下鍵事件"""
        selection = self.stock_tree.selection()
        if not selection:
            # 如果沒有選中項，選擇第一個
            first_item = self.stock_tree.get_children()[0]
            self.stock_tree.selection_set(first_item)
            self.stock_tree.see(first_item)
            self.on_stock_select(None)
            return
        
        # 獲取當前選中項的索引
        current_index = self.stock_tree.index(selection[0])
        items = self.stock_tree.get_children()
        if current_index < len(items) - 1:
            # 選擇下一個項目
            next_item = items[current_index + 1]
            self.stock_tree.selection_set(next_item)
            self.stock_tree.see(next_item)  # 確保項目可見
            self.on_stock_select(None)  # 觸發選擇事件

    def search_volatility(self):
        """查詢指定日期的波動率"""
        # 設置功能選單為"當日波動5%以上"
        self.function_var.set("當日波動5%以上")
        # 觸發功能變更事件
        self.on_function_change(None)

def main():
    root = tk.Tk()
    app = StockApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()