import mysql.connector
from FinMind.data import DataLoader

def update_stock_info():
    dl = DataLoader()
    # 獲取所有台股股票資訊
    stock_info = dl.taiwan_stock_info()
    
    # 連接到資料庫
    connection = mysql.connector.connect(
        host='192.168.1.145',
        user='root',
        password='fender1106',
        database='stock_code'
    )
    cursor = connection.cursor()
    
    try:
        # 更新資料庫中的股票名稱
        for _, row in stock_info.iterrows():
            try:
                code = row['stock_id']
                name = row['stock_name']
                # 檢查是否有 market_type 欄位，如果沒有，使用 type 欄位
                market = row.get('market_type', row.get('type', 'Unknown'))
                
                print(f"處理股票: {code} {name} ({market})")
                
                # 更新或插入資料
                cursor.execute("""
                    INSERT INTO stock_code (code, name, market_type) 
                    VALUES (%s, %s, %s)
                    ON DUPLICATE KEY UPDATE name = VALUES(name), market_type = VALUES(market_type)
                """, (code, name, market))
                
            except Exception as e:
                print(f"處理股票 {code} 時發生錯誤: {e}")
                continue
        
        connection.commit()
        print("股票資訊更新完成")
        
    except Exception as e:
        print(f"更新失敗: {e}")
        # 印出更多錯誤信息
        print("資料欄位:", stock_info.columns.tolist())
        print("資料範例:", stock_info.iloc[0].to_dict())
    finally:
        cursor.close()
        connection.close()

if __name__ == "__main__":
    update_stock_info() 