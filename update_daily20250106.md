我要重新再開發一次股票價格資訊回補的功能(update_daily20250106.py)
1. 資料庫主機: 192.168.1.145 mysql 帳號: root 密碼: fender1106 資料庫: stock4(建立一個.env的檔案，方便之後修改)
2. 目前資料庫裡面有2個資料表:
     (1)`stack_Ncode` : 記錄股票代碼 0000~9999， 對應"沒有股票"的代碼
        CREATE TABLE `stack_Ncode` (
            `id` int(11) NOT NULL comment '流水號',
            `code` varchar(10) NOT NULL comment '股票代碼',
            PRIMARY KEY (`id`)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;
     (2)`stack_code` : 記錄股票代碼 0000~9999， 對應到"有"股票的代碼
        CREATE TABLE `stock_code` (
            `id` int(11) NOT NULL comment '流水號',
            `code` varchar(8) NOT NULL comment '股票代碼',
            `name` varchar(50) DEFAULT NULL comment '股票名稱',
            `market_type` varchar(10) DEFAULT NULL comment '股票市場類型-上市twse, 上櫃-tpex'
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

3. 個股每日的股價資訊回補統一使用下列table 格式:
    CREATE TABLE `stock_xxxx`  xxxx 為股票代碼
        `date` date NOT NULL comment '日期',
        `open` decimal(10,2) DEFAULT NULL comment '開盤價',
        `high` decimal(10,2) DEFAULT NULL comment '最高價',
        `low` decimal(10,2) DEFAULT NULL comment '最低價',
        `close` decimal(10,2) DEFAULT NULL comment '收盤價',
        `volume` bigint(20) DEFAULT NULL comment '成交量',
        PRIMARY KEY (`date`)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;






備註1: temp_table_create.py 是建立所有個股每日交易資料表的程式
