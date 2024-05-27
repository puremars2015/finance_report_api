## import excel 資料


def convert(excel_file_path, excel_file_output):

    import os
    import pandas as pd
    import sqlite3

    # Excel文件路徑
    # excel_file_path = 'source.xlsx'
    # SQLite數據庫文件路徑
    sqlite_db_path = 'data0520.db'
    # excel輸出名稱
    # excel_file_output = 'converted.xlsx'

    # 使用pandas的ExcelFile對象讀取Excel文件
    excel_file = pd.ExcelFile(excel_file_path)

    # 創建或打開SQLite數據庫
    conn = sqlite3.connect(sqlite_db_path)

    # 遍歷所有分頁
    for sheet_name in excel_file.sheet_names:
        # 讀取當前分頁到DataFrame
        df = excel_file.parse(sheet_name)
        
        # 將DataFrame導入到以分頁名稱命名的SQLite表中
        # 如果該表不存在，它將會被創建
        df.to_sql(sheet_name, conn, if_exists='replace', index=False)

    # 關閉數據庫連接
    conn.close()

    print('Excel數據已成功導入到SQLite數據庫。')





    ## 新增表格

    conn = sqlite3.connect(sqlite_db_path)
    c = conn.cursor()


    # 創建一個名為 'users' 的表格
    # 如果表格已存在，則先刪除該表格
    c.execute('''DROP TABLE IF EXISTS 各單位科目與金額V2''')
    c.execute('''CREATE TABLE 各單位科目與金額V2 (
                五大部門分類 TEXT,
                成本歸屬部門代碼 TEXT,
                成本歸屬部門 TEXT,
                計薪人數 INTEGER,
                在職人數 NUMERIC,
                項目 TEXT,
                借方科目代號 TEXT,
                貸方科目代號 TEXT,
                NTD INTEGER
            )''')

    # 提交（保存）更改
    conn.commit()

    # 關閉連線
    conn.close()







    ## convert table

    conn = sqlite3.connect(sqlite_db_path)
    c = conn.cursor()

    c.execute('''INSERT INTO 各單位科目與金額V2
    SELECT * FROM (

    SELECT  
    d1.`五大部門分類`,
    d1.`成本歸屬部門代碼`,
    d1.`成本歸屬部門`,
    d1.`計薪人數`,
    d1.`在職人數`,
    '本薪' AS 項目,
    CAST(d1.`本薪會科` AS INTEGER) AS 借方科目代號,
    CAST(d1.`本薪會科貸方` AS INTEGER) AS 貸方科目代號,
    d1.`本薪` AS NTD
    FROM 各單位科目與金額 d1
    WHERE d1.`五大部門分類` IS NOT NULL

    UNION ALL
    SELECT  
    d2.`五大部門分類`,
    d2.`成本歸屬部門代碼`,
    d2.`成本歸屬部門`,
    d2.`計薪人數`,
    d2.`在職人數`,
    '伙食津貼' AS 項目,
    CAST(d2.`伙食津貼會科` AS INTEGER) AS 借方科目代號,
    CAST(d2.`伙食津貼會科貸方` AS INTEGER) AS 貸方科目代號,
    d2.`伙食津貼` AS NTD
    FROM 各單位科目與金額 d2
    WHERE d2.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d3.`五大部門分類`,
    d3.`成本歸屬部門代碼`,
    d3.`成本歸屬部門`,
    d3.`計薪人數`,
    d3.`在職人數`,
    '全勤獎金' AS 項目,
    CAST(d3.`全勤獎金會科` AS INTEGER) AS 借方科目代號,
    CAST(d3.`全勤獎金會科貸方` AS INTEGER) AS 貸方科目代號,
    d3.`全勤獎金` AS NTD
    FROM 各單位科目與金額 d3
    WHERE d3.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d4.`五大部門分類`,
    d4.`成本歸屬部門代碼`,
    d4.`成本歸屬部門`,
    d4.`計薪人數`,
    d4.`在職人數`,
    '其他津貼' AS 項目,
    CAST(d4.`其他津貼會科` AS INTEGER) AS 借方科目代號,
    CAST(d4.`其他津貼會科貸方` AS INTEGER) AS 貸方科目代號,
    d4.`其他津貼` AS NTD
    FROM 各單位科目與金額 d4
    WHERE d4.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d5.`五大部門分類`,
    d5.`成本歸屬部門代碼`,
    d5.`成本歸屬部門`,
    d5.`計薪人數`,
    d5.`在職人數`,
    '主管加給' AS 項目,
    CAST(d5.`主管加給會科` AS INTEGER) AS 借方科目代號,
    CAST(d5.`主管加給會科貸方` AS INTEGER) AS 貸方科目代號,
    d5.`主管加給` AS NTD
    FROM 各單位科目與金額 d5
    WHERE d5.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d6.`五大部門分類`,
    d6.`成本歸屬部門代碼`,
    d6.`成本歸屬部門`,
    d6.`計薪人數`,
    d6.`在職人數`,
    '應稅加項' AS 項目,
    CAST(d6.`應稅加項會科` AS INTEGER) AS 借方科目代號,
    CAST(d6.`應稅加項會科貸方` AS INTEGER) AS 貸方科目代號,
    d6.`應稅加項` AS NTD
    FROM 各單位科目與金額 d6
    WHERE d6.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d7.`五大部門分類`,
    d7.`成本歸屬部門代碼`,
    d7.`成本歸屬部門`,
    d7.`計薪人數`,
    d7.`在職人數`,
    '法院扣款' AS 項目,
    CAST(d7.`法院扣款會科` AS INTEGER) AS 借方科目代號,
    CAST(d7.`法院扣款會科貸方` AS INTEGER) AS 貸方科目代號,
    d7.`法院扣款` AS NTD
    FROM 各單位科目與金額 d7
    WHERE d7.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d8.`五大部門分類`,
    d8.`成本歸屬部門代碼`,
    d8.`成本歸屬部門`,
    d8.`計薪人數`,
    d8.`在職人數`,
    '應稅扣項' AS 項目,
    CAST(d8.`應稅扣項會科` AS INTEGER) AS 借方科目代號,
    CAST(d8.`應稅扣項會科貸方` AS INTEGER) AS 貸方科目代號,
    d8.`應稅扣項` AS NTD
    FROM 各單位科目與金額 d8
    WHERE d8.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d9.`五大部門分類`,
    d9.`成本歸屬部門代碼`,
    d9.`成本歸屬部門`,
    d9.`計薪人數`,
    d9.`在職人數`,
    '免稅扣項' AS 項目,
    CAST(d9.`代扣款會科` AS INTEGER)  AS 借方科目代號,
    CAST(d9.`代扣款會科貸方` AS INTEGER)  AS 貸方科目代號,
    d9.`免稅扣項` AS NTD
    FROM 各單位科目與金額 d9
    WHERE d9.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d10.`五大部門分類`,
    d10.`成本歸屬部門代碼`,
    d10.`成本歸屬部門`,
    d10.`計薪人數`,
    d10.`在職人數`,
    '輪班加班費' AS 項目,
    CAST(d10.`輪班加班費會科` AS INTEGER) AS 借方科目代號,
    CAST(d10.`輪班加班費會科貸方` AS INTEGER) AS 貸方科目代號,
    d10.`輪班加班費` AS NTD
    FROM 各單位科目與金額 d10
    WHERE d10.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d11.`五大部門分類`,
    d11.`成本歸屬部門代碼`,
    d11.`成本歸屬部門`,
    d11.`計薪人數`,
    d11.`在職人數`,
    '班別津貼' AS 項目,
    CAST(d11.`班別津貼會科` AS INTEGER) AS 借方科目代號,
    CAST(d11.`班別津貼會科貸方` AS INTEGER) AS 貸方科目代號,
    d11.`班別津貼` AS NTD
    FROM 各單位科目與金額 d11
    WHERE d11.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d12.`五大部門分類`,
    d12.`成本歸屬部門代碼`,
    d12.`成本歸屬部門`,
    d12.`計薪人數`,
    d12.`在職人數`,
    '雇主應付勞保費' AS 項目,
    CAST(d12.`雇主應付勞保費會科` AS INTEGER) AS 借方科目代號,
    CAST(d12.`雇主應付勞保費會科貸方` AS INTEGER) AS 貸方科目代號,
    d12.`雇主應付勞保費` AS NTD
    FROM 各單位科目與金額 d12
    WHERE d12.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d13.`五大部門分類`,
    d13.`成本歸屬部門代碼`,
    d13.`成本歸屬部門`,
    d13.`計薪人數`,
    d13.`在職人數`,
    '雇主應付健保費' AS 項目,
    CAST(d13.`雇主應付健保費會科` AS INTEGER) AS 借方科目代號,
    CAST(d13.`雇主應付健保費會科貸方` AS INTEGER) AS 貸方科目代號,
    d13.`雇主應付健保費` AS NTD
    FROM 各單位科目與金額 d13
    WHERE d13.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d14.`五大部門分類`,
    d14.`成本歸屬部門代碼`,
    d14.`成本歸屬部門`,
    d14.`計薪人數`,
    d14.`在職人數`,
    '雇主勞退新制提撥' AS 項目,
    CAST(d14.`雇主勞退新制提撥會科` AS INTEGER) AS 借方科目代號,
    CAST(d14.`雇主勞退新制提撥會科貸方` AS INTEGER) AS 貸方科目代號,
    d14.`雇主勞退新制提撥` AS NTD
    FROM 各單位科目與金額 d14
    WHERE d14.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d15.`五大部門分類`,
    d15.`成本歸屬部門代碼`,
    d15.`成本歸屬部門`,
    d15.`計薪人數`,
    d15.`在職人數`,
    '伙食扣款' AS 項目,
    CAST(d15.`伙食扣款會科` AS INTEGER) AS 借方科目代號,
    CAST(d15.`伙食扣款會科貸方` AS INTEGER) AS 貸方科目代號,
    d15.`伙食扣款` AS NTD
    FROM 各單位科目與金額 d15
    WHERE d15.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d16.`五大部門分類`,
    d16.`成本歸屬部門代碼`,
    d16.`成本歸屬部門`,
    d16.`計薪人數`,
    d16.`在職人數`,
    '免稅加班費' AS 項目,
    CAST(d16.`免稅加班費會科` AS INTEGER) AS 借方科目代號,
    CAST(d16.`免稅加班費會科貸方` AS INTEGER) AS 貸方科目代號,
    d16.`免稅加班費` AS NTD
    FROM 各單位科目與金額 d16
    WHERE d16.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d17.`五大部門分類`,
    d17.`成本歸屬部門代碼`,
    d17.`成本歸屬部門`,
    d17.`計薪人數`,
    d17.`在職人數`,
    '應稅加班費' AS 項目,
    CAST(d17.`應稅加班費會科` AS INTEGER) AS 借方科目代號,
    CAST(d17.`應稅加班費會科貸方` AS INTEGER) AS 貸方科目代號,
    d17.`應稅加班費` AS NTD
    FROM 各單位科目與金額 d17
    WHERE d17.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d18.`五大部門分類`,
    d18.`成本歸屬部門代碼`,
    d18.`成本歸屬部門`,
    d18.`計薪人數`,
    d18.`在職人數`,
    '請假扣款小計' AS 項目,
    CAST(d18.`請假扣款小計會科` AS INTEGER) AS 借方科目代號,
    CAST(d18.`請假扣款小計會科貸方` AS INTEGER) AS 貸方科目代號,
    d18.`請假扣款小計` AS NTD
    FROM 各單位科目與金額 d18
    WHERE d18.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d19.`五大部門分類`,
    d19.`成本歸屬部門代碼`,
    d19.`成本歸屬部門`,
    d19.`計薪人數`,
    d19.`在職人數`,
    '未休年假折現' AS 項目,
    CAST(d19.`未休年假折現會科` AS INTEGER) AS 借方科目代號,
    CAST(d19.`未休年假折現會科貸方` AS INTEGER) AS 貸方科目代號,
    d19.`未休年假折現` AS NTD
    FROM 各單位科目與金額 d19
    WHERE d19.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d20.`五大部門分類`,
    d20.`成本歸屬部門代碼`,
    d20.`成本歸屬部門`,
    d20.`計薪人數`,
    d20.`在職人數`,
    '未休補休假折現' AS 項目,
    CAST(d20.`未休補休假折現會科` AS INTEGER) AS 借方科目代號,
    CAST(d20.`未休補休假折現會科貸方` AS INTEGER) AS 貸方科目代號,
    d20.`未休補休假折現` AS NTD
    FROM 各單位科目與金額 d20
    WHERE d20.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d21.`五大部門分類`,
    d21.`成本歸屬部門代碼`,
    d21.`成本歸屬部門`,
    d21.`計薪人數`,
    d21.`在職人數`,
    '職工福利金' AS 項目,
    CAST(d21.`職工福利金會科` AS INTEGER) AS 借方科目代號,
    CAST(d21.`職工福利金會科貸方` AS INTEGER) AS 貸方科目代號,
    d21.`職工福利金` AS NTD
    FROM 各單位科目與金額 d21
    WHERE d21.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d22.`五大部門分類`,
    d22.`成本歸屬部門代碼`,
    d22.`成本歸屬部門`,
    d22.`計薪人數`,
    d22.`在職人數`,
    '勞保費' AS 項目,
    CAST(d22.`勞保費會科` AS INTEGER) AS 借方科目代號,
    CAST(d22.`勞保費會科貸方` AS INTEGER) AS 貸方科目代號,
    d22.`勞保費` AS NTD
    FROM 各單位科目與金額 d22
    WHERE d22.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d23.`五大部門分類`,
    d23.`成本歸屬部門代碼`,
    d23.`成本歸屬部門`,
    d23.`計薪人數`,
    d23.`在職人數`,
    '健保費' AS 項目,
    CAST(d23.`健保費會科` AS INTEGER) AS 借方科目代號,
    CAST(d23.`健保費會科貸方` AS INTEGER) AS 貸方科目代號,
    d23.`健保費` AS NTD
    FROM 各單位科目與金額 d23
    WHERE d23.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d24.`五大部門分類`,
    d24.`成本歸屬部門代碼`,
    d24.`成本歸屬部門`,
    d24.`計薪人數`,
    d24.`在職人數`,
    '個人提撥' AS 項目,
    CAST(d24.`個人提撥會科` AS INTEGER) AS 借方科目代號,
    CAST(d24.`個人提撥會科貸方` AS INTEGER) AS 貸方科目代號,
    d24.`個人提撥` AS NTD
    FROM 各單位科目與金額 d24
    WHERE d24.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d25.`五大部門分類`,
    d25.`成本歸屬部門代碼`,
    d25.`成本歸屬部門`,
    d25.`計薪人數`,
    d25.`在職人數`,
    '所得稅' AS 項目,
    CAST(d25.`所得稅會科` AS INTEGER) AS 借方科目代號,
    CAST(d25.`所得稅會科貸方` AS INTEGER) AS 貸方科目代號,
    d25.`所得稅` AS NTD
    FROM 各單位科目與金額 d25
    WHERE d25.`五大部門分類` IS NOT NULL
        
    UNION ALL
    SELECT  
    d26.`五大部門分類`,
    d26.`成本歸屬部門代碼`,
    d26.`成本歸屬部門`,
    d26.`計薪人數`,
    d26.`在職人數`,
    '就業安定費' AS 項目,
    CAST(d26.`就業安定費會科` AS INTEGER) AS 借方科目代號,
    CAST(d26.`就業安定費會科貸方` AS INTEGER) AS 貸方科目代號,
    d26.`就業安定費` AS NTD
    FROM 各單位科目與金額 d26
    WHERE d26.`五大部門分類` IS NOT NULL

    ) AS T1''')

    # 提交（保存）更改
    conn.commit()

    # 關閉連線
    conn.close()


    print('table轉換成功。')

    ## 匯出excel

    import pandas as pd
    import numpy as np

    # 創建或打開SQLite數據庫連接
    conn = sqlite3.connect(sqlite_db_path)

    c = conn.cursor()

    c.execute('''
    SELECT DISTINCT 項目 FROM 各單位科目與金額V2
    WHERE 借方科目代號 IS NOT NULL''')

    rows = c.fetchall()

    pages = []

    count = 0

    for row in rows:
        c.execute(f'''
        SELECT T1.五大部門分類,T1.成本歸屬部門代碼,T1.成本歸屬部門,T1.在職人數,T1.項目,case when NTD >= 0 then '借' else '貸' end 借貸方,T2.借方科目名稱,T1.借方科目代號,ABS(T1.NTD) NTD 
        FROM 各單位科目與金額V2 AS T1
        INNER JOIN 名稱對照 AS T2 ON T1.項目 = T2.名稱
        WHERE 項目 = '{row[0]}' AND T1.NTD != 0 AND T1.NTD IS NOT NULL
        ''')
        table = c.fetchall()
        df = pd.DataFrame(table)
        if df.empty:
            continue
        # print(df)
        total_sum = df[8].sum()
        # 添加新行
        new_row = pd.DataFrame({0:[''], 1: [''], 2: [''], 3: [''], 4: [''], 5: [''], 6: [''], 7:[''], 8: [total_sum]}) 
        df = pd.concat([df, new_row], ignore_index=True)
        pages.append(df)

        c.execute(f'''
        SELECT T1.五大部門分類,T1.成本歸屬部門代碼,T1.成本歸屬部門,T1.在職人數,T1.項目,case when NTD >= 0 then '貸' else '借' end 借貸方,T2.貸方科目名稱,T1.貸方科目代號,ABS(T1.NTD) NTD
        FROM 各單位科目與金額V2 AS T1
        INNER JOIN 名稱對照 AS T2 ON T1.項目 = T2.名稱
        WHERE 項目 = '{row[0]}' AND T1.NTD != 0 AND T1.NTD IS NOT NULL
        ''')
        table = c.fetchall()
        df = pd.DataFrame(table)
        if df.empty:
            continue
        # print(df)
        total_sum = df[8].sum()
        # 添加新行
        new_row = pd.DataFrame({0:[''], 1: [''], 2: [''], 3: [''], 4: [''], 5: [''], 6: [''], 7:[''], 8: [total_sum]}) 
        df = pd.concat([df, new_row], ignore_index=True)
        pages.append(df)


    dfx = pd.DataFrame({0:['五大部門分類'], 1: ['成本歸屬部門代碼'], 2: ['成本歸屬部門'], 3:['在職人數'], 4: ['項目'], 5:['借貸方'], 6:['科目名稱'], 7: ['科目代號'], 8:['金額']}) 

    for page in pages:
        dfx = pd.concat([dfx, page], ignore_index=True)



    dfx.to_excel(excel_file_output, index=False, header=False)
    conn.close()
