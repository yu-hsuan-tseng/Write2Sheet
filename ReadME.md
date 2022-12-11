## Description


本資料夾包含兩個主要執行的程式以及一隻定義function的程式

please run *tableau_daily_clickout_partner_1.py* and *run_sheet.py*

*warning . DO NOT RUN *write_to_sheet.py*

### 1. tableau_daily_clickout_partner_1.py 每日8:50AM 自動寄信到seanforlineec@gmail.com 
    若需要更改收件人信箱，於第236行修改 msg['To'] = xxx

### 2. run_sheet.py 每日9:00AM 自動執行26羅漢寫入google sheet 程式
    若要修改執行時間，於run_sheet.py 第13行更改 if hr == "xx" and mini == "xx":
    若需要修改function的功能，於write_to_sheet.py進行修改

### 3. 
    drive_key.json 用於串接 寫入google sheet 的api
    client_secrets2.json 用於Google Analytics api 
