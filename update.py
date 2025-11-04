# update.py - يضيف صف يومي تلقائيًا
import yfinance as yf
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import time
import os
import json
import numpy as np

# === تحميل الـ JSON ===
creds = None

# 1. من GitHub Secret
creds_json = os.getenv('GOOGLE_SERVICE_ACCOUNT')
if creds_json:
    with open('/tmp/creds.json', 'w') as f:
        f.write(creds_json)
    creds = Credentials.from_service_account_file('/tmp/creds.json', scopes=[
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ])
    print("تم تحميل الـ JSON من GitHub")

# 2. من Google Drive (Colab فقط)
else:
    try:
        from google.colab import drive
        drive.mount('/content/drive')
        json_path = '/content/drive/MyDrive/stock-tracker-2025-c2e6fce3f1a7.json'
        if os.path.exists(json_path):
            creds = Credentials.from_service_account_file(json_path, scopes=[
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ])
            print("تم تحميل الـ JSON من Google Drive")
        else:
            raise FileNotFoundError
    except Exception as e:
        raise ValueError(f"فشل تحميل الـ JSON: {e}")

# === إعداد الشيت ===
client = gspread.authorize(creds)
spreadsheet_id = os.getenv('SPREADSHEET_ID', '1St7HrYmMsfqV5LgZRUm57JZxYopOP10cjS0oYPSC0UM')
sh = client.open_by_key(spreadsheet_id)
sheet = sh.sheet1

# === الأسهم ===
stocks = [
    'ADIB.CA', 'SAUD.CA', 'AMIA.CA', 'ATLC.CA', 'FAITA.CA',
    'AIFI.CA', 'CAED.CA', 'GIHD.CA', 'MBSC.CA', 'AMES.CA',
    'DCRC.CA', 'ZEOT.CA', 'MOSC.CA', 'CCRS.CA', 'NDRL.CA',
    'CLHO.CA', 'MCRO.CA', 'AXPH.CA', 'AJWA.CA', 'SPMD.CA'
]

# === جلب البيانات ===
update_time = datetime.now().strftime('%Y-%m-%d %H:%M')
new_rows = []

for symbol in stocks:
    try:
        ticker = yf.Ticker(symbol)
        hist = ticker.history(period="2d")
        if hist.empty:
            raise ValueError("لا بيانات")
        latest = hist.iloc[-1]
        close = latest['Close']
        open_p = latest['Open']
        if pd.isna(close) or pd.isna(open_p) or close <= 0:
            raise ValueError("قيمة غير صالحة")
        price = round(float(close), 2)
        change = round(float(close - open_p), 2)
        pct = round(float((close - open_p) / open_p * 100), 2)
        new_rows.append([update_time, symbol, price, change, pct])
        print(f"{symbol}: {price} ج.م")
    except Exception as e:
        new_rows.append([update_time, symbol, "غير متاح", 0, 0])
    time.sleep(0.6)

# === إضافة الصفوف ===
sheet.append_rows(new_rows, value_input_option='RAW')
print(f"تم إضافة {len(new_rows)} صف جديد!")

# === تنسيق العناوين (مرة واحدة) ===
try:
    if not sheet.row_values(1):
        headers = ['التاريخ والوقت', 'الرمز', 'السعر (ج.م)', 'التغيير (ج.م)', 'التغيير %']
        sheet.insert_row(headers, 1)
        sheet.format("A1:E1", {"textFormat": {"bold": True}})
except:
    pass
