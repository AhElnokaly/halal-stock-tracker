import yfinance as yf
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import time
import os

# === تثبيت واستخدام curl_cffi ===
try:
    from curl_cffi import requests as cffi_requests
except ImportError:
    print("جاري تثبيت curl_cffi...")
    import subprocess
    subprocess.run(["pip", "install", "curl_cffi", "-q"], check=True)
    from curl_cffi import requests as cffi_requests

# === إعداد الجلسة بـ curl_cffi ===
session = cffi_requests.Session(impersonate="chrome110")
yf.pdr_override = lambda: None

# === تحميل الـ Credentials ===
if os.path.exists('/tmp/creds.json'):
    creds = Credentials.from_service_account_file('/tmp/creds.json', scopes=[
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ])
    print("تم تحميل الـ JSON من GitHub Actions")
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

# === الاتصال بجوجل شيت ===
client = gspread.authorize(creds)
spreadsheet_id = '1St7HrYmMsfqV5LgZRUm57JZxYopOP10cjS0oYPSC0UM'
sh = client.open_by_key(spreadsheet_id)
sheet = sh.sheet1

# === قائمة العطلات الرسمية في مصر 2025 ===
EGYPT_HOLIDAYS_2025 = [
    "2025-01-07", "2025-01-25", "2025-04-20", "2025-04-21",
    "2025-05-01", "2025-06-30", "2025-07-01", "2025-07-23",
    "2025-09-27", "2025-10-06"
]

today = datetime.now().date()
today_str = today.strftime('%Y-%m-%d')

# === التحقق: من الأحد إلى الخميس فقط + لا عطلة رسمية ===
if today.weekday() >= 5:  # جمعة أو سبت
    print(f"اليوم {today} هو جمعة أو سبت → لا يوجد تداول.")
    raise SystemExit("خارج أيام العمل.")

if today_str in EGYPT_HOLIDAYS_2025:
    print(f"اليوم {today} عطلة رسمية → لا يوجد تداول.")
    raise SystemExit("عطلة رسمية.")

print(f"اليوم {today} ضمن أيام العمل → جاري التحديث...")

# === تنظيف الشيت لو أكتر من 1000 صف ===
rows = len(sheet.get_all_values())
if rows > 1000:
    print(f"عدد الصفوف = {rows} → سيتم حذف {rows - 1000} صفوف قديمة.")
    keep = 1000
    delete_count = rows - keep
    try:
        sheet.delete_rows(2, 2 + delete_count - 1)
        print("تم تنظيف الشيت بنجاح (دفعة واحدة).")
    except Exception as e:
        print(f"فشل الحذف الجماعي: {e}")
        for _ in range(min(delete_count, 100)):
            sheet.delete_rows(2)
            time.sleep(0.5)

# === قائمة الأسهم ===
stocks = [
    'ADIB.CA', 'SAUD.CA', 'AMIA.CA', 'ATLC.CA', 'FAITA.CA',
    'AIFI.CA', 'CAED.CA', 'GIHD.CA', 'MBSC.CA', 'AMES.CA',
    'DCRC.CA', 'ZEOT.CA', 'MOSC.CA', 'CCRS.CA', 'NDRL.CA',
    'CLHO.CA', 'MCRO.CA', 'AXPH.CA', 'AJWA.CA', 'SPMD.CA'
]

update_time = datetime.now().strftime('%Y-%m-%d %H:%M')
new_rows = []
skipped_stocks = []

print("\nجاري جلب الأسعار...")

for symbol in stocks:
    try:
        ticker = yf.Ticker(symbol, session=session)
        hist = ticker.history(period="5d", timeout=15)

        if hist.empty:
            raise ValueError("لا توجد بيانات")

        latest = hist.iloc[-1]
        close = latest['Close']
        open_p = latest['Open']

        if pd.isna(close) or pd.isna(open_p) or close <= 0:
            raise ValueError("قيم غير صالحة")

        price = round(float(close), 2)
        change = round(float(close - open_p), 2)
        pct = round(float((close - open_p) / open_p * 100), 2)

        # إضافة فقط إذا كان السهم متاح
        new_rows.append([update_time, symbol, price, change, pct])
        print(f"{symbol}: {price} ج.م ({pct:+.2f}%)")

    except Exception as e:
        skipped_stocks.append(symbol)
        print(f"{symbol}: تم تخطيه (غير متاح)")

    time.sleep(1.5)

# === إضافة العناوين (5 أعمدة فقط) ===
headers = ['التاريخ والوقت', 'الرمز', 'السعر (ج.م)', 'التغيير (ج.م)', 'التغيير %']
current_headers = sheet.row_values(1)

if not current_headers or current_headers != headers:
    if current_headers:
        sheet.delete_rows(1)
    sheet.insert_row(headers, 1)
    sheet.format("A1:E1", {"textFormat": {"bold": True}})
    print("تم إضافة/تحديث العناوين.")

# === إضافة البيانات المتاحة فقط ===
if new_rows:
    sheet.append_rows(new_rows, value_input_option='RAW')
    print(f"\nتم إضافة {len(new_rows)} سهم متاح بنجاح!")
else:
    print("\nتحذير: لا توجد أسهم متاحة للإضافة اليوم.")

# === طباعة ملخص الأسهم المُتخطاة ===
if skipped_stocks:
    print(f"\nالأسهم المُتخطاة ({len(skipped_stocks)}): {', '.join(skipped_stocks)}")
else:
    print("\nجميع الأسهم متاحة!")

print(f"\nتم التحديث في: {update_time}")
