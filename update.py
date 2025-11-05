import yfinance as yf
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import time
import os

# === تثبيت curl_cffi ===
try:
    from curl_cffi import requests as cffi_requests
except ImportError:
    print("جاري تثبيت curl_cffi...")
    import subprocess
    subprocess.run(["pip", "install", "curl_cffi", "-q"], check=True)
    from curl_cffi import requests as cffi_requests

# === إعداد الجلسة ===
session = cffi_requests.Session(impersonate="chrome110")
yf.pdr_override = lambda: None

# === تحميل الـ Credentials ===
creds = None

# 1. في GitHub Actions → /tmp/creds.json
if os.path.exists('/tmp/creds.json'):
    creds = Credentials.from_service_account_file('/tmp/creds.json', scopes=[
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ])
    print("تم تحميل الـ JSON من GitHub Actions")

# 2. في Google Colab → من Google Drive
elif 'google.colab' in str(get_ipython()):
    try:
        from google.colab import drive
        drive.mount('/content/drive')
        json_path = '/content/drive/MyDrive/stock-tracker-2025-c2e6fce3f1a7.json'  # غيّر المسار لو حابب
        if os.path.exists(json_path):
            creds = Credentials.from_service_account_file(json_path, scopes=[
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ])
            print("تم تحميل الـ JSON من Google Drive (Colab)")
        else:
            raise FileNotFoundError(f"الملف غير موجود: {json_path}")
    except Exception as e:
        print(f"خطأ في تحميل JSON من Drive: {e}")
        raise

# إذا مفيش أي مصدر → خطأ
if creds is None:
    raise FileNotFoundError("لم يتم العثور على ملف creds.json في /tmp أو Google Drive")

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
if today.weekday() >= 5:
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

# === قاموس الأسماء بالعربي ===
STOCK_NAMES_AR = {
    'ADIB.CA': 'بنك أبوظبي الإسلامي',
    'SAUD.CA': 'البنك السعودي للاستثمار',
    'AMIA.CA': 'الإسكندرية للاستثمارات',
    'ATLC.CA': 'أطلس للاستثمار',
    'FAITA.CA': 'فيتا للصناعات',
    'AIFI.CA': 'الإسكندرية للاستثمار',
    'CAED.CA': 'القاهرة للاستثمار',
    'GIHD.CA': 'جي آي إتش دي',
    'MBSC.CA': 'مصر بني سويف',
    'AMES.CA': 'أميس',
    'DCRC.CA': 'ديجيتال كابيتال',
    'ZEOT.CA': 'زيوت مصر',
    'MOSC.CA': 'مصر للأسمنت',
    'CCRS.CA': 'القاهرة للاستثمار',
    'NDRL.CA': 'الدلتا للسكر',
    'CLHO.CA': 'كلين هاوس',
    'MCRO.CA': 'مايكرو',
    'AXPH.CA': 'الإسكندرية للأدوية',
    'AJWA.CA': 'أجوا',
    'SPMD.CA': 'سبيد ميديكال'
}

stocks = list(STOCK_NAMES_AR.keys())
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
        arabic_name = STOCK_NAMES_AR.get(symbol, symbol)

        new_rows.append([update_time, arabic_name, symbol, price, change, pct])
        print(f"{symbol} - {arabic_name}: {price} ج.م ({pct:+.2f}%)")

    except Exception as e:
        skipped_stocks.append(symbol)
        print(f"{symbol}: تم تخطيه (غير متاح)")

    time.sleep(1.5)

# === إضافة العناوين ===
headers = ['التاريخ والوقت', 'اسم السهم', 'الرمز', 'السعر (ج.م)', 'التغيير (ج.م)', 'التغيير %']
current_headers = sheet.row_values(1)

if not current_headers or current_headers != headers:
    if current_headers:
        sheet.delete_rows(1)
    sheet.insert_row(headers, 1)
    sheet.format("A1:F1", {"textFormat": {"bold": True}})
    print("تم إضافة/تحديث العناوين.")

# === إضافة البيانات ===
if new_rows:
    sheet.append_rows(new_rows, value_input_option='RAW')
    print(f"\nتم إضافة {len(new_rows)} سهم متاح بنجاح!")
else:
    print("\nتحذير: لا توجد أسهم متاحة اليوم.")

if skipped_stocks:
    print(f"\nالأسهم المُتخطاة ({len(skipped_stocks)}): {', '.join(skipped_stocks)}")
else:
    print("\nجميع الأسهم متاحة!")

print(f"\nتم التحديث في: {update_time}")
