# -*- coding: utf-8 -*-
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
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "curl_cffi", "-q"])
    from curl_cffi import requests as cffi_requests

# === إعداد الجلسة ===
session = cffi_requests.Session(impersonate="chrome110")

# === تحميل الـ Credentials (GitHub Actions / Colab) ===
creds = None

# 1. GitHub Actions → /tmp/creds.json
if os.path.exists('/tmp/creds.json'):
    creds = Credentials.from_service_account_file(
        '/tmp/creds.json',
        scopes=['https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive']
    )
    print("تم تحميل الـ JSON من GitHub Actions")

# 2. Google Colab → من Google Drive
elif 'google.colab' in str(get_ipython()):
    try:
        from google.colab import drive
        drive.mount('/content/drive')
        json_path = '/content/drive/MyDrive/stock-tracker-2025-c2e6fce3f1a7.json'
        if os.path.exists(json_path):
            creds = Credentials.from_service_account_file(
                json_path,
                scopes=['https://www.googleapis.com/auth/spreadsheets',
                        'https://www.googleapis.com/auth/drive']
            )
            print("تم تحميل الـ JSON من Google Drive (Colab)")
        else:
            raise FileNotFoundError(f"الملف غير موجود: {json_path}")
    except Exception as e:
        print(f"خطأ في تحميل JSON من Drive: {e}")
        raise

if creds is None:
    raise FileNotFoundError("لم يتم العثور على ملف creds.json")

# === الاتصال بجوجل شيت ===
client = gspread.authorize(creds)
spreadsheet_id = '1St7HrYmMsfqV5LgZRUm57JZxYopOP10cjS0oYPSC0UM'
sh = client.open_by_key(spreadsheet_id)
sheet = sh.sheet1

# ------------------------------------------------------------------
# 1. احتفظ بالعناوين فقط + احذف كل البيانات القديمة
# ------------------------------------------------------------------
def clear_old_data():
    all_vals = sheet.get_all_values()
    if len(all_vals) <= 1:                     # لا يوجد بيانات سوى العناوين
        return

    # احذف كل الصفوف من الصف 2 إلى نهاية الشيت
    rows_to_delete = len(all_vals) - 1
    try:
        # حذف جماعي (أسرع)
        sheet.delete_rows(2, 2 + rows_to_delete - 1)
        print(f"تم حذف {rows_to_delete} صفوف قديمة.")
    except Exception as e:
        print(f"فشل الحذف الجماعي: {e}")
        # fallback: حذف صفًا صفًا (بحد أقصى 100 صف لتجنب الـ timeout)
        for _ in range(min(rows_to_delete, 100)):
            sheet.delete_rows(2)
            time.sleep(0.3)

clear_old_data()

# ------------------------------------------------------------------
# 2. ضع العناوين الجديدة (مع السيولة، الدعم، المقاومة)
# ------------------------------------------------------------------
headers = ['التاريخ والوقت', 'اسم السهم', 'الرمز', 'السعر (ج.م)', 'التغيير (ج.م)', 'التغيير %', 
           'السيولة (حجم)', 'الدعم (Support)', 'المقاومة (Resistance)']
if sheet.row_values(1) != headers:
    sheet.update('A1:I1', [headers])
    sheet.format("A1:I1", {"textFormat": {"bold": True}})
    print("تم إضافة/تحديث العناوين (مع الإضافات الجديدة).")
else:
    print("العناوين موجودة مسبقًا.")

# ------------------------------------------------------------------
# 3. قاموس الأسماء بالعربي
# ------------------------------------------------------------------
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

print("\nجاري جلب الأسعار والسيولة والدعم/المقاومة...")

for symbol in stocks:
    try:
        ticker = yf.Ticker(symbol, session=session)
        # جلب تاريخ شهر كامل للدعم/المقاومة، ويومين للسعر الحالي
        hist_month = ticker.history(period="1mo", interval="1d", timeout=15)
        hist_day = ticker.history(period="2d", interval="1d", timeout=15)

        if hist_month.empty or len(hist_month) < 5 or hist_day.empty:
            raise ValueError("بيانات غير كافية")

        # السعر الحالي والتغيير
        latest = hist_day.iloc[-1]
        close = latest['Close']
        open_p = hist_day.iloc[0]['Open']  # افتتاح اليوم الأول في الفترة

        if pd.isna(close) or pd.isna(open_p) or close <= 0:
            raise ValueError("قيم غير صالحة")

        price = round(float(close), 2)
        change = round(float(close - open_p), 2)
        pct = round(float((close - open_p) / open_p * 100), 2)

        # السيولة: حجم التداول اليومي الأخير
        volume = latest['Volume'] if not pd.isna(latest['Volume']) else hist_day.iloc[-2]['Volume']
        volume = int(volume) if not pd.isna(volume) else 0

        # الدعم: أدنى سعر في الشهر (Low min)
        support = round(float(hist_month['Low'].min()), 2)

        # المقاومة: أعلى سعر في الشهر (High max)
        resistance = round(float(hist_month['High'].max()), 2)

        arabic_name = STOCK_NAMES_AR.get(symbol, symbol)

        new_rows.append([update_time, arabic_name, symbol, price, change, pct, volume, support, resistance])
        print(f"{symbol} - {arabic_name}: سعر {price} ج.م | سيولة {volume:,} | دعم {support} | مقاومة {resistance} ({pct:+.2f}%)")

    except Exception as e:
        skipped_stocks.append(symbol)
        print(f"{symbol}: تم تخطيه → {e}")

    time.sleep(1.5)   # لتجنب الحظر

# ------------------------------------------------------------------
# 4. إضافة البيانات الجديدة (آخر تحديث فقط)
# ------------------------------------------------------------------
if new_rows:
    sheet.append_rows(new_rows, value_input_option='RAW')
    print(f"\nتم إضافة {len(new_rows)} سهم (مع السيولة، الدعم، المقاومة)!")
else:
    print("\nتحذير: لا توجد أسهم متاحة اليوم.")

if skipped_stocks:
    print(f"\nالأسهم المُتخطاة ({len(skipped_stocks)}): {', '.join(skipped_stocks)}")
else:
    print("\nجميع الأسهم متاحة!")

print(f"\nتم التحديث في: {update_time}")
