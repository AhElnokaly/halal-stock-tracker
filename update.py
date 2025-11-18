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
session = cffi_requests.Session(impersonate="chrome120")

# === تحميل الـ Credentials ===
creds = None
if os.path.exists('/tmp/creds.json'):
    creds = Credentials.from_service_account_file('/tmp/creds.json',
        scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
    print("تم تحميل الـ JSON من GitHub Actions")
elif 'google.colab' in str(get_ipython()):
    try:
        from google.colab import drive
        drive.mount('/content/drive')
        json_path = '/content/drive/MyDrive/stock-tracker-2025-c2e6fce3f1a7.json'
        if os.path.exists(json_path):
            creds = Credentials.from_service_account_file(json_path,
                scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
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

# ==================================================================
# استكمال: تاب EGX مع قائمة كاملة 223 سهم + تصنيف دقيق + أسماء عربية
# ==================================================================

print("\n" + "="*70)
print("بدء تحديث تاب 'EGX' - جميع أسهم البورصة المصرية (223 سهم)")
print("="*70)

# --- 1. إنشاء أو فتح تاب EGX ---
try:
    egx_sheet = sh.worksheet("EGX")
    print("تم العثور على تاب 'EGX' موجود مسبقًا.")
except gspread.exceptions.WorksheetNotFound:
    egx_sheet = sh.add_worksheet(title="EGX", rows=350, cols=12)
    print("تم إنشاء تاب جديد باسم 'EGX'.")

# --- 2. مسح البيانات القديمة ---
def clear_egx_data():
    all_vals = egx_sheet.get_all_values()
    if len(all_vals) > 1:
        rows_to_delete = len(all_vals) - 1
        try:
            egx_sheet.delete_rows(2, 2 + rows_to_delete - 1)
            print(f"تم حذف {rows_to_delete} صف من تاب EGX.")
        except:
            for _ in range(min(rows_to_delete, 100)):
                egx_sheet.delete_rows(2)
                time.sleep(0.3)
clear_egx_data()

# --- 3. العناوين ---
egx_headers = ['التاريخ والوقت', 'اسم السهم', 'الرمز', 'السعر (ج.م)', 'التغيير (ج.م)', 
               'التغيير %', 'السيولة', 'الدعم', 'المقاومة', 'المؤشر']

if egx_sheet.row_values(1) != egx_headers:
    egx_sheet.update(values=[egx_headers], range_name='A1:J1')
    egx_sheet.format("A1:J1", {"textFormat": {"bold": True}, 
                               "backgroundColor": {"red": 0.1, "green": 0.5, "blue": 0.9}})
    print("تم تحديث عناوين تاب EGX.")

# --- 4. قائمة كاملة لـ 223 سهم (نوفمبر 2025) ---
all_stocks = [
    'COMI.CA', 'SWDY.CA', 'TMGH.CA', 'EAST.CA', 'ETEL.CA', 'QNBE.CA', 'EGAL.CA', 'MFPC.CA', 'ALCN.CA', 'ABUK.CA',
    'EFIH.CA', 'FWRY.CA', 'ORAS.CA', 'EMFD.CA', 'HDBK.CA', 'HRHO.CA', 'GPPL.CA', 'IRON.CA', 'EKHO.CA', 'EKHOA.CA',
    'FERC.CA', 'EFID.CA', 'BTFH.CA', 'ADIB.CA', 'GBCO.CA', 'ORHD.CA', 'CIEB.CA', 'JUFO.CA', 'CANA.CA', 'FAIT.CA',
    'FAITA.CA', 'SCTS.CA', 'PHDC.CA', 'OCDI.CA', 'EXPA.CA', 'EFIC.CA', 'ARCC.CA', 'EGCH.CA', 'VALU.CA', 'TAQA.CA',
    'SKPC.CA', 'UBEE.CA', 'SCEM.CA', 'CLHO.CA', 'ORWE.CA', 'HELI.CA', 'CCAP.CA', 'CCAPP.CA', 'MTIE.CA', 'MBSC.CA',
    'PHAR.CA', 'MCQE.CA', 'RAYA.CA', 'ATQA.CA', 'POUL.CA', 'EGSA.CA', 'TALM.CA', 'SAUD.CA', 'ISPH.CA', 'CIRA.CA',
    'CSAG.CA', 'MHOT.CA', 'MASR.CA', 'ELEC.CA', 'IFAP.CA', 'OLFI.CA', 'AMOC.CA', 'EGTS.CA', 'MOIL.CA', 'CICH.CA',
    'BINV.CA', 'AMES.CA', 'SUGR.CA', 'EGAS.CA', 'DOMT.CA', 'RMDA.CA', 'EGBE.CA', 'ACAP.CA', 'BONY.CA', 'ISMQ.CA',
    'MOIN.CA', 'ZMID.CA', 'ACRO.CA', 'CNFN.CA', 'MPRC.CA', 'BIOC.CA', 'OIH.CA', 'ENGC.CA', 'GDWA.CA', 'DSCW.CA',
    'NAPR.CA', 'PHTV.CA', 'SPHT.CA', 'GSSC.CA', 'MPCI.CA', 'PRCL.CA', 'AXPH.CA', 'CPCI.CA', 'MIPH.CA', 'KABO.CA',
    'SPIN.CA', 'UEFM.CA', 'WCDF.CA', 'SVCE.CA', 'GGRN.CA', 'SAIB.CA', 'MICH.CA', 'NIPH.CA', 'MFSC.CA', 'OFH.CA',
    'AJWA.CA', 'ADCI.CA', 'UNIT.CA', 'KZPC.CA', 'ARAB.CA', 'NINH.CA', 'ELKA.CA', 'AFMC.CA', 'ASCM.CA', 'OCPH.CA',
    'ACGC.CA', 'AMIA.CA', 'LCSW.CA', 'NAHO.CA', 'ELSH.CA', 'EDFM.CA', 'CFGH.CA', 'EALR.CA', 'DAPH.CA', 'SMFR.CA',
    'WKOL.CA', 'ECAP.CA', 'RACC.CA', 'PHGC.CA', 'CEFM.CA', 'SCFM.CA', 'MPCO.CA', 'NHPS.CA', 'AMER.CA', 'ADPC.CA',
    'ATLC.CA', 'INFI.CA', 'EHDR.CA', 'ACAMD.CA', 'SNFC.CA', 'IDRE.CA', 'ETRS.CA', 'DEIN.CA', 'AALR.CA', 'MILS.CA',
    'NCCW.CA', 'ALRA.CA', 'KRDI.CA', 'CERA.CA', 'MENA.CA', 'SDTI.CA', 'MOSC.CA', 'ISMA.CA', 'MEPA.CA', 'ODIN.CA',
    'ZEOT.CA', 'UEGC.CA', 'POCO.CA', 'SEIG.CA', 'SEIGA.CA', 'NARE.CA', 'APSW.CA', 'RTVC.CA', 'GGCC.CA', 'COSG.CA',
    'MAAL.CA', 'CAED.CA', 'GTEX.CA', 'ICLE.CA', 'MEGM.CA', 'MOED.CA', 'RAKT.CA', 'PRCL.CA'
]

print(f"تم تحميل {len(all_stocks)} سهم كاملة.")

# --- 5. قوائم المؤشرات (محدثة 2025) ---
EGX30_SYMBOLS = {'SWDY.CA', 'COMI.CA', 'FWRY.CA', 'HRHO.CA', 'TMGH.CA', 'ETEL.CA', 'MFPC.CA', 'EKHO.CA', 'ADIB.CA',
                 'SIDI.CA', 'PHDC.CA', 'ORHD.CA', 'EAST.CA', 'HELI.CA', 'MCQE.CA', 'EGAL.CA', 'SKPC.CA', 'EMFD.CA',
                 'ORWE.CA', 'MTIE.CA', 'EXPA.CA', 'AUTO.CA', 'JUFO.CA', 'POUL.CA', 'EFIC.CA', 'ELEC.CA', 'CESI.CA',
                 'AIND.CA', 'ARAB.CA', 'CICH.CA'}

EGX70_SYMBOLS = {'ACTF.CA', 'KRDI.CA', 'ASCM.CA', 'ATQA.CA', 'PRMH.CA', 'OBRI.CA', 'SCEM.CA', 'EGS.CA', 'MOSC.CA', 'ZEOT.CA',
                 'ACRO.CA', 'SPMD.CA', 'AXPH.CA', 'CLHO.CA', 'MCRO.CA', 'NDRL.CA', 'FAIT.CA', 'AMIA.CA', 'GIHD.CA', 'MBSC.CA',
                 'CAED.CA', 'AIFI.CA', 'DCRC.CA', 'CCRS.CA', 'AJWA.CA', 'SAUD.CA', 'CAN.CA', 'RACC.CA', 'RAYA.CA', 'BTFH.CA',
                 'ISMQ.CA', 'INRT.CA', 'EGID.CA', 'SEIG.CA', 'OLFI.CA', 'KZVH.CA', 'APSA.CA', 'ARPI.CA', 'ASPI.CA', 'ATLC.CA',
                 'BIO.CA', 'BIOC.CA', 'CIEB.CA', 'CIRA.CA', 'CNIN.CA', 'DOMT.CA', 'DSCW.CA', 'ECAP.CA', 'EDBM.CA', 'EFIH.CA',
                 'EGCH.CA', 'EGTS.CA', 'EHDR.CA', 'EIUD.CA', 'ELKA.CA', 'ELSH.CA', 'EMRI.CA', 'ENGC.CA', 'EPPK.CA', 'ESRS.CA',
                 'ETRS.CA', 'FAIT.CA', 'FIRE.CA', 'GOCO.CA', 'GRCA.CA', 'GTHE.CA', 'ICID.CA', 'ICMI.CA', 'IDRE.CA', 'IFAP.CA',
                 'INEG.CA', 'IRAX.CA', 'ISPH.CA', 'KWIN.CA', 'LCSW.CA', 'MENA.CA', 'MEPA.CA', 'MICH.CA', 'MILS.CA', 'MIPR.CA',
                 'MIRA.CA', 'MIXC.CA', 'MNHD.CA', 'MOIL.CA', 'MPCI.CA', 'MPRC.CA', 'MTRS.CA', 'NAMA.CA', 'OCDI.CA', 'ODIN.CA',
                 'OIH.CA', 'PACH.CA', 'PHAR.CA', 'PHTV.CA', 'PORT.CA', 'PRCL.CA', 'QNBA.CA', 'RAKT.CA', 'REAC.CA', 'RMDA.CA',
                 'SAFM.CA', 'SCFM.CA', 'SIPC.CA', 'SMFR.CA', 'SNFC.CA', 'SPHT.CA', 'SRWA.CA', 'SUCE.CA', 'SVOT.CA', 'TALM.CA',
                 'TASF.CA', 'TORA.CA', 'TRST.CA', 'UAEC.CA', 'UPI.CA', 'VODE.CA', 'WATP.CA', 'WCDF.CA', 'WKOL.CA', 'ZMID.CA'}

# --- 6. قاموس الأسماء العربية والإنجليزية (موسع لأكثر من 200 سهم) ---
NAMES = {
    'AALR.CA': 'العامة لاستصلاح الأراضي',  # Arabic
    'ABUK.CA': 'Abou Kir Fertilizers & Chemical Industries Co.',  # English
    'ACAMD.CA': 'Arab Co. for Asset Management & Development',
    'ACAP.CA': 'إيه كابيتال القابضة',
    'ACGC.CA': 'Arab Cotton Ginning Co.',
    'ACRO.CA': 'اكرو مصر',
    'ACTF.CA': 'اكت فاينانشال',
    'ADCI.CA': 'Arab Pharmaceuticals',
    'ADIB.CA': 'Abu Dhabi Islamic Bank-Egypt',
    'ADPC.CA': 'Arab Dairy Products Co. Arab Dairy - Panda',
    'ADRI.CA': 'أراب للتنمية و الاستثمار العقاري',
    'AFDI.CA': 'Alahly For Development & Investment',
    'AFMC.CA': 'Alexandria Flour Mills Co.',
    'AIDC.CA': 'Arabia for Investment and Development',
    'AIFI.CA': 'Atlas for Investment & Food Industries SAE',
    'AIH.CA': 'Arabia Investments Holding SAE',
    'AJWA.CA': 'Ajwa for Food Industries Co. Egypt',
    'ALCN.CA': 'الاسكندرية لتداول الحاويات',
    'ALUM.CA': 'Arab Aluminum Co. SAE',
    'AMER.CA': 'Amer Group Holding',
    'AMES.CA': 'Alexandria New Medical Center Co.',
    'AMIA.CA': 'Arab Moltaqa Investments Company',
    'AMOC.CA': 'أموك',
    'AMPI.CA': 'AL Moasher Pay for Electronic Payment and Collection (S.A.E)',
    'ANFI.CA': 'Alexandria National Co. for Financial Investment',
    'APSW.CA': 'Arab Polvara Spinning & Weaving Co.',
    'ARAB.CA': 'Arab Developers Holding',
    'ARCC.CA': 'Arabian Cement Company',
    'AREH.CA': 'Egyptian Real Estate Group',
    'ARVA.CA': 'Arab Valves Co.',
    'ASCM.CA': 'ASEC Co. for Mining',
    'ASPI.CA': 'Aspire Capital Holding for Financial Investments',
    'ATLC.CA': 'Al Tawfeek Leasing Company-A.T.LEASE',
    'ATQA.CA': 'Misr National Steel',
    'AXPH.CA': 'Alexandria Company for Pharmaceuticals and Chemical Industries',
    'BIDI.CA': 'El Badr Investment and Development - BID',
    'BIGP.CA': 'ElBarbary Investment Group',
    'BINV.CA': 'B Investments Holding SAE',
    'BIOC.CA': 'GlaxoSmithKline S.A.E.',
    'BONY.CA': 'Bonyan for Development and Trade',
    'BTFH.CA': 'Beltone Holding',
    'CAED.CA': 'Cairo Educational Services',
    'CANA.CA': 'Suez Canal Bank SAE',
    'CCAP.CA': 'QALA For Financial Investments',
    'CCAPP.CA': 'QALA For Financial Investments - Preferred',
    'CCRS.CA': 'Gulf Canadian Real Estate Investment Co.',
    'CEFM.CA': 'Middle Egypt Flour Mills',
    'CERA.CA': 'Arab Ceramic Co. - Ceramica Remas',
    'CFGH.CA': 'Concrete Fashion Group for Commercial and Industrial Investments S.A.E',
    'CICH.CA': 'CI Capital Holding for Financial Investments',
    'CIEB.CA': 'Credit Agricole Egypt',
    'CIRA.CA': 'Cairo For Investment And Real Estate Developments -CIRA Education',
    'CLHO.CA': 'Cleopatra Hospital Company',
    'CNFN.CA': 'Contact Financial Holding SAE',
    'COMI.CA': 'Commercial International Bank - Egypt (CIB) S.A.E.',
    'COSG.CA': 'Cairo Oils & Soap',
    'CPCI.CA': 'Kahira Pharmaceuticals & Chemical Industries Co.',
    'CSAG.CA': 'Canal Shipping Agencies Co.',
    'DAPH.CA': 'Development & Engineering Consultants',
    'DEIN.CA': 'Delta Insurance',
    'DOMT.CA': 'Arabian Food Industries Co.',
    'DSCW.CA': 'Dice Sports & Casual Wear Manufacturers SAE',
    'EALR.CA': 'El Arabia for Land Reclamation',
    'EAST.CA': 'Eastern Company',
    'ECAP.CA': 'El Ezz Ceramics & Porcelain Co. (Gemma)',
    'EDFM.CA': 'East Delta Flour Mills Co.',
    'EFIC.CA': 'Egyptian Financial & Industrial Co.',
    'EFID.CA': 'Edita Food Industries SAE',
    'EFIH.CA': 'e-finance for Digital and Financial Investments S.A.E.',
    'EGAL.CA': 'Egypt Aluminum',
    'EGAS.CA': 'Egypt Gas Co.',
    'EGCH.CA': 'Egyptian Chemical Industries',
    'EGSA.CA': 'Egyptian Satellite Co.',
    'EGTS.CA': 'Egyptian for Tourism Resorts',
    'EHDR.CA': 'Egyptians Housing Development & Reconstruction',
    'EKHO.CA': 'Egypt Kuwait Holding Co. SAE',
    'EKHOA.CA': 'Egypt Kuwait Holding Co. SAE - Preferred',
    'ELEC.CA': 'Electro Cable Egypt',
    'ELKA.CA': 'El Kahera Housing',
    'ELSH.CA': 'El-Shams Housing & Development SA',
    'EMFD.CA': 'Emaar Misr for Development SAE',
    'ENGC.CA': 'Industrial Engineering Co. for Construction & Development',
    'ESRS.CA': 'Ezz Steel',
    'ETEL.CA': 'Telecom Egypt',
    'ETRS.CA': 'Egyptian Iron & Steel',
    'EXPA.CA': 'Export Development Bank of Egypt',
    'FAIT.CA': 'Faisal Islamic Bank of Egypt',
    'FAITA.CA': 'Faisal Islamic Bank of Egypt - Preferred',
    'FERC.CA': 'Fertilizers & Chemicals',
    'FWRY.CA': 'Fawry for Banking Technology and Electronic Payments SAE',
    'GBCO.CA': 'GB AUTO',
    'GDWA.CA': 'Gadwa For Industrial Development',
    'GGCC.CA': 'Giza General Contracting & Real Estate Investment',
    'GIHD.CA': 'غزل المحلة',  # Keep Arabic if preferred
    'GPPL.CA': 'Golden Pyramids Plaza',
    'GSSC.CA': 'General Silos & Storage Co.',
    'GTEX.CA': 'GTEX For Commercial And Industrial Investments',
    'HDBK.CA': 'Housing & Development Bank',
    'HELI.CA': 'Heliopolis Housing',
    'HRHO.CA': 'EFG Hermes Holding SAE',
    'ICLE.CA': 'International Co. For Leasing (InLease)',
    'IDRE.CA': 'Arab Real Estate Investment Co. (ALICO)',
    'IFAP.CA': 'International Agricultural Products',
    'INFI.CA': 'Ismailia National Co. for Food Industries',
    'IRON.CA': 'Egyptian Iron & Steel',
    'ISMA.CA': 'Ismailia Misr Poultry',
    'ISMQ.CA': 'Iron And Steel for Mines And Quarries',
    'ISPH.CA': 'Ibn Sina Pharma',
    'JUFO.CA': 'Juhayna Food Industries',
    'KABO.CA': 'El Nasr Clothing & Textiles (Kabo)',
    'KRDI.CA': 'Al-Khair River For Development Agricultural Investment & Environmental Services Company',
    'KZPC.CA': 'Kafr El Zayat Pesticides',
    'LCSW.CA': 'Lecico Egypt SAE',
    'MAAL.CA': 'Marseille Almasreia Alkhalegeya For Holding Investment SAE',
    'MASR.CA': 'Madinet Masr For Housing and Development',
    'MBSC.CA': 'Misr Beni Suef Cement',
    'MCQE.CA': 'Misr Cement (Qena)',
    'MEGM.CA': 'Middle East Glass Manufacturing Co.',
    'MEPA.CA': 'Medical Packaging Company',
    'MENA.CA': 'Mena Touristic & Real Estate Investment',
    'MFPC.CA': 'Misr Chemical Industries',
    'MFSC.CA': 'Egypt For Poultry',
    'MHOT.CA': 'Misr Hotels',
    'MICH.CA': 'Misr Chemical Industries',
    'MILS.CA': 'North Cairo Mills',
    'MIPH.CA': 'Misr Intercontinental for Granite & Marble (EGY-STON)',
    'MMAT.CA': 'M.M Group for Industry & International Trade',
    'MNHD.CA': 'Madinet Nasr Housing & Development',
    'MOED.CA': 'Egyptian Modern Education Systems',
    'MOIL.CA': 'Maridive & Oil Services SAE',
    'MOSC.CA': 'Misr Oils & Soap',
    'MPCO.CA': 'Mansourah Poultry',
    'MPCI.CA': 'Memphis Pharmaceuticals & Chemical Industries',
    'MPRC.CA': 'Egyptian Media Production City',
    'MTIE.CA': 'MM Group for Industry and International Trade SAE',
    'NAHO.CA': 'Nasr City Housing',
    'NCCW.CA': 'Nasr Company for Civil Works',
    'NHPS.CA': 'National Housing for Professional Syndicates',
    'NINH.CA': 'Nozha International Hospital',
    'NIPH.CA': 'Ebn Sina for Medical',
    'OCDI.CA': 'Six of October Development & Investment',
    'OCPH.CA': 'October Pharma SAE',
    'ODIN.CA': 'El Orouba for Development & Investment',
    'OFH.CA': 'Orascom Financial Holding',
    'OIH.CA': 'Orascom Investment Holding SAE',
    'OLFI.CA': 'Obour Land for Food Industries SAE',
    'ORAS.CA': 'Orascom Construction PLC',
    'ORHD.CA': 'Orascom Development Egypt SAE',
    'ORWE.CA': 'Oriental Weavers',
    'PHAR.CA': 'Egyptian International Pharmaceuticals (EIPICO)',
    'PHDC.CA': 'Palm Hills Developments SAE',
    'PHGC.CA': 'Port Said Agricultural Development & Construction',
    'POCO.CA': 'Port Saied for Development & Construction',
    'POUL.CA': 'Cairo Poultry',
    'PRCL.CA': 'General Company For Ceramic & Porcelain Products',
    'QNBE.CA': 'Qatar National Bank Alahly',
    'RACC.CA': 'Raya Contact Center',
    'RAKT.CA': 'Rakta Paper Manufacturing',
    'RAYA.CA': 'Raya Holding for Financial Investments SAE',
    'RMDA.CA': 'Tenth of Ramadan Pharmaceutical Industries&Diagnostic-Rameda',
    'RTVC.CA': 'Remco for Touristic Villages Construction',
    'SAIB.CA': 'Societe Arabe Internationale de Banque SAE',
    'SAUD.CA': 'Saudi Egyptian Investment & Finance',
    'SCEM.CA': 'South Cairo & Giza Mills & Bakeries',
    'SCFM.CA': 'South Cairo & Giza Mills & Bakeries',
    'SCTS.CA': 'Suez Canal Company For Technology Settling',
    'SDTI.CA': 'Sharm Dreams Co. for Tourism Investment',
    'SEIG.CA': 'El Sewedy Electric Co.',
    'SEIGA.CA': 'El Sewedy Electric Co. - Preferred',
    'SMFR.CA': 'Samad Misr - EGYFERT',
    'SNFC.CA': 'Sharkia National Food',
    'SPHT.CA': 'El Shams Pyramids For Hotels&Touristic Projects',
    'SPIN.CA': 'Alexandria Spinning & Weaving (SPINALEX)',
    'SUGR.CA': 'Delta Sugar',
    'SVCE.CA': 'South Valley Cement',
    'SWDY.CA': 'El Sewedy Electric Co.',
    'TALM.CA': 'Taaleem Management Services',
    'TMGH.CA': 'Talaat Moustafa Group',
    'UEFM.CA': 'Upper Egypt Flour Mills',
    'UEGC.CA': 'United for Housing & Development',
    'UNIT.CA': 'United Housing & Development',
    'WCDF.CA': 'Middle & West Delta Flour Mills',
    'WKOL.CA': 'Zawya for Industrial Development',
    'ZMID.CA': 'Zahraa El Maadi Investment & Development Co SAE',
    # Add more from previous if needed
}

# --- 7. جلب البيانات ---
egx_update_time = datetime.now().strftime('%Y-%m-%d %H:%M')
egx_rows = []
failed = []

print(f"\nجاري جلب بيانات {len(all_stocks)} سهم... (قد يستغرق 5-10 دقائق)")

for symbol in all_stocks:
    try:
        ticker = yf.Ticker(symbol, session=session)
        hist_month = ticker.history(period="1mo", interval="1d")
        hist_day = ticker.history(period="2d", interval="1d")

        if hist_month.empty or hist_day.empty or len(hist_month) < 3:
            raise ValueError("بيانات غير كافية")

        latest = hist_day.iloc[-1]
        prev_open = hist_day.iloc[0]['Open']
        close = latest['Close']
        volume = latest['Volume']

        if pd.isna(close) or close <= 0:
            raise ValueError("سعر غير صالح")

        price = round(close, 2)
        change = round(close - prev_open, 2)
        pct = round((close - prev_open) / prev_open * 100, 2)
        volume = int(volume) if not pd.isna(volume) else 0
        support = round(hist_month['Low'].min(), 2)
        resistance = round(hist_month['High'].max(), 2)

        name = NAMES.get(symbol, "غير معروف")
        if name == "غير معروف":
            try:
                english_name = ticker.info.get('longName', 'Unknown')
                name = english_name
            except:
                name = "Unknown"

        if symbol in EGX30_SYMBOLS:
            index = "EGX30"
        elif symbol in EGX70_SYMBOLS:
            index = "EGX70"
        elif symbol in EGX30_SYMBOLS | EGX70_SYMBOLS:  # EGX100 هو دمج
            index = "EGX100"
        else:
            index = "غير مصنف"

        egx_rows.append([egx_update_time, name, symbol.replace('.CA',''), price, change, pct, volume, support, resistance, index])

        print(f"{symbol.replace('.CA',''):>8} │ {price:>6} │ {pct:+6.2f}% │ {index}")

    except Exception as e:
        failed.append(symbol)
        print(f"{symbol}: تم تخطيه → {str(e)[:50]}")
    time.sleep(1.2)

# --- 8. كتابة البيانات ---
if egx_rows:
    egx_sheet.append_rows(egx_rows, value_input_option='RAW')
    print(f"\nتم بنجاح تحديث تاب 'EGX' بـ {len(egx_rows)} سهم!")
else:
    print("\nتحذير: لم يتم جلب أي بيانات!")

# --- 9. مراجعة الشيت وتحديث الأسماء "غير معروف" أو "Unknown" إلى الإنجليزي ---
print("\nجاري مراجعة الشيت وتحديث الأسماء غير المعروفة...")
all_vals = egx_sheet.get_all_values()
updated_count = 0

for row_idx, row in enumerate(all_vals[1:], start=2):  # من الصف 2 (البيانات)
    current_name = row[1]  # العمود B (index 1)
    if current_name in ["غير معروف", "Unknown"]:  
        symbol = row[2] + '.CA'  # العمود C (index 2)
        try:
            ticker = yf.Ticker(symbol, session=session)
            english_name = ticker.info.get('longName', 'Unknown')
            egx_sheet.update_cell(row_idx, 2, english_name)  # تحديث العمود B
            print(f"تم تحديث {symbol.replace('.CA', '')}: {english_name}")
            updated_count += 1
            time.sleep(0.5)  # تأخير لتجنب الحظر
        except Exception as e:
            print(f"فشل تحديث {symbol}: {e}")

print(f"\nتم تحديث {updated_count} اسم سهم غير معروف إلى الإنجليزي!")

if failed:
    print(f"الأسهم الفاشلة ({len(failed)}): {', '.join([s.replace('.CA','') for s in failed[:20]])}...")
else:
    print("تم جلب جميع الأسهم بنجاح!")

print(f"\nتم التحديث النهائي في: {egx_update_time}")
print("="*70)
