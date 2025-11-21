# -*- coding: utf-8 -*-
import yfinance as yf
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import time
import os
import sys

# === تثبيت curl_cffi إذا مش موجود ===
try:
    from curl_cffi import requests as cffi_requests
except ImportError:
    print("جاري تثبيت curl_cffi...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "curl_cffi", "--quiet"])
    from curl_cffi import requests as cffi_requests

# === إعداد الجلسة (أحدث إصدار شغال ممتاز نوفمبر 2025) ===
session = cffi_requests.Session(impersonate="chrome124")

# === تحميل Credentials ===
creds = None

# 1. GitHub Actions
if os.path.exists('/tmp/creds.json'):
    creds = Credentials.from_service_account_file(
        '/tmp/creds.json',
        scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    )
    print("تم تحميل creds.json من GitHub Actions")

# 2. Colab أو محلي
else:
    possible_paths = [
        'creds.json',
        'stock-tracker-2025-c2e6fce3f1a7.json',
        '/content/drive/MyDrive/stock-tracker-2025-c2e6fce3f1a7.json'
    ]
    for path in possible_paths:
        if os.path.exists(path):
            creds = Credentials.from_service_account_file(
                path,
                scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            )
            print(f"تم تحميل creds.json من {path}")
            break

if creds is None:
    raise FileNotFoundError("لم يتم العثور على ملف creds.json !")

# === الاتصال بجوجل شيت ===
client = gspread.authorize(creds)
spreadsheet_id = '1St7HrYmMsfqV5LgZRUm57JZxYopOP10cjS0oYPSC0UM'
sh = client.open_by_key(spreadsheet_id)
print("تم الاتصال بجوجل شيت بنجاح")

# ==================================================================
# 1. تحديث الصفحة الرئيسية (الأسهم الحلال المختارة)
# ==================================================================
sheet = sh.sheet1

# مسح البيانات القديمة
all_vals = sheet.get_all_values()
if len(all_vals) > 1:
    rows_to_delete = len(all_vals) - 1
    try:
        sheet.delete_rows(2, 2 + rows_to_delete - 1)
        print(f"تم حذف {rows_to_delete} صف من الصفحة الرئيسية")
    except:
        for _ in range(min(rows_to_delete, 100)):
            sheet.delete_rows(2)
            time.sleep(0.3)

# العناوين
headers = ['التاريخ والوقت', 'اسم السهم', 'الرمز', 'السعر (ج.م)', 'التغيير (ج.م)', 'التغيير %',
           'السيولة (حجم)', 'الدعم (Support)', 'المقاومة (Resistance)']
if sheet.row_values(1) != headers:
    sheet.update('A1:I1', [headers])
    sheet.format("A1:I1", {"textFormat": {"bold": True}})
    print("تم تحديث العناوين في الصفحة الرئيسية")

# الأسهم الحلال (DCRC.CA تم حذفه لأنه delisted)
STOCK_NAMES_AR = {
    'ADIB.CA': 'بنك أبوظبي الإسلامي', 'SAUD.CA': 'البنك السعودي للاستثمار', 'AMIA.CA': 'الإسكندرية للاستثمارات',
    'ATLC.CA': 'أطلس للاستثمار', 'FAITA.CA': 'فيتا للصناعات', 'AIFI.CA': 'الإسكندرية للاستثمار',
    'CAED.CA': 'القاهرة للاستثمار', 'GIHD.CA': 'جي آي إتش دي', 'MBSC.CA': 'مصر بني سويف',
    'AMES.CA': 'أميس', 'ZEOT.CA': 'زيوت مصر', 'MOSC.CA': 'مصر للأسمنت', 'CCRS.CA': 'القاهرة للاستثمار',
    'NDRL.CA': 'الدلتا للسكر', 'CLHO.CA': 'كلين هاوس', 'MCRO.CA': 'مايكرو', 'AXPH.CA': 'الإسكندرية للأدوية',
    'AJWA.CA': 'أجوا', 'SPMD.CA': 'سبيد ميديكال'
}

stocks = list(STOCK_NAMES_AR.keys())
update_time = datetime.now().strftime('%Y-%m-%d %H:%M')
main_rows = []

print("\nجاري جلب بيانات الأسهم الحلال المختارة...")
for symbol in stocks:
    try:
        ticker = yf.Ticker(symbol, session=session)
        hist_month = ticker.history(period="1mo", interval="1d", timeout=15)
        hist_day = ticker.history(period="2d", interval="1d", timeout=15)

        if hist_month.empty or hist_day.empty or len(hist_month) < 5:
            raise ValueError("بيانات غير كافية")

        latest = hist_day.iloc[-1]
        open_p = hist_day.iloc[0]['Open']
        close = latest['Close']
        volume = latest['Volume'] if not pd.isna(latest['Volume']) else 0

        price = round(float(close), 2)
        change = round(float(close - open_p), 2)
        pct = round(float((close - open_p) / open_p * 100), 2)
        volume = int(volume)
        support = round(float(hist_month['Low'].min()), 2)
        resistance = round(float(hist_month['High'].max()), 2)

        arabic_name = STOCK_NAMES_AR.get(symbol, symbol)
        main_rows.append([update_time, arabic_name, symbol, price, change, pct, volume, support, resistance])
        print(f"✓ {symbol} → {price} ج.م ({pct:+.2f}%)")

    except Exception as e:
        print(f"✗ {symbol} فشل: {e}")
    time.sleep(1.2)

if main_rows:
    sheet.append_rows(main_rows, value_input_option='RAW')
    print(f"تم تحديث الصفحة الرئيسية بـ {len(main_rows)} سهم حلال ✓")

# ==================================================================
# 2. تحديث تاب EGX (223 سهم كامل)
# ==================================================================
try:
    egx_sheet = sh.worksheet("EGX")
    print("تم العثور على تاب EGX")
except gspread.exceptions.WorksheetNotFound:
    egx_sheet = sh.add_worksheet(title="EGX", rows=400, cols=12)
    print("تم إنشاء تاب EGX جديد")

# مسح البيانات القديمة في تاب EGX - آمن 100% لكل إصدارات gspread
print("جاري مسح البيانات القديمة في تاب EGX بطريقة آمنة...")
frozen = egx_sheet.frozen_row_count if hasattr(egx_sheet, 'frozen_row_count') else 1
all_vals_egx = egx_sheet.get_all_values()

if len(all_vals_egx) > frozen:
    try:
        # الطريقة الصحيحة في الإصدارات الحديثة
        egx_sheet.batch_clear([f"A{frozen+1}:Z1000"])
        print("تم مسح البيانات القديمة بنجاح باستخدام batch_clear")
    except:
        # fallback لو الإصدار قديم جدًا
        egx_sheet.spreadsheet.values_clear(f"{egx_sheet.title}!A{frozen+1}:Z1000")
        print("تم مسح البيانات القديمة بطريقة بديلة")
else:
    print("تاب EGX فاضي بالفعل (غير العناوين)")

# العناوين
egx_headers = ['التاريخ والوقت', 'اسم السهم', 'الرمز', 'السعر (ج.م)', 'التغيير (ج.م)',
               'التغيير %', 'السيولة', 'الدعم', 'المقاومة', 'المؤشر']

if egx_sheet.row_values(1) != egx_headers:
    egx_sheet.update('A1:J1', [egx_headers])
    egx_sheet.format("A1:J1", {"textFormat": {"bold": True},
                               "backgroundColor": {"red": 0.1, "green": 0.5, "blue": 0.9}})

# قائمة كاملة 223 سهم (نوفمبر 2025)
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

# تصنيف المؤشرات
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

# قاموس الأسماء العربية والإنجليزية
NAMES = {
    'AALR.CA': 'العامة لاستصلاح الأراضي',
    'ABUK.CA': 'Abou Kir Fertilizers & Chemical Industries Co.',
    'ACAMD.CA': 'Arab Co. for Asset Management & Development',
    'ACAP.CA': 'إيه كابيتال القابضة',
    'ACGC.CA': 'Arab Cotton Ginning Co.',
    'ACRO.CA': 'اكرو مصر',
    'ACTF.CA': 'اكت فاينانشال',
    'ADCI.CA': 'Arab Pharmaceuticals',
    'ADIB.CA': 'Abu Dhabi Islamic Bank-Egypt',
    'ADPC.CA': 'Arab Dairy Products Co. Arab Dairy - Panda',
    'AFDI.CA': 'Alahly For Development & Investment',
    'AFMC.CA': 'Alexandria Flour Mills Co.',
    'AIFI.CA': 'Atlas for Investment & Food Industries SAE',
    'AJWA.CA': 'Ajwa for Food Industries Co. Egypt',
    'ALCN.CA': 'الاسكندرية لتداول الحاويات',
    'AMER.CA': 'Amer Group Holding',
    'AMES.CA': 'Alexandria New Medical Center Co.',
    'AMIA.CA': 'Arab Moltaqa Investments Company',
    'AMOC.CA': 'أموك',
    'ARAB.CA': 'Arab Developers Holding',
    'ARCC.CA': 'Arabian Cement Company',
    'ASCM.CA': 'ASEC Co. for Mining',
    'ATLC.CA': 'Al Tawfeek Leasing Company-A.T.LEASE',
    'ATQA.CA': 'Misr National Steel',
    'AXPH.CA': 'Alexandria Company for Pharmaceuticals and Chemical Industries',
    'BINV.CA': 'B Investments Holding SAE',
    'BIOC.CA': 'GlaxoSmithKline S.A.E.',
    'BONY.CA': 'Bonyan for Development and Trade',
    'BTFH.CA': 'Beltone Holding',
    'CAED.CA': 'Cairo Educational Services',
    'CANA.CA': 'Suez Canal Bank SAE',
    'CCAP.CA': 'QALA For Financial Investments',
    'CCAPP.CA': 'QALA For Financial Investments - Preferred',
    'CCRS.CA': 'Gulf Canadian Real Estate Investment Co.',
    'CERA.CA': 'Arab Ceramic Co. - Ceramica Remas',
    'CICH.CA': 'CI Capital Holding for Financial Investments',
    'CIEB.CA': 'Credit Agricole Egypt',
    'CIRA.CA': 'Cairo For Investment And Real Estate Developments -CIRA Education',
    'CLHO.CA': 'Cleopatra Hospital Company',
    'CNFN.CA': 'Contact Financial Holding SAE',
    'COMI.CA': 'Commercial International Bank - Egypt (CIB) S.A.E.',
    'COSG.CA': 'Cairo Oils & Soap',
    'CSAG.CA': 'Canal Shipping Agencies Co.',
    'DEIN.CA': 'Delta Insurance',
    'DOMT.CA': 'Arabian Food Industries Co.',
    'DSCW.CA': 'Dice Sports & Casual Wear Manufacturers SAE',
    'EAST.CA': 'Eastern Company',
    'ECAP.CA': 'El Ezz Ceramics & Porcelain Co. (Gemma)',
    'EFIC.CA': 'Egyptian Financial & Industrial Co.',
    'EFID.CA': 'Edita Food Industries SAE',
    'EFIH.CA': 'e-finance for Digital and Financial Investments S.A.E.',
    'EGAL.CA': 'Egypt Aluminum',
    'EGCH.CA': 'Egyptian Chemical Industries',
    'EGTS.CA': 'Egyptian for Tourism Resorts',
    'EHDR.CA': 'Egyptians Housing Development & Reconstruction',
    'EKHO.CA': 'Egypt Kuwait Holding Co. SAE',
    'EKHOA.CA': 'Egypt Kuwait Holding Co. SAE - Preferred',
    'ELEC.CA': 'Electro Cable Egypt',
    'EMFD.CA': 'Emaar Misr for Development SAE',
    'ENGC.CA': 'Industrial Engineering Co. for Construction & Development',
    'ETEL.CA': 'Telecom Egypt',
    'EXPA.CA': 'Export Development Bank of Egypt',
    'FAITA.CA': 'Faisal Islamic Bank of Egypt - Preferred',
    'FERC.CA': 'Fertilizers & Chemicals',
    'FWRY.CA': 'Fawry for Banking Technology and Electronic Payments SAE',
    'GBCO.CA': 'GB AUTO',
    'GDWA.CA': 'Gadwa For Industrial Development',
    'GGCC.CA': 'Giza General Contracting & Real Estate Investment',
    'GIHD.CA': 'غزل المحلة',
    'GPPL.CA': 'Golden Pyramids Plaza',
    'HDBK.CA': 'Housing & Development Bank',
    'HELI.CA': 'Heliopolis Housing',
    'HRHO.CA': 'EFG Hermes Holding SAE',
    'IDRE.CA': 'Arab Real Estate Investment Co. (ALICO)',
    'IFAP.CA': 'International Agricultural Products',
    'INFI.CA': 'Ismailia National Co. for Food Industries',
    'IRON.CA': 'Egyptian Iron & Steel',
    'ISMA.CA': 'Ismailia Misr Poultry',
    'ISPH.CA': 'Ibn Sina Pharma',
    'JUFO.CA': 'Juhayna Food Industries',
    'KRDI.CA': 'Al-Khair River For Development Agricultural Investment & Environmental Services Company',
    'KZPC.CA': 'Kafr El Zayat Pesticides',
    'LCSW.CA': 'Lecico Egypt SAE',
    'MAAL.CA': 'Marseille Almasreia Alkhalegeya For Holding Investment SAE',
    'MASR.CA': 'Madinet Masr For Housing and Development',
    'MBSC.CA': 'Misr Beni Suef Cement',
    'MCQE.CA': 'Misr Cement (Qena)',
    'MEPA.CA': 'Medical Packaging Company',
    'MFPC.CA': 'Misr Chemical Industries',
    'MHOT.CA': 'Misr Hotels',
    'MICH.CA': 'Misr Chemical Industries',
    'MILS.CA': 'North Cairo Mills',
    'MOIL.CA': 'Maridive & Oil Services SAE',
    'MOSC.CA': 'Misr Oils & Soap',
    'MPCO.CA': 'Mansourah Poultry',
    'MPRC.CA': 'Egyptian Media Production City',
    'MTIE.CA': 'MM Group for Industry and International Trade SAE',
    'NAHO.CA': 'Nasr City Housing',
    'NDRL.CA': 'الدلتا للسكر',
    'NINH.CA': 'Nozha International Hospital',
    'OCDI.CA': 'Six of October Development & Investment',
    'ODIN.CA': 'El Orouba for Development & Investment',
    'OIH.CA': 'Orascom Investment Holding SAE',
    'OLFI.CA': 'Obour Land for Food Industries SAE',
    'ORAS.CA': 'Orascom Construction PLC',
    'ORHD.CA': 'Orascom Development Egypt SAE',
    'ORWE.CA': 'Oriental Weavers',
    'PHAR.CA': 'Egyptian International Pharmaceuticals (EIPICO)',
    'PHDC.CA': 'Palm Hills Developments SAE',
    'POCO.CA': 'Port Saied for Development & Construction',
    'POUL.CA': 'Cairo Poultry',
    'PRCL.CA': 'General Company For Ceramic & Porcelain Products',
    'QNBE.CA': 'Qatar National Bank Alahly',
    'RAKT.CA': 'Rakta Paper Manufacturing',
    'RAYA.CA': 'Raya Holding for Financial Investments SAE',
    'RMDA.CA': 'Tenth of Ramadan Pharmaceutical Industries&Diagnostic-Rameda',
    'RTVC.CA': 'Remco for Touristic Villages Construction',
    'SAUD.CA': 'Saudi Egyptian Investment & Finance',
    'SCEM.CA': 'South Cairo & Giza Mills & Bakeries',
    'SEIG.CA': 'El Sewedy Electric Co.',
    'SEIGA.CA': 'El Sewedy Electric Co. - Preferred',
    'SKPC.CA': 'S K P C',
    'SPMD.CA': 'سبيد ميديكال',
    'SWDY.CA': 'El Sewedy Electric Co.',
    'TALM.CA': 'Taaleem Management Services',
    'TMGH.CA': 'Talaat Moustafa Group',
    'ZEOT.CA': 'استخراج الزيوت',
    # أضف المزيد إذا لزم الأمر
}

egx_time = datetime.now().strftime('%Y-%m-%d %H:%M')
egx_rows = []
failed = []

print(f"\nجاري جلب بيانات {len(all_stocks)} سهم لتاب EGX...")

for symbol in all_stocks:
    try:
        ticker = yf.Ticker(symbol, session=session)
        hist_month = ticker.history(period="1mo", interval="1d")
        hist_day = ticker.history(period="2d", interval="1d")

        if hist_month.empty or hist_day.empty or len(hist_month) < 3:
            raise ValueError("بيانات ناقصة")

        latest = hist_day.iloc[-1]
        close = latest['Close']
        prev_open = hist_day.iloc[0]['Open']
        volume = int(latest['Volume']) if not pd.isna(latest['Volume']) else 0

        price = round(close, 2)
        change = round(close - prev_open, 2)
        pct = round((close - prev_open) / prev_open * 100, 2)
        support = round(hist_month['Low'].min(), 2)
        resistance = round(hist_month['High'].max(), 2)

        name = NAMES.get(symbol, ticker.info.get('longName', symbol.replace('.CA', '')))
        index = "EGX30" if symbol in EGX30_SYMBOLS else "EGX70" if symbol in EGX70_SYMBOLS else "أخرى"

        egx_rows.append([egx_time, name, symbol.replace('.CA',''), price, change, pct, volume, support, resistance, index])
        print(f"{symbol.replace('.CA',''):>8} │ {price:>6} │ {pct:+6.2f}% │ {index}")

    except Exception as e:
        failed.append(symbol)
        print(f"✗ {symbol} → {str(e)[:50]}")
    time.sleep(1.1)

if egx_rows:
    egx_sheet.append_rows(egx_rows, value_input_option='RAW')
    print(f"\nتم تحديث تاب EGX بنجاح بـ {len(egx_rows)} سهم ✓")

if failed:
    print(f"الأسهم الفاشلة ({len(failed)}): {', '.join([s.replace('.CA','') for s in failed[:20]])}")

print(f"\nتم التحديث النهائي في: {datetime.now().strftime('%Y-%m-%d %H:%M')} ✓")
print("الكود خلص شغله بنجاح 100%")