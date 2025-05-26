import pandas as pd
import django
import os
import math

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'ups_project.settings')
django.setup()

from branches.models import Branch, UPSBrand, BatteryModel

# خواندن فایل اکسل
df = pd.read_excel('ups2.xlsx')

# حذف داده‌های قبلی
Branch.objects.all().delete()
UPSBrand.objects.all().delete()
BatteryModel.objects.all().delete()

# تبدیل امن به عدد صحیح
def safe_int(val):
    try:
        return int(val)
    except:
        return None

# بررسی معتبر بودن مقدار سلول
def is_valid(value):
    if value is None:
        return False
    if isinstance(value, float) and math.isnan(value):
        return False
    return str(value).strip() != ''

# حلقه پردازش سطرهای فایل اکسل
for index, row in df.iterrows():
    # برند UPS
    ups_brand_name = str(row['مدل ups']).strip() if is_valid(row['مدل ups']) else ''
    ups_brand_obj = None
    if ups_brand_name:
        ups_brand_obj, _ = UPSBrand.objects.get_or_create(name=ups_brand_name)

    # مدل باتری
    battery_model_name = str(row['مدل  و برند \nباطری']).strip() if is_valid(row['مدل  و برند \nباطری']) else ''
    battery_models = []
    if battery_model_name:
        bm_obj, _ = BatteryModel.objects.get_or_create(name=battery_model_name)
        battery_models.append(bm_obj)

    # ایجاد شعبه بدون اضافه کردن فیلدهای ManyToMany
    branch = Branch.objects.create(
        branch_code=str(row['شماره شعبه']).strip(),
        name=str(row['نام شعبه']).strip(),
        expert=str(row['کارشناس']).strip(),
        ups_power=str(row['توان UPS(KVA )']).strip(),
        install_date=str(row['تاریخ راه اندازی و نصب\nUPS']).strip(),
        last_battery_installed_date=str(row['تاریخ آخرین نصب باطری']).strip(),
        battery_count=safe_int(row['تعداد باطری\nنصب شده']),
        battery_amp=str(row['میزان آمپر هر باطری']).strip(),
        battery_voltage=str(row['ولتاژ باطری']).strip(),
        ups_serial=str(row['شماره سریال UPS']).strip(),
        charge_duration=str(row['مدت زمان \nشارژ دهی به دقیقه']).strip(),
        address=str(row['آدرس']).strip(),
        grade=str(row['درجه شعبه']).strip(),
        phone=str(row['تلفن']).strip(),
        battery_status=str(row["وضعیت نصب باطری"]).strip(),
    )

    # افزودن رابطه‌های ManyToMany
    if ups_brand_obj:
        branch.ups_brand.add(ups_brand_obj)

    for bm in battery_models:
        branch.battery_model.add(bm)
