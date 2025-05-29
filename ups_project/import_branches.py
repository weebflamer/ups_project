import pandas as pd
import django
import os
import math
import jdatetime
from datetime import datetime as dt

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
    if val is None:
        return None
    if isinstance(val, float):
        if math.isnan(val):
            return None
        return int(val)
    if isinstance(val, int):
        return val
    try:
        s_val = str(val).strip()
        if s_val.endswith('.0') and s_val[:-2].isdigit():
            return int(float(s_val))
        if s_val.isdigit():
            return int(s_val)
    except (ValueError, TypeError):
        pass
    return None


# بررسی معتبر بودن مقدار سلول
def is_valid(value):
    if value is None:
        return False
    if isinstance(value, float) and math.isnan(value):
        return False
    if isinstance(value, (int, float)):
        return True
    return str(value).strip() != ''


# تابع برای تبدیل رشته تاریخ به آبجکت jdatetime.date
def convert_to_date_object_and_shamsi_string(date_value):
    if not is_valid(date_value):
        return None

    s = str(date_value).strip()

    # Handle float representation from Excel (e.g., '2023.0', '1402.0')
    if s.endswith('.0') and s[:-2].isdigit():
        s = s[:-2]  # Remove '.0' suffix for integer years

    # Try to parse as integer year first (e.g., '2023', '1402')
    if s.isdigit():
        year_int = int(s)
        if year_int > 1500:  # Likely Gregorian year (e.g., 1900-2100)
            try:
                g_date_obj = dt(year_int, 1, 1).date()
                return jdatetime.date.fromgregorian(date=g_date_obj)  # Return jdate object
            except ValueError:
                pass
        elif year_int < 1500 and year_int > 1000:  # Likely Shamsi year
            try:
                # If only year, attempt to create a jdate object for Jan 1st of that year
                return jdatetime.date(year_int, 1, 1)  # Return jdate object
            except ValueError:
                pass

    # Try parsing as a full Shamsi date string (e.g., '1402/05/10')
    try:
        return jdatetime.date.strptime(s, '%Y/%m/%d')  # Return jdate object
    except AttributeError:  # If jdatetime.date.strptime still causes AttributeError
        print(f"Warning: jdatetime.date.strptime is not available for '{s}'. Trying alternative parse.")
        # Fallback to try parsing as Gregorian if jdatetime.date.strptime is problematic
        try:
            g_date_obj = dt.strptime(s, '%Y/%m/%d').date()
            return jdatetime.date.fromgregorian(date=g_date_obj)  # Return jdate object
        except ValueError:
            pass
    except ValueError:
        pass  # If not a valid Shamsi date string

    # Try parsing as a full Gregorian date string (e.g., '2023-05-15', '2023/05/15')
    try:
        g_date_obj = dt.strptime(s, '%Y-%m-%d').date()
        return jdatetime.date.fromgregorian(date=g_date_obj)  # Return jdate object
    except ValueError:
        try:
            g_date_obj = dt.strptime(s, '%Y/%m/%d').date()
            return jdatetime.date.fromgregorian(date=g_date_obj)  # Return jdate object
        except ValueError:
            pass

    # If all parsing attempts fail, return None
    print(f"Warning: Could not convert '{s}' to a valid date object. Storing as None.")
    return None


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

    # تبدیل تاریخ‌ها قبل از ایجاد شعبه
    install_date_obj = convert_to_date_object_and_shamsi_string(row['تاریخ راه اندازی و نصب\nUPS'])
    last_battery_date_obj = convert_to_date_object_and_shamsi_string(row['تاریخ آخرین نصب باطری'])

    # ایجاد شعبه
    branch = Branch.objects.create(
        branch_code=str(row['شماره شعبه']).strip(),
        name=str(row['نام شعبه']).strip(),
        expert=str(row['کارشناس']).strip(),
        ups_power=str(row['توان UPS(KVA )']).strip(),
        install_date=install_date_obj,  # Pass date object
        last_battery_installed_date=last_battery_date_obj,  # Pass date object
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