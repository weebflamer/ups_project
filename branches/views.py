# branches/views.py
import os
from django.shortcuts import render, get_object_or_404, redirect  # redirect را اضافه کنید
from django.db.models import Q
from django.http import HttpResponse

from django.conf import settings
from django.http import FileResponse

from .forms import BranchForm

import pandas as pd

from .forms import UploadFileForm

from jalali_date import datetime2jalali, date2jalali  # Keep these imports
import jdatetime  # Add this import
from datetime import datetime as dt  # Add this import

import openpyxl
from .models import Branch

EXPECTED_COLUMNS = [
    'شماره شعبه', 'نام شعبه', 'درجه شعبه', 'تلفن', 'مدل ups', 'محل استفاده',
    'توان UPS(KVA )', 'تاریخ راه اندازی و نصب UPS', 'تعداد باطری نصب شده',
    'مدل و برند باطری', 'میزان آمپر هر باطری', 'ولتاژ باطری',
    'تاریخ تولید باطری (فقط شمسی)', 'تاریخ تولید باطری',
    'تاریخ آخرین نصب باطری', 'برق خروجی', 'وضعیت نصب باطری',
    'مدت زمان شارژ دهی به دقیقه', 'شماره سریال UPS', 'میزان ارت شعبه',
    'شماره تماس کارشناس شعبه', 'ملکی - استیجاری', 'کد پستی', 'آدرس', 'کارشناس'
]


# تابع برای تبدیل امن به عدد صحیح (از import_branches.py کپی شده)
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


# تابع برای بررسی معتبر بودن مقدار سلول (از import_branches.py کپی شده)
def is_valid(value):
    if value is None:
        return False
    if isinstance(value, float) and math.isnan(value):
        return False
    if isinstance(value, (int, float)):
        return True
    return str(value).strip() != ''


# تابع برای تبدیل رشته تاریخ به آبجکت jdatetime.date (از import_branches.py کپی شده)
def convert_to_date_object(date_value):
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
                return jdatetime.date(year_int, 1, 1)  # Return jdate object
            except ValueError:
                pass

    # Try parsing as a full Shamsi date string (e.g., '1402/05/10')
    try:
        return jdatetime.date.strptime(s, '%Y/%m/%d')  # Return jdate object
    except AttributeError:
        print(f"Warning: jdatetime.date.strptime is not available for '{s}'. Trying alternative parse.")
        try:
            g_date_obj = dt.strptime(s, '%Y/%m/%d').date()
            return jdatetime.date.fromgregorian(date=g_date_obj)  # Return jdate object
        except ValueError:
            pass
    except ValueError:
        pass

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


def delete_branch(request, pk):
    branch = get_object_or_404(Branch, pk=pk)
    if request.method == 'POST':
        branch.delete()
        return redirect('branch_list')
    return render(request, 'branches/confirm_delete.html', {'branch': branch})


def branch_detail(request, pk):
    branch = get_object_or_404(Branch, pk=pk)

    next_branch = Branch.objects.filter(pk__gt=pk).order_by('pk').first()
    prev_branch = Branch.objects.filter(pk__lt=pk).order_by('-pk').first()

    return render(request, 'branches/branch_detail.html', {
        'branch': branch,
        'next_branch': next_branch,
        'prev_branch': prev_branch
    })


def branch_list(request):
    branches = Branch.objects.all()
    return render(request, 'branches/branch_list.html', {'branches': branches})


def branch_list(request):
    query = request.GET.get('q')
    if query:
        branches = Branch.objects.filter(
            Q(name__icontains=query) |
            Q(branch_code__icontains=query)
        )
    else:
        branches = Branch.objects.all()
    return render(request, 'branches/branch_list.html', {'branches': branches})


def add_branch(request):
    if request.method == 'POST':
        form = BranchForm(request.POST)
        if form.is_valid():
            branch = form.save()  # شعبه را ذخیره کنید
            return redirect('branch_detail', pk=branch.pk)  # به جزئیات شعبه جدید هدایت کنید
    else:
        form = BranchForm()
    # این خط برای درخواست GET یا در صورت نامعتبر بودن فرم است
    return render(request, 'branches/add_branch.html', {'form': form})


def edit_branch(request, pk):
    branch = get_object_or_404(Branch, pk=pk)
    if request.method == 'POST':
        form = BranchForm(request.POST, instance=branch)
        if form.is_valid():
            form.save()
            return redirect('branch_detail', pk=branch.pk)
    else:
        form = BranchForm(instance=branch)
    return render(request, 'branches/edit_branch.html', {'form': form, 'branch': branch})


def upload_excel(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            try:
                df = pd.read_excel(request.FILES['file'])
            except Exception as e:
                return render(request, 'branches/upload_excel.html', {
                    'form': form,
                    'error': 'فایل قابل خواندن نیست. لطفاً یک فایل اکسل معتبر بارگذاری کنید.'
                })

            missing_columns = [col for col in EXPECTED_COLUMNS if col not in df.columns]
            if missing_columns:
                return render(request, 'branches/upload_excel.html', {
                    'form': form,
                    'error': f'ستون‌های زیر در فایل وجود ندارند: {", ".join(missing_columns)}'
                })

            for _, row in df.iterrows():
                install_date_obj = convert_to_date_object(row['تاریخ راه اندازی و نصب UPS'])
                last_battery_date_obj = convert_to_date_object(row['تاریخ آخرین نصب باطری'])

                Branch.objects.create(
                    branch_code=str(row["شماره شعبه"]),
                    name=row["نام شعبه"],
                    grade=row.get("درجه شعبه"),
                    phone=row.get("تلفن"),
                    ups_power=row.get("توان UPS(KVA )"),
                    install_date=install_date_obj,  # Pass date object
                    last_battery_installed_date=last_battery_date_obj,  # Pass date object
                    battery_count=safe_int(row.get("تعداد باطری نصب شده")),
                    # Assuming battery_model column is handled later as ManyToMany
                    # if not, it should be adjusted based on models.py
                    battery_amp=row.get("میزان آمپر هر باطری"),
                    battery_voltage=row.get("ولتاژ باطری"),
                    ups_serial=row.get("شماره سریال UPS"),
                    charge_duration=row.get("مدت زمان شارژ دهی به دقیقه"),
                    address=row.get("آدرس"),
                    expert=row.get("کارشناس")
                )
            return redirect('branch_list')
    else:
        form = UploadFileForm()
    return render(request, 'branches/upload_excel.html', {'form': form})


def download_template(request):
    file_path = os.path.join(settings.BASE_DIR, 'temp.xlsx')
    if os.path.exists(file_path):
        return FileResponse(open(file_path, 'rb'), as_attachment=True, filename='قالب_شعب.xlsx')
    else:
        return HttpResponse("فایل قالب پیدا نشد.", status=404)


def export_to_excel(request):
    branches = Branch.objects.all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Branches"

    # سرستون‌ها
    ws.append([
        'نام شعبه', 'کد شعبه', 'کارشناس', 'تعداد باتری', 'مدل باتری',
        'برند UPS', 'آدرس', 'تلفن', 'تاریخ نصب', 'آخرین تاریخ نصب باتری',
        'توان UPS', 'مدت شارژدهی', 'سریال UPS', 'آمپر باتری', 'ولتاژ باتری'
    ])

    for branch in branches:
        battery_models = ", ".join([str(bm) for bm in branch.battery_model.all()])
        ups_brands = ", ".join([str(ub) for ub in branch.ups_brand.all()])

        # تبدیل jDateField به رشته برای اکسل
        install_date_str = None
        if branch.install_date:
            install_date_str = date2jalali(branch.install_date).strftime('%Y/%m/%d')

        last_battery_installed_date_str = None
        if branch.last_battery_installed_date:
            last_battery_installed_date_str = date2jalali(branch.last_battery_installed_date).strftime('%Y/%m/%d')

        ws.append([
            branch.name,
            branch.branch_code,
            branch.expert,
            branch.battery_count,
            battery_models,
            ups_brands,
            branch.address,
            branch.phone,
            install_date_str,  # استفاده از رشته شمسی
            last_battery_installed_date_str,  # استفاده از رشته شمسی
            branch.ups_power,
            branch.charge_duration,
            branch.ups_serial,
            branch.battery_amp,
            branch.battery_voltage
        ])

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=branches.xlsx'
    wb.save(response)
    return response