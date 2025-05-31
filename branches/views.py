# branches/views.py
import os
from django.shortcuts import render, get_object_or_404, redirect
from django.db.models import Q
from django.http import HttpResponse
from django.contrib import messages
from django.conf import settings
from django.http import FileResponse

from .forms import BranchForm
import pandas as pd
import math
from .forms import UploadFileForm
from jalali_date import datetime2jalali, date2jalali
import jdatetime
from datetime import datetime as dt
import openpyxl
from .models import Branch, UPSBrand, BatteryModel

EXPECTED_COLUMNS = [
    'شماره شعبه', 'نام شعبه', 'درجه شعبه', 'تلفن', 'مدل ups', 'محل استفاده',
    'توان UPS(KVA )', 'تاریخ راه اندازی و نصب UPS', 'تعداد باطری نصب شده',
    'مدل و برند باطری', 'میزان آمپر هر باطری', 'ولتاژ باطری',
    'تاریخ تولید باطری (فقط شمسی)', 'تاریخ تولید باطری',
    'تاریخ آخرین نصب باطری', 'برق خروجی', 'وضعیت نصب باطری',
    'مدت زمان شارژ دهی به دقیقه', 'شماره سریال UPS', 'میزان ارت شعبه',
    'شماره تماس کارشناس شعبه', 'ملکی - استیجاری', 'کد پستی', 'آدرس', 'کارشناس'
]

REQUIRED_COLUMNS_FOR_DATA_ENTRY = [
    'شماره شعبه',
    'نام شعبه',
    'درجه شعبه',
    'تلفن',
    'شماره تماس کارشناس شعبه',
    'کارشناس'
]


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


def is_valid(value):
    if value is None:
        return False
    if isinstance(value, float) and math.isnan(value):
        return False
    if isinstance(value, (int, float)):
        return True
    return str(value).strip() != ''


def convert_to_date_object(date_value):
    if not is_valid(date_value):
        return None

    s = str(date_value).strip()

    if s.endswith('.0') and s[:-2].isdigit():
        s = s[:-2]

    if s.isdigit():
        year_int = int(s)
        if year_int > 1500:
            try:
                g_date_obj = dt(year_int, 1, 1).date()
                return jdatetime.date.fromgregorian(date=g_date_obj)
            except ValueError:
                pass
        elif year_int < 1500 and year_int > 1000:
            try:
                return jdatetime.date(year_int, 1, 1)
            except ValueError:
                pass

    try:
        return jdatetime.date.strptime(s, '%Y/%m/%d')
    except AttributeError:
        try:
            g_date_obj = dt.strptime(s, '%Y/%m/%d').date()
            return jdatetime.date.fromgregorian(date=g_date_obj)
        except ValueError:
            pass
    except ValueError:
        pass

    try:
        g_date_obj = dt.strptime(s, '%Y-%m-%d').date()
        return jdatetime.date.fromgregorian(date=g_date_obj)
    except ValueError:
        try:
            g_date_obj = dt.strptime(s, '%Y/%m/%d').date()
            return jdatetime.date.fromgregorian(date=g_date_obj)
        except ValueError:
            pass

    print(f"Warning: Could not convert '{s}' to a valid date object. Storing as None.")
    return None


def delete_branch(request, pk):
    branch = get_object_or_404(Branch, pk=pk)
    if request.method == 'POST':
        branch.delete()
        messages.success(request, f'شعبه {branch.name} با موفقیت حذف شد.')
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
            branch = form.save()
            messages.success(request, f'شعبه {branch.name} با موفقیت اضافه شد.')
            return redirect('branch_detail', pk=branch.pk)
    else:
        form = BranchForm()
    return render(request, 'branches/add_branch.html', {'form': form})


def edit_branch(request, pk):
    branch = get_object_or_404(Branch, pk=pk)
    if request.method == 'POST':
        form = BranchForm(request.POST, instance=branch)
        if form.is_valid():
            form.save()
            messages.success(request, f'شعبه {branch.name} با موفقیت ویرایش شد.')
            return redirect('branch_detail', pk=branch.pk)
    else:
        form = BranchForm(instance=branch)
    return render(request, 'branches/edit_branch.html', {'form': form, 'branch': branch})


def upload_excel(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            uploaded_file = request.FILES['file']
            file_extension = os.path.splitext(uploaded_file.name)[1].lower()

            df = None
            try:
                if file_extension == '.xlsx' or file_extension == '.xls':
                    df = pd.read_excel(uploaded_file)
                elif file_extension == '.csv':
                    encodings = ['utf-8-sig', 'utf-8', 'cp1256']
                    delimiters = [',', ';', '\t']

                    read_successful = False
                    for enc in encodings:
                        for delim in delimiters:
                            try:
                                uploaded_file.seek(0)
                                df = pd.read_csv(uploaded_file, encoding=enc, sep=delim)
                                df.columns = df.columns.str.strip()
                                if any(col in df.columns for col in EXPECTED_COLUMNS):
                                    read_successful = True
                                    break
                            except Exception as e:
                                pass
                        if read_successful:
                            break

                    if not read_successful:
                        messages.error(request,
                                       'فایل CSV قابل خواندن نیست یا فرمت آن پشتیبانی نمی‌شود. لطفاً انکدینگ یا جداکننده را بررسی کنید.')
                        return render(request, 'branches/upload_excel.html', {'form': form})

                else:
                    messages.error(request,
                                   'فرمت فایل پشتیبانی نمی‌شود. لطفاً فایل اکسل (.xlsx, .xls) یا CSV (.csv) بارگذاری کنید.')
                    return render(request, 'branches/upload_excel.html', {'form': form})
            except Exception as e:
                messages.error(request,
                               f'خطایی در پردازش فایل رخ داد: {e}. لطفاً از فرمت فایل و سلامت آن اطمینان حاصل کنید.')
                return render(request, 'branches/upload_excel.html', {'form': form})

            df.columns = df.columns.str.strip()

            df.dropna(how='all', inplace=True)

            missing_columns_in_file = [col for col in EXPECTED_COLUMNS if col not in df.columns]
            if missing_columns_in_file:
                messages.error(request,
                               f'ساختار فایل با قالب مورد انتظار مطابقت ندارد. ستون‌های زیر در فایل وجود ندارند: {", ".join(missing_columns_in_file)}')
                return render(request, 'branches/upload_excel.html',
                              {'form': form, 'missing_columns': missing_columns_in_file})

            errors = []
            new_branches_count = 0

            for index, row in df.iterrows():
                row_errors = []

                excel_row_number = index + 2

                # --- START: Modified branch_code processing ---
                branch_code_raw = row.get("شماره شعبه", None)
                branch_code = None
                if pd.isna(branch_code_raw):
                    row_errors.append(f'ستون "شماره شعبه" نمی‌تواند خالی باشد (مقدار NaN).')
                else:
                    try:
                        # Try to convert to int if it's a float like 3700000.0
                        if isinstance(branch_code_raw, float) and branch_code_raw == int(branch_code_raw):
                            branch_code = str(int(branch_code_raw)).strip()
                        else:
                            branch_code = str(branch_code_raw).strip()

                        if not branch_code:  # Check after stripping
                            row_errors.append(f'ستون "شماره شعبه" نمی‌تواند خالی باشد.')
                    except ValueError:
                        row_errors.append(
                            f'ستون "شماره شعبه" ({branch_code_raw}) مقدار نامعتبر دارد و نمی‌تواند به عدد تبدیل شود.')
                # --- END: Modified branch_code processing ---

                branch_name_raw = row.get("نام شعبه", None)
                if pd.isna(branch_name_raw):
                    row_errors.append(f'ستون "نام شعبه" نمی‌تواند خالی باشد (مقدار NaN).')
                else:
                    branch_name = str(branch_name_raw).strip()
                    if not branch_name:
                        row_errors.append(f'ستون "نام شعبه" نمی‌تواند خالی باشد.')

                for col_name in REQUIRED_COLUMNS_FOR_DATA_ENTRY:
                    if col_name in ["شماره شعبه", "نام شعبه"]:
                        continue

                    cell_value = row.get(col_name, None)
                    if pd.isna(cell_value) or not str(cell_value).strip():
                        row_errors.append(f'ستون "{col_name}" نمی‌تواند خالی باشد.')

                if row_errors:
                    errors.append(f'ردیف {excel_row_number} (خطا در فیلدهای اجباری): ' + '; '.join(row_errors))
                    continue

                try:
                    # branch_code and branch_name are guaranteed not to be empty and correctly formatted if we reached here
                    # Use the 'branch_code' variable that was processed above
                    # branch_name is also processed above

                    if Branch.objects.filter(branch_code=branch_code).exists():
                        errors.append(
                            f'ردیف {excel_row_number}: شعبه با کد {branch_code} ("{branch_name}") قبلاً وجود دارد و اضافه نشد.')
                        continue

                    grade = str(row.get("درجه شعبه", "")).strip() or None
                    phone_val = str(row.get("تلفن", "")).strip() or None
                    expert_val = str(row.get("کارشناس", "")).strip() or None

                    ups_power = str(row.get("توان UPS(KVA )", "")).strip() or None
                    install_date_obj = convert_to_date_object(row.get('تاریخ راه اندازی و نصب UPS', None))
                    battery_count = safe_int(row.get("تعداد باطری نصب شده", None))
                    battery_amp = str(row.get("میزان آمپر هر باطری", "")).strip() or None
                    battery_voltage = str(row.get("ولتاژ باطری", "")).strip() or None
                    last_battery_date_obj = convert_to_date_object(row.get('تاریخ آخرین نصب باطری', None))
                    ups_serial = str(row.get("شماره سریال UPS", "")).strip() or None
                    charge_duration = str(row.get("مدت زمان شارژ دهی به دقیقه", "")).strip() or None
                    address = str(row.get("آدرس", "")).strip() or None
                    battery_status = str(row.get("وضعیت نصب باطری", "")).strip() or None

                    branch = Branch.objects.create(
                        branch_code=branch_code,  # Use the cleaned branch_code
                        name=branch_name,  # Use the cleaned branch_name
                        grade=grade,
                        phone=phone_val,
                        expert=expert_val,

                        ups_power=ups_power,
                        install_date=install_date_obj,
                        last_battery_installed_date=last_battery_date_obj,
                        battery_count=battery_count,
                        battery_amp=battery_amp,
                        battery_voltage=battery_voltage,
                        ups_serial=ups_serial,
                        charge_duration=charge_duration,
                        address=address,
                        battery_status=battery_status,
                    )

                    ups_brand_names = str(row.get("مدل ups", "")).strip()
                    if ups_brand_names:
                        for brand_name in ups_brand_names.split(','):
                            brand_name = brand_name.strip()
                            if brand_name:
                                ups_brand_obj, created = UPSBrand.objects.get_or_create(name=brand_name)
                                branch.ups_brand.add(ups_brand_obj)

                    battery_model_names = str(row.get("مدل و برند باطری", "")).strip()
                    if battery_model_names:
                        for model_name in battery_model_names.split(','):
                            model_name = model_name.strip()
                            if model_name:
                                battery_model_obj, created = BatteryModel.objects.get_or_create(name=model_name)
                                branch.battery_model.add(battery_model_obj)

                    new_branches_count += 1

                except Exception as e:
                    errors.append(f'ردیف {excel_row_number}: خطایی در ذخیره رکورد رخ داد: {e}')

            if not errors:
                messages.success(request, f'{new_branches_count} شعبه جدید با موفقیت از فایل اکسل اضافه شد.')
                return redirect('branch_list')
            else:
                messages.warning(request,
                                 f'{new_branches_count} شعبه جدید اضافه شد، اما {len(errors)} خطا در پردازش فایل اکسل وجود داشت.')
                return render(request, 'branches/upload_excel.html', {'form': form, 'errors': errors})

        else:  # If form is NOT valid for POST request (e.g., no file selected)
            return render(request, 'branches/upload_excel.html', {'form': form})

    else:  # GET request
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
        'توان UPS', 'مدت شارژدهی', 'سریال UPS', 'آمپر باتری', 'ولتاژ باتری',
        'وضعیت باتری'
    ])

    for branch in branches:
        battery_models = ", ".join([str(bm) for bm in branch.battery_model.all()])
        ups_brands = ", ".join([str(ub) for ub in branch.ups_brand.all()])

        # --- START: Modified date conversion for export ---
        # تبدیل install_date (jmodels.jDateField) به رشته شمسی
        install_date_str = None
        if branch.install_date:
            # Convert jdate object to string in 'YYYY/MM/DD' format
            # branch.install_date is already a jdatetime.date object because of jmodels.jDateField
            install_date_str = branch.install_date.strftime('%Y/%m/%d')

        # تبدیل last_battery_installed_date (jmodels.jDateField) به رشته شمسی
        last_battery_installed_date_str = None
        if branch.last_battery_installed_date:
            # Convert jdate object to string in 'YYYY/MM/DD' format
            last_battery_installed_date_str = branch.last_battery_installed_date.strftime('%Y/%m/%d')
        # --- END: Modified date conversion for export ---

        ws.append([
            branch.name,
            branch.branch_code,
            branch.expert,
            branch.battery_count,
            battery_models,
            ups_brands,
            branch.address,
            branch.phone,
            install_date_str,  # استفاده از رشته شمسی تبدیل شده
            last_battery_installed_date_str,  # استفاده از رشته شمسی تبدیل شده
            branch.ups_power,
            branch.charge_duration,
            branch.ups_serial,
            branch.battery_amp,
            branch.battery_voltage,
            branch.battery_status
        ])

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=branches.xlsx'
    wb.save(response)
    return response
