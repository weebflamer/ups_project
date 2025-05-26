import os
from django.shortcuts import render, get_object_or_404, redirect
from django.db.models import Q
from django.http import HttpResponse

from django.conf import settings
from django.http import FileResponse

from .forms import BranchForm

import pandas as pd

from .forms import UploadFileForm

from jalali_date import datetime2jalali, date2jalali


import openpyxl


EXPECTED_COLUMNS = [
    'شماره شعبه', 'نام شعبه', 'درجه شعبه', 'تلفن', 'مدل ups', 'محل استفاده',
    'توان UPS(KVA )', 'تاریخ راه اندازی و نصب UPS', 'تعداد باطری نصب شده',
    'مدل و برند باطری', 'میزان آمپر هر باطری', 'ولتاژ باطری',
    'تاریخ تولید باطری (فقط شمسی)', 'تاریخ تولید باطری',
    'تاریخ آخرین نصب باطری', 'برق خروجی', 'وضعیت نصب باطری',
    'مدت زمان شارژ دهی به دقیقه', 'شماره سریال UPS', 'میزان ارت شعبه',
    'شماره تماس کارشناس شعبه', 'ملکی - استیجاری', 'کد پستی', 'آدرس', 'کارشناس'
]



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
            form.save()
            return render(request, 'add_branch.html', {'form': form, 'title': 'افزودن شعبه جدید'})

    else:
        form = BranchForm()
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
                Branch.objects.create(
                    branch_code=str(row["شماره شعبه"]),
                    name=row["نام شعبه"],
                    grade=row.get("درجه شعبه"),
                    phone=row.get("تلفن"),
                    ups_power=row.get("توان UPS(KVA )"),
                    install_date=row.get("تاریخ راه اندازی و نصب UPS"),
                    battery_count=row.get("تعداد باطری نصب شده"),
                    battery_model=row.get("مدل و برند باطری"),
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



import openpyxl
from django.http import HttpResponse
from .models import Branch


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

        ws.append([
            branch.name,
            branch.branch_code,
            branch.expert,
            branch.battery_count,
            battery_models,
            ups_brands,
            branch.address,
            branch.phone,
            branch.install_date,
            branch.last_battery_installed_date,
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


