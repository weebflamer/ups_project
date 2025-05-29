# branches/forms.py
from django import forms
from .models import Branch, BatteryModel, UPSBrand
# دیگر از jalali_date.fields یا jalali_date.widgets چیزی وارد نمی‌کنیم.

class UploadFileForm(forms.Form):
    file = forms.FileField(label='آپلود فایل اکسل')

class BranchForm(forms.ModelForm):
    battery_model = forms.ModelMultipleChoiceField(
        queryset=BatteryModel.objects.all(),
        widget=forms.CheckboxSelectMultiple,
        required=False,
        label='مدل باتری'
    )

    ups_brand = forms.ModelMultipleChoiceField(
        queryset=UPSBrand.objects.all(),
        widget=forms.CheckboxSelectMultiple,
        required=False,
        label='برند UPS'
    )

    install_date = forms.CharField(
        widget=forms.TextInput(attrs={'class': 'form-control'}),
        required=False, # این خط را حفظ کنید
        label='تاریخ نصب'
    )

    last_battery_installed_date = forms.CharField(
        widget=forms.TextInput(attrs={'class': 'form-control'}),
        required=False, # این خط را حفظ کنید
        label='آخرین تاریخ نصب باتری'
    )

    class Meta:
        model = Branch
        fields = '__all__'
        labels = {
            'branch_code': 'کد شعبه',
            'name': 'نام شعبه',
            'grade': 'درجه',
            'phone': 'تلفن',
            'ups_power': 'توان UPS',
            'battery_count': 'تعداد باتری',
            'battery_model': 'مدل باتری',
            'battery_status': 'وضعیت باتری',
            'charge_duration': 'مدت شارژدهی (دقیقه)',
            'address': 'آدرس',
            'expert': 'کارشناس سداد',
            'ups_brand': 'برند UPS',
            'ups_serial': 'سریال UPS',
            'battery_amp': 'آمپر باتری',
            'battery_voltage': 'ولتاژ باتری',
        }
        widgets = {
            'charge_duration': forms.NumberInput(attrs={'min': 0}),
            'battery_status': forms.TextInput(attrs={'placeholder': 'مثلاً سالم / نیاز به تعویض'}),
            'phone': forms.TextInput(attrs={'type': 'tel'}),
            'address': forms.Textarea(attrs={'rows': 2}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        for name, field in self.fields.items():
            field.required = False # این خط باید اینجا باشد تا همه فیلدها را غیرضروری کند
            if not isinstance(field.widget, forms.CheckboxSelectMultiple):
                current_classes = field.widget.attrs.get('class', '').split()
                if 'form-control' not in current_classes:
                    current_classes.append('form-control')
                field.widget.attrs['class'] = ' '.join(current_classes)

    # متدهای clean_FIELDNAME برای مدیریت فیلدهای تاریخ خالی
    def clean_install_date(self):
        install_date = self.cleaned_data.get('install_date')
        if not install_date: # اگر فیلد خالی بود (رشته خالی)
            return None # آن را به None تبدیل کن
        return install_date # در غیر این صورت، مقدار موجود را برگردان

    def clean_last_battery_installed_date(self):
        last_battery_installed_date = self.cleaned_data.get('last_battery_installed_date')
        if not last_battery_installed_date: # اگر فیلد خالی بود (رشته خالی)
            return None # آن را به None تبدیل کن
        return last_battery_installed_date # در غیر این صورت، مقدار موجود را برگردان