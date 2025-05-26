from django import forms
from .models import Branch, BatteryModel, UPSBrand

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
        widget=forms.TextInput(attrs={'class': 'form-control datepicker'}),
        required=False,
        label='تاریخ نصب'
    )

    last_battery_installed_date = forms.CharField(
        widget=forms.TextInput(attrs={'class': 'form-control datepicker'}),
        required=False,
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
            'install_date': 'تاریخ نصب',
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
            'last_battery_installed_date': 'آخرین تاریخ نصب باتری',
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
            field.required = False
            if not isinstance(field.widget, forms.CheckboxSelectMultiple):
                field.widget.attrs.update({'class': 'form-control'})
