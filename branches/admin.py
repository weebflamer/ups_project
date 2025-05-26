from django.contrib import admin
from .models import Branch, UPSBrand, BatteryModel

@admin.register(Branch)
class BranchAdmin(admin.ModelAdmin):
    filter_horizontal = ('battery_model', 'ups_brand')

admin.site.register(UPSBrand)
admin.site.register(BatteryModel)
