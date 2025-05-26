from django.db import models

class UPSBrand(models.Model):
    name = models.CharField(max_length=100, unique=True)

    def __str__(self):
        return self.name

class BatteryModel(models.Model):
    name = models.CharField(max_length=100, unique=True)

    def __str__(self):
        return self.name

class Branch(models.Model):
    branch_code = models.CharField(max_length=20)
    name = models.CharField(max_length=100)
    expert = models.CharField(max_length=100, blank=True, null=True)
    ups_brand = models.ManyToManyField(UPSBrand, blank=True, verbose_name='برند UPS')
    ups_power = models.CharField(max_length=20, blank=True, null=True)
    install_date = models.CharField(max_length=100, blank=True, null=True)
    last_battery_installed_date = models.CharField(max_length=100, blank=True, null=True)
    battery_count = models.IntegerField(blank=True, null=True)
    battery_model = models.ManyToManyField(BatteryModel, blank=True, verbose_name='مدل باتری')
    battery_amp = models.CharField(max_length=20, blank=True, null=True)
    battery_voltage = models.CharField(max_length=20, blank=True, null=True)
    ups_serial = models.CharField(max_length=50, blank=True, null=True)
    charge_duration = models.CharField(max_length=20, blank=True, null=True)
    address = models.TextField(max_length=200,blank=True, null=True)
    grade = models.CharField(max_length=50, blank=True)
    phone = models.CharField(max_length=50, blank=True)
    battery_status = models.CharField(max_length=100, blank=True)

    def __str__(self):
        return f"{self.branch_code} - {self.name}"
