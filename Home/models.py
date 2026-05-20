from decimal import Decimal
from django.db import models
from django.contrib.auth.models import AbstractUser
from django.core.validators import RegexValidator
from django.utils import timezone
from django.conf import settings
from Users.models import CustomUser

# =========================================================
# VALIDATORS
# =========================================================

kenyan_phone_validator = RegexValidator(
    regex=r'^2547\d{8}$',
    message="Phone number must be in format: 2547XXXXXXXX"
)



# =========================================================
# DEPARTMENTS
# =========================================================
from django.db import models

class Department(models.Model):
    # Changed to a standard CharField without fixed choices
    name = models.CharField(max_length=50, unique=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['name']

    def __str__(self):
        return self.name

# =========================================================
# EMPLOYEES
# =========================================================

class Employee(models.Model):
    department = models.ForeignKey(Department, on_delete=models.CASCADE, related_name='employees')
    first_name = models.CharField(max_length=100)
    last_name = models.CharField(max_length=100)
    id_number = models.CharField(max_length=20, unique=True)
    phone_number = models.CharField(max_length=12, validators=[kenyan_phone_validator])
    email = models.EmailField(blank=True, null=True)
    address = models.TextField(blank=True)
    basic_salary = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))
    hire_date = models.DateField()
    is_active = models.BooleanField(default=True)
    
    # Changed from User to settings.AUTH_USER_MODEL
    created_by = models.ForeignKey(CustomUser, on_delete=models.SET_NULL, null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['first_name', 'last_name']
        indexes = [
            models.Index(fields=['department']),
            models.Index(fields=['id_number']),
        ]

    def __str__(self):
        return f"{self.first_name} {self.last_name}"

# =========================================================
# ATTENDANCE & PAYSLIPS
# =========================================================

class Attendance(models.Model):
    employee = models.ForeignKey(Employee, on_delete=models.CASCADE, related_name='attendance_records')
    date = models.DateField()
    present = models.BooleanField(default=True)
    remarks = models.CharField(max_length=255, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        unique_together = ('employee', 'date')
        ordering = ['-date']

    def __str__(self):
        return f"{self.employee} - {self.date}"

class Payslip(models.Model):
    STATUS_CHOICES = [('pending', 'Pending'), ('paid', 'Paid')]
    employee = models.ForeignKey(Employee, on_delete=models.CASCADE, related_name='payslips')
    month_year = models.DateField()
    allowances = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))
    deductions = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))
    overtime_amount = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))
    bonus = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))
    status = models.CharField(max_length=10, choices=STATUS_CHOICES, default='pending')
    payment_date = models.DateField(blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        unique_together = ('employee', 'month_year')
        ordering = ['-month_year']

    @property
    def gross_pay(self):
        return self.employee.basic_salary + self.allowances + self.overtime_amount + self.bonus

    @property
    def net_pay(self):
        return self.gross_pay - self.deductions

    def __str__(self):
        return f"{self.employee} - {self.month_year.strftime('%B %Y')}"

# =========================================================
# ASSETS & LOGS
# =========================================================

class Asset(models.Model):
    STATUS_CHOICES = [('active', 'Active'), ('maintenance', 'Maintenance'), ('disposed', 'Disposed')]
    department = models.ForeignKey(Department, on_delete=models.CASCADE, related_name='assets')
    name = models.CharField(max_length=100)
    reg_number = models.CharField(max_length=20, unique=True)
    purchase_value = models.DecimalField(max_digits=12, decimal_places=2)
    purchase_date = models.DateField()
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='active')
    description = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['name']
        indexes = [models.Index(fields=['reg_number'])]

    def __str__(self):
        return f"{self.name} [{self.reg_number}]"

class MaintenanceRecord(models.Model):
    asset = models.ForeignKey(Asset, on_delete=models.CASCADE, related_name='maintenance_records')
    service_date = models.DateField()
    cost = models.DecimalField(max_digits=10, decimal_places=2)
    description = models.TextField()
    serviced_by = models.CharField(max_length=255, blank=True)
    next_service_date = models.DateField(blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)

class TransportLog(models.Model):
    asset = models.ForeignKey(Asset, on_delete=models.CASCADE, related_name='transport_logs')
    date = models.DateField()
    fuel_cost = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))
    mileage_reading = models.PositiveIntegerField()
    distance_covered = models.PositiveIntegerField(default=0)
    loading_fees = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))
    unforeseen_expenses = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))
    driver_name = models.CharField(max_length=255, blank=True)
    destination = models.CharField(max_length=255, blank=True)
    description = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    @property
    def total_expenses(self):
        return self.fuel_cost + self.loading_fees + self.unforeseen_expenses

# =========================================================
# FARMING ACTIVITIES
# =========================================================

class FarmingActivity(models.Model):
    CATEGORY_CHOICES = [('livestock', 'Livestock'), ('crops', 'Crops')]
    category = models.CharField(max_length=20, choices=CATEGORY_CHOICES)
    activity_name = models.CharField(max_length=100)
    feeds_cost = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))
    vet_expenses = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))
    land_prep_cost = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))
    labor_cost = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))
    fertilizer_cost = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))
    seed_cost = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0.00'))
    sales_income = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0.00'))
    date_started = models.DateField()
    closing_date = models.DateField(blank=True, null=True)
    is_closed = models.BooleanField(default=False)
    notes = models.TextField(blank=True,null=True)
    created_at = models.DateTimeField(auto_now_add=True)

    @property
    def total_costs(self):
        return (self.feeds_cost + self.vet_expenses + self.land_prep_cost + 
                self.labor_cost + self.fertilizer_cost + self.seed_cost)

    @property
    def net_margin(self):
        return self.sales_income - self.total_costs

# =========================================================
# FINANCE
# =========================================================

class Expense(models.Model):
    department = models.ForeignKey(Department, on_delete=models.CASCADE, related_name='expenses')
    category = models.CharField(max_length=100)
    amount = models.DecimalField(max_digits=12, decimal_places=2)
    expense_date = models.DateField()
    description = models.TextField(blank=True)
    # Changed from User to settings.AUTH_USER_MODEL
    recorded_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True)
    created_at = models.DateTimeField(auto_now_add=True)


from django.db import models
from django.contrib.contenttypes.models import ContentType
from django.contrib.contenttypes.fields import GenericForeignKey
from Users.models import CustomUser
from django.db import models
from django.contrib.auth.models import User

class ChatMessage(models.Model):
    sender = models.ForeignKey(CustomUser, on_delete=models.CASCADE, related_name='sent_messages')
    message = models.TextField()
    timestamp = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['timestamp']

    def __str__(self):
        return f"{self.sender.username}: {self.message[:20]}"
class GeneralLedger(models.Model):
    # Your existing choices and fields...
    TRANSACTION_TYPES = [('income', 'Income'), ('expense', 'Expense')]
    
    transaction_type = models.CharField(max_length=10, choices=TRANSACTION_TYPES)
    amount = models.DecimalField(max_length=12, decimal_places=2, max_digits=12)
    description = models.TextField()
    ledger_source = models.CharField(max_length=50, blank=True, null=True)
    reference_number = models.CharField(max_length=50)
    department = models.ForeignKey('Department', on_delete=models.SET_NULL, null=True, blank=True)
    timestamp = models.DateTimeField(auto_now_add=True)
    created_by = models.ForeignKey(CustomUser, on_delete=models.PROTECT)

    # Add these fields to handle the generic framework relationship:
    content_type = models.ForeignKey(ContentType, on_delete=models.CASCADE, null=True, blank=True)
    object_id = models.PositiveIntegerField(null=True, blank=True)
    content_object = GenericForeignKey('content_type', 'object_id')

    def __str__(self):
        return f"{self.transaction_type.upper()} - {self.reference_number}"