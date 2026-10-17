from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone


class MenuCategory(models.Model):
    """Categories: Breakfast, Main, Drinks, Choma"""
    name = models.CharField(max_length=50, unique=True)
    display_order = models.IntegerField(default=0)
    
    class Meta:
        verbose_name_plural = "Menu Categories"
        ordering = ['display_order']
    
    def __str__(self):
        return self.name


class MenuItem(models.Model):
    """Individual menu items with prices"""
    category = models.ForeignKey(MenuCategory, on_delete=models.CASCADE, related_name='items')
    name = models.CharField(max_length=100)
    price = models.DecimalField(max_digits=10, decimal_places=2)
    is_available = models.BooleanField(default=True)
    display_order = models.IntegerField(default=0)
    
    class Meta:
        ordering = ['category__display_order', 'display_order']
    
    def __str__(self):
        return f"{self.name} - {self.price} TZS"


class Order(models.Model):
    """Individual order entry — waiter is now a User with role='waiter'"""
    LOCATION_CHOICES = [
        ('table', 'Table'),
        ('takeaway', 'Takeaway'),
        ('other', 'Other'),
    ]
    
    menu_item = models.ForeignKey(MenuItem, on_delete=models.CASCADE)
    quantity = models.IntegerField(default=1)
    unit_price = models.DecimalField(max_digits=10, decimal_places=2)
    total_price = models.DecimalField(max_digits=10, decimal_places=2)
    
    # Waiter is now a User (with role='waiter') — null for non-waiter submitters who choose "Self"
    waiter_user = models.ForeignKey(
        User, 
        on_delete=models.SET_NULL, 
        null=True, 
        blank=True, 
        related_name='orders_as_waiter',
        help_text="The waiter serving this order (a User with role='waiter', or self)"
    )
    
    location_type = models.CharField(max_length=20, choices=LOCATION_CHOICES, default='table')
    location_detail = models.CharField(max_length=100, blank=True, null=True, help_text="Table number, takeaway number, or other location")
    notes = models.TextField(blank=True, null=True)
    status = models.CharField(max_length=10, choices=[('open', 'Open'), ('closed', 'Closed')], default='open')
    closed_at = models.DateTimeField(null=True, blank=True)
    
    created_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='orders_created')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    session_id = models.CharField(max_length=100, blank=True, null=True)
    
    class Meta:
        ordering = ['-created_at']
    
    def __str__(self):
        location = self.get_full_location_display()
        waiter_name = self.get_waiter_display_name()
        return f"{self.menu_item.name} x{self.quantity} - {waiter_name} - {location}"
    
    def save(self, *args, **kwargs):
        if not self.total_price:
            self.total_price = self.unit_price * self.quantity
        super().save(*args, **kwargs)
    
    def get_waiter_display_name(self):
        """Get a friendly display name for the waiter"""
        if self.waiter_user:
            full_name = self.waiter_user.get_full_name()
            return full_name if full_name else self.waiter_user.username
        return "N/A"
    
    def get_full_location_display(self):
        """Get full location display"""
        if self.location_type == 'other':
            return self.location_detail or 'Other'
        elif self.location_detail:
            return f"{self.get_location_type_display()} #{self.location_detail}"
        else:
            return self.get_location_type_display()


class UserProfile(models.Model):
    """Extended user profile with role"""
    ROLE_CHOICES = [
        ('admin', 'Administrator'),
        ('manager', 'Manager'),
        ('cashier', 'Cashier'),
        ('waiter', 'Waiter'),
    ]
    
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='profile')
    role = models.CharField(max_length=20, choices=ROLE_CHOICES, default='cashier')
    phone = models.CharField(max_length=20, blank=True, null=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    
    def __str__(self):
        return f"{self.user.username} - {self.get_role_display()}"
    
    @property
    def is_admin(self):
        return self.role == 'admin'
    
    @property
    def is_manager(self):
        return self.role == 'manager'
    
    @property
    def is_cashier(self):
        return self.role == 'cashier'
    
    @property
    def is_waiter(self):
        return self.role == 'waiter'
    
    @property
    def can_view_reports(self):
        return self.role in ['admin', 'manager']
    
    @property
    def can_manage_users(self):
        return self.role == 'admin'
    
    @property
    def can_manage_menu(self):
        return self.role in ['admin', 'manager']
    
    @property
    def can_record_expenses(self):
        return self.role in ['admin', 'manager', 'cashier']
    
    @property
    def display_name(self):
        """Friendly display name for this user"""
        full_name = self.user.get_full_name()
        return full_name if full_name else self.user.username


from django.db.models.signals import post_save
from django.dispatch import receiver


@receiver(post_save, sender=User)
def create_user_profile(sender, instance, created, **kwargs):
    """Auto-create profile when user is created"""
    UserProfile.objects.get_or_create(user=instance)


class Expense(models.Model):
    """Daily expense records"""
    CATEGORY_CHOICES = [
        ('food', 'Food & Ingredients'),
        ('utilities', 'Utilities (Water, Electricity)'),
        ('rent', 'Rent'),
        ('salaries', 'Staff Salaries'),
        ('transport', 'Transport'),
        ('maintenance', 'Maintenance & Repairs'),
        ('supplies', 'Supplies & Packaging'),
        ('marketing', 'Marketing'),
        ('other', 'Other'),
    ]
    
    category = models.CharField(max_length=30, choices=CATEGORY_CHOICES, default='other')
    amount = models.DecimalField(max_digits=12, decimal_places=2)
    description = models.CharField(max_length=255)
    date = models.DateField(default=timezone.now)
    receipt_ref = models.CharField(max_length=100, blank=True, null=True, help_text="Receipt/Bill number")
    recorded_by = models.ForeignKey(User, on_delete=models.CASCADE)
    created_at = models.DateTimeField(auto_now_add=True)
    
    class Meta:
        ordering = ['-date', '-created_at']
    
    def __str__(self):
        return f"{self.get_category_display()} - {self.amount} TZS ({self.date})"


class AuditLog(models.Model):
    """Tracks UPDATE and DELETE actions on database rows"""
    ACTION_CHOICES = [
        ('update', 'Update'),
        ('delete', 'Delete'),
    ]
    
    action = models.CharField(max_length=20, choices=ACTION_CHOICES)
    model_name = models.CharField(max_length=50)
    object_id = models.CharField(max_length=50)
    object_repr = models.CharField(max_length=255, help_text="Human-readable name of the object")
    
    data_before = models.JSONField(null=True, blank=True)
    data_after = models.JSONField(null=True, blank=True)
    
    performed_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='audit_logs')
    performed_at = models.DateTimeField(auto_now_add=True)
    
    notes = models.CharField(max_length=255, blank=True, null=True)
    ip_address = models.GenericIPAddressField(null=True, blank=True)
    
    class Meta:
        ordering = ['-performed_at']
        verbose_name = "Audit Log"
        verbose_name_plural = "Audit Logs"
        indexes = [
            models.Index(fields=['-performed_at']),
            models.Index(fields=['model_name', 'object_id']),
            models.Index(fields=['action']),
        ]
    
    def __str__(self):
        return f"[{self.get_action_display()}] {self.model_name} #{self.object_id} by {self.performed_by}"
    
    def get_changed_fields(self):
        """Returns list of fields that changed between before and after"""
        if self.action != 'update' or not self.data_before or not self.data_after:
            return []
        
        changed = []
        all_keys = set(self.data_before.keys()) | set(self.data_after.keys())
        
        for key in all_keys:
            before = self.data_before.get(key)
            after = self.data_after.get(key)
            if before != after:
                changed.append({
                    'field': key,
                    'before': before,
                    'after': after,
                })
        return changed