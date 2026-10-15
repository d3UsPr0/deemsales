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


class Waiter(models.Model):
    """Waiters with full name and display name"""
    full_name = models.CharField(max_length=100)
    display_name = models.CharField(max_length=50)
    is_active = models.BooleanField(default=True)
    
    def __str__(self):
        return self.display_name


class Order(models.Model):
    """Individual order entry"""
    LOCATION_CHOICES = [
        ('table', 'Table'),
        ('takeaway', 'Takeaway'),
        ('other', 'Other'),
    ]
    
    menu_item = models.ForeignKey(MenuItem, on_delete=models.CASCADE)
    quantity = models.IntegerField(default=1)
    unit_price = models.DecimalField(max_digits=10, decimal_places=2)
    total_price = models.DecimalField(max_digits=10, decimal_places=2)
    waiter = models.ForeignKey(Waiter, on_delete=models.SET_NULL, null=True, blank=True)
    location_type = models.CharField(max_length=20, choices=LOCATION_CHOICES, default='table')
    location_detail = models.CharField(max_length=100, blank=True, null=True, help_text="Table number, takeaway number, or other location")
    notes = models.TextField(blank=True, null=True)
    status = models.CharField(max_length=10, choices=[('open', 'Open'), ('closed', 'Closed')], default='open')
    closed_at = models.DateTimeField(null=True, blank=True)
    created_by = models.ForeignKey(User, on_delete=models.CASCADE)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    session_id = models.CharField(max_length=100, blank=True, null=True)  # To group orders submitted together
    
    class Meta:
        ordering = ['-created_at']
    
    def __str__(self):
        location = self.get_location_display()
        if self.location_detail:
            location += f" ({self.location_detail})"
        return f"{self.menu_item.name} x{self.quantity} - {self.waiter} - {location}"
    
    def save(self, *args, **kwargs):
        if not self.total_price:
            self.total_price = self.unit_price * self.quantity
        super().save(*args, **kwargs)
    
    def get_location_display(self):
        """Get full location display"""
        if self.location_type == 'other':
            return self.location_detail or 'Other'
        elif self.location_detail:
            return f"{self.get_location_type_display()} #{self.location_detail}"
        else:
            return self.get_location_type_display()
        
from django.db.models.signals import post_save
from django.dispatch import receiver


class UserProfile(models.Model):
    """Extended user profile with role"""
    ROLE_CHOICES = [
        ('admin', 'Administrator'),
        ('manager', 'Manager'),
        ('cashier', 'Cashier'),
    ]
    
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='profile')
    role = models.CharField(max_length=20, choices=ROLE_CHOICES, default='cashier')
    phone = models.CharField(max_length=20, blank=True, null=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    
    def __str__(self):
        return f"{self.user.username} - {self.get_role_display()}"
    
    # Helper properties
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
    def can_view_reports(self):
        """Admin and Manager can view reports"""
        return self.role in ['admin', 'manager']
    
    @property
    def can_manage_users(self):
        """Only Admin can manage users"""
        return self.role == 'admin'


@receiver(post_save, sender=User)
def create_user_profile(sender, instance, created, **kwargs):
    """Auto-create profile when user is created"""
    if created:
        UserProfile.objects.get_or_create(user=instance)
    else:
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
    
    MODEL_CHOICES = [
        ('Order', 'Order'),
        ('MenuItem', 'Menu Item'),
        ('MenuCategory', 'Menu Category'),
        ('Waiter', 'Waiter'),
        ('Expense', 'Expense'),
        ('User', 'User'),
    ]
    
    action = models.CharField(max_length=20, choices=ACTION_CHOICES)
    model_name = models.CharField(max_length=50)
    object_id = models.CharField(max_length=50)
    object_repr = models.CharField(max_length=255, help_text="Human-readable name of the object")
    
    # Snapshot before and after (for update) or just before (for delete)
    data_before = models.JSONField(null=True, blank=True)
    data_after = models.JSONField(null=True, blank=True)
    
    # Who did it
    performed_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='audit_logs')
    performed_at = models.DateTimeField(auto_now_add=True)
    
    # Extra info
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