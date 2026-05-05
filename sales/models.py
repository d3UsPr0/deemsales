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