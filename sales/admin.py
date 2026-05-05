from django.contrib import admin
from .models import MenuCategory, MenuItem, Waiter, Order


@admin.register(MenuCategory)
class MenuCategoryAdmin(admin.ModelAdmin):
    list_display = ['name', 'display_order']
    list_editable = ['display_order']


@admin.register(MenuItem)
class MenuItemAdmin(admin.ModelAdmin):
    list_display = ['name', 'category', 'price', 'is_available', 'display_order']
    list_filter = ['category', 'is_available']
    list_editable = ['price', 'is_available', 'display_order']


@admin.register(Waiter)
class WaiterAdmin(admin.ModelAdmin):
    list_display = ['full_name', 'display_name', 'is_active']


@admin.register(Order)
class OrderAdmin(admin.ModelAdmin):
    list_display = ['menu_item', 'quantity', 'total_price', 'waiter', 'created_by', 'created_at']
    list_filter = ['created_at', 'waiter', 'menu_item__category']
    readonly_fields = ['total_price', 'created_at', 'updated_at']