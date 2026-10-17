from django.contrib import admin
from .models import MenuCategory, MenuItem, Order, UserProfile, Expense, AuditLog


@admin.register(MenuCategory)
class MenuCategoryAdmin(admin.ModelAdmin):
    list_display = ['name', 'display_order']
    list_editable = ['display_order']


@admin.register(MenuItem)
class MenuItemAdmin(admin.ModelAdmin):
    list_display = ['name', 'category', 'price', 'is_available', 'display_order']
    list_filter = ['category', 'is_available']
    list_editable = ['price', 'is_available', 'display_order']


@admin.register(Order)
class OrderAdmin(admin.ModelAdmin):
    list_display = ['menu_item', 'quantity', 'total_price', 'waiter_user', 'created_by', 'status', 'created_at']
    list_filter = ['created_at', 'status', 'waiter_user', 'menu_item__category']
    readonly_fields = ['total_price', 'created_at', 'updated_at']
    raw_id_fields = ['waiter_user', 'created_by', 'menu_item']
    search_fields = ['menu_item__name', 'waiter_user__username', 'waiter_user__first_name', 'waiter_user__last_name']


@admin.register(UserProfile)
class UserProfileAdmin(admin.ModelAdmin):
    list_display = ['user', 'role', 'phone', 'is_active']
    list_filter = ['role', 'is_active']
    search_fields = ['user__username', 'user__first_name', 'user__last_name']
    raw_id_fields = ['user']


@admin.register(Expense)
class ExpenseAdmin(admin.ModelAdmin):
    list_display = ['date', 'category', 'description', 'amount', 'recorded_by']
    list_filter = ['date', 'category']
    search_fields = ['description', 'receipt_ref']
    raw_id_fields = ['recorded_by']


@admin.register(AuditLog)
class AuditLogAdmin(admin.ModelAdmin):
    list_display = ['performed_at', 'action', 'model_name', 'object_id', 'object_repr', 'performed_by']
    list_filter = ['action', 'model_name', 'performed_at']
    search_fields = ['object_repr', 'performed_by__username']
    readonly_fields = ['action', 'model_name', 'object_id', 'object_repr', 
                       'data_before', 'data_after', 'performed_by', 'performed_at', 'ip_address']
    raw_id_fields = ['performed_by']
    
    def has_add_permission(self, request):
        # Prevent manual creation of audit logs
        return False
    
    def has_change_permission(self, request, obj=None):
        # Audit logs are read-only
        return False