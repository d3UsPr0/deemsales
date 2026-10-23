from django.urls import path
from . import views

urlpatterns = [
    # Auth
    path('', views.home_view, name='home'),
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('keep-alive/', views.keep_alive, name='keep_alive'),

    # Orders
    path('submit-orders/', views.submit_orders, name='submit_orders'),
    path('edit-order/<int:order_id>/', views.edit_order, name='edit_order'),
    path('update-order/<int:order_id>/', views.update_order, name='update_order'),
    path('delete-order/<int:order_id>/', views.delete_order, name='delete_order'),
    path('delete-group/<str:session_id>/', views.delete_group_orders, name='delete_group_orders'),
    path('view-group/<str:session_id>/', views.view_group_orders, name='view_group_orders'),
    path('order-group-info/<str:session_id>/', views.get_order_group_info, name='order_group_info'),
    path('add-to-order/<str:session_id>/', views.add_to_order, name='add_to_order'),
    path('toggle-group-status/<str:session_id>/', views.toggle_group_status, name='toggle_group_status'),
    path('order-group-full/<str:session_id>/', views.get_order_group_full, name='get_order_group_full'),
    path('update-order-group/<str:session_id>/', views.update_order_group, name='update_order_group'),

    # Reports / Export
    path('export-excel/', views.export_excel, name='export_excel'),

    # User Management (admin only)
    path('users/', views.user_management, name='user_management'),
    path('users/add/', views.add_user, name='add_user'),
    path('users/edit/<int:user_id>/', views.edit_user, name='edit_user'),
    path('users/toggle/<int:user_id>/', views.toggle_user_status, name='toggle_user_status'),
    path('users/delete/<int:user_id>/', views.delete_user, name='delete_user'),
    path('users/reset-password/<int:user_id>/', views.reset_password, name='reset_password'),

    # Menu Management (admin + manager)
    path('menu/', views.menu_management, name='menu_management'),

    # Categories
    path('menu/category/add/', views.add_category, name='add_category'),
    path('menu/category/edit/<int:category_id>/', views.edit_category, name='edit_category'),
    path('menu/category/delete/<int:category_id>/', views.delete_category, name='delete_category'),

    # Menu Items
    path('menu/item/add/', views.add_menu_item, name='add_menu_item'),
    path('menu/item/edit/<int:item_id>/', views.edit_menu_item, name='edit_menu_item'),
    path('menu/item/delete/<int:item_id>/', views.delete_menu_item, name='delete_menu_item'),
    path('menu/item/toggle/<int:item_id>/', views.toggle_item_availability, name='toggle_item_availability'),

    # Expenses
    path('expenses/', views.expenses_view, name='expenses_view'),
    path('expenses/add/', views.add_expense, name='add_expense'),
    path('expenses/edit/<int:expense_id>/', views.edit_expense, name='edit_expense'),
    path('expenses/delete/<int:expense_id>/', views.delete_expense, name='delete_expense'),
    path('expenses/get/<int:expense_id>/', views.get_expense, name='get_expense'),

    # Audit Log
    path('audit-log/', views.audit_log_view, name='audit_log'),
    path('audit-log/detail/<int:log_id>/', views.audit_log_detail, name='audit_log_detail'),
    
    # Get orders page (for pagination)
    path('orders-page/', views.get_orders_page, name='get_orders_page'),
]