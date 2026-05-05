from django.urls import path
from . import views

urlpatterns = [
    path('', views.home_view, name='home'),
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('submit-orders/', views.submit_orders, name='submit_orders'),
    path('edit-order/<int:order_id>/', views.edit_order, name='edit_order'),
    path('update-order/<int:order_id>/', views.update_order, name='update_order'),
    path('delete-order/<int:order_id>/', views.delete_order, name='delete_order'),
    path('export-excel/', views.export_excel, name='export_excel'),
    path('delete-group/<str:session_id>/', views.delete_group_orders, name='delete_group_orders'),
    path('view-group/<str:session_id>/', views.view_group_orders, name='view_group_orders'),
    path('order-group-info/<str:session_id>/', views.get_order_group_info, name='order_group_info'),
    path('add-to-order/<str:session_id>/', views.add_to_order, name='add_to_order'),
    path('toggle-group-status/<str:session_id>/', views.toggle_group_status, name='toggle_group_status'),
    
]