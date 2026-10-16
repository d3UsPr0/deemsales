from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required, user_passes_test
from django.contrib import messages
from django.http import HttpResponse, JsonResponse
from django.db import transaction
from django.utils import timezone
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import json
from django.db.models import Q
from .models import MenuCategory, MenuItem, Waiter, Order
from collections import defaultdict, OrderedDict
OD = OrderedDict
from functools import wraps
from django.core.exceptions import PermissionDenied


def get_user_role(user):
    """Get the role of a user"""
    if user.is_superuser:
        return 'admin'
    try:
        return user.profile.role
    except:
        return 'cashier'  # Default role


def role_required(*allowed_roles):
    """Decorator to restrict access by role"""
    def decorator(view_func):
        @wraps(view_func)
        @login_required
        def wrapper(request, *args, **kwargs):
            role = get_user_role(request.user)
            if role not in allowed_roles:
                messages.error(request, 'You do not have permission to access this page.')
                return redirect('home')
            return view_func(request, *args, **kwargs)
        return wrapper
    return decorator


def admin_required(view_func):
    """Only admin can access"""
    return role_required('admin')(view_func)


def reports_required(view_func):
    """Admin and Manager can access"""
    return role_required('admin', 'manager')(view_func)


def is_admin(user):
    """Check if user is admin role"""
    return get_user_role(user) == 'admin'


def can_view_reports(user):
    """Check if user can view reports"""
    return get_user_role(user) in ['admin', 'manager']


def is_admin(user):
    return user.is_staff or user.is_superuser


def login_view(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        password = request.POST.get('password')
        user = authenticate(request, username=username, password=password)
        
        if user is not None:
            login(request, user)
            request.session['last_activity'] = timezone.now().timestamp()
            messages.success(request, f'Welcome, {user.username}!')
            return redirect('home')
        else:
            messages.error(request, 'Invalid credentials')
    
    # Check for timeout
    if request.GET.get('timeout') == '1':
        messages.warning(request, 'Your session expired due to inactivity. Please log in again.')
    
    return render(request, 'sales/login.html')


def logout_view(request):
    logout(request)
    messages.info(request, 'You have been logged out')
    return redirect('login')

@login_required
def keep_alive(request):
    """Endpoint to keep session alive when user clicks 'Stay Logged In'"""
    request.session['last_activity'] = timezone.now().timestamp()
    request.session.modified = True
    return JsonResponse({'status': 'ok', 'message': 'Session extended'})


@login_required
def home_view(request):
    """Main order entry page"""
    categories = MenuCategory.objects.all()
    menu_items = MenuItem.objects.filter(is_available=True).select_related('category')
    waiters = Waiter.objects.filter(is_active=True)
    
    # Get today's orders grouped by session_id
    today = timezone.now().date()
    today_orders = Order.objects.filter(
        created_at__date=today
    ).select_related('menu_item', 'waiter', 'created_by').order_by('-created_at')
    
    # Group orders by session_id and combine duplicate items
    # from collections import defaultdict, OrderedDict as OD
    from django.db.models import Sum
    
    grouped_orders = []
    session_groups = defaultdict(list)
    
    for order in today_orders:
        session_key = order.session_id if order.session_id else f"single_{order.id}"
        session_groups[session_key].append(order)
    
    # Process each group and combine duplicates
    order_number = 1  # Order counter
    
    for session_key, orders in session_groups.items():
        # Combine duplicate menu items by summing quantities
        combined_items = OrderedDict()
        
        for order in orders:
            item_key = order.menu_item.id
            
            if item_key in combined_items:
                # Update existing item
                combined_items[item_key]['quantity'] += order.quantity
                combined_items[item_key]['total_price'] += float(order.total_price)
                combined_items[item_key]['order_ids'].append(order.id)
            else:
                # Add new item
                combined_items[item_key] = {
                    'menu_item': order.menu_item,
                    'quantity': order.quantity,
                    'unit_price': float(order.unit_price),
                    'total_price': float(order.total_price),
                    'category': order.menu_item.category,
                    'order_ids': [order.id],
                }
        
        # Create group data with combined items
        combined_list = list(combined_items.values())
        
                # Determine payment status
        all_closed = all(order.status == 'closed' for order in orders)
        any_closed = any(order.status == 'closed' for order in orders)
        if all_closed:
            payment_status = 'closed'
        elif any_closed:
            payment_status = 'partial'
        else:
            payment_status = 'open'

        group_data = {
            'order_number': order_number,
            'session_id': session_key,
            'combined_items': combined_list,
            'created_at': orders[0].created_at,
            'waiter': orders[0].waiter,
            'created_by': orders[0].created_by,
            'location_type': orders[0].location_type,
            'location_detail': orders[0].location_detail,
            'location_display': orders[0].get_location_display(),
            'notes': orders[0].notes if orders[0].notes else '',
            'total_quantity': sum(item['quantity'] for item in combined_list),
            'total_amount': sum(item['total_price'] for item in combined_list),
            'item_count': len(combined_list),
            'is_group': len(orders) > 1 or len(combined_list) > 1,
            'first_order_id': orders[-1].id,
            'last_order_id': orders[0].id,
            'payment_status': payment_status,
        }
        grouped_orders.append(group_data)
        order_number += 1
    
    # Sort by order number (newest first based on creation time)
    grouped_orders.sort(key=lambda x: x['created_at'], reverse=True)
    # Re-assign order numbers after sorting
    for idx, group in enumerate(grouped_orders, 1):
        group['order_number'] = idx
    
    # Calculate total collection for today
    total_collection = today_orders.aggregate(total=Sum('total_price'))['total'] or 0
        # Count open/closed
    open_count = 0
    closed_count = 0
    for session_key, orders in session_groups.items():
        if all(order.status == 'closed' for order in orders):
            closed_count += 1
        else:
            open_count += 1

        # ============================================================
    # UNPAID ORDERS — all OPEN orders from any date
    # ============================================================
    unpaid_orders = Order.objects.filter(
        status='open'
    ).select_related('menu_item', 'waiter', 'created_by').order_by('-created_at')
    
    # Group unpaid orders by session_id
    unpaid_groups = defaultdict(list)
    for order in unpaid_orders:
        session_key = order.session_id if order.session_id else f"single_{order.id}"
        unpaid_groups[session_key].append(order)
    
    unpaid_grouped = []
    unpaid_order_num = 1
    
    for session_key, orders in unpaid_groups.items():
        combined_items = OD()
        for order in orders:
            item_key = order.menu_item.id
            if item_key in combined_items:
                combined_items[item_key]['quantity'] += order.quantity
                combined_items[item_key]['total_price'] += float(order.total_price)
                combined_items[item_key]['order_ids'].append(order.id)
            else:
                combined_items[item_key] = {
                    'menu_item': order.menu_item,
                    'quantity': order.quantity,
                    'unit_price': float(order.unit_price),
                    'total_price': float(order.total_price),
                    'category': order.menu_item.category,
                    'order_ids': [order.id],
                }
        
        combined_list = list(combined_items.values())
        first_order = orders[0]
        
        # Location display
        if first_order.location_type == 'other':
            location_display = first_order.location_detail or 'Other'
        elif first_order.location_detail:
            location_display = f"{first_order.get_location_type_display()} #{first_order.location_detail}"
        else:
            location_display = first_order.get_location_type_display()
        
        unpaid_grouped.append({
            'order_number': unpaid_order_num,
            'session_id': session_key,
            'combined_items': combined_list,
            'created_at': first_order.created_at,
            'waiter': first_order.waiter,
            'created_by': first_order.created_by,
            'location_type': first_order.location_type,
            'location_detail': first_order.location_detail,
            'location_display': location_display,
            'notes': first_order.notes if first_order.notes else '',
            'total_quantity': sum(item['quantity'] for item in combined_list),
            'total_amount': sum(item['total_price'] for item in combined_list),
            'item_count': len(combined_list),
            'is_group': len(orders) > 1 or len(combined_list) > 1,
            'first_order_id': orders[-1].id,
            'last_order_id': orders[0].id,
            'days_old': (today - first_order.created_at.date()).days,
        })
        unpaid_order_num += 1
    
    # Sort unpaid by newest first
    unpaid_grouped.sort(key=lambda x: x['created_at'], reverse=True)
    for idx, g in enumerate(unpaid_grouped, 1):
        g['order_number'] = idx
    
    unpaid_total = sum(g['total_amount'] for g in unpaid_grouped)
    
    context = {
        'categories': categories,
        'menu_items': menu_items,
        'waiters': waiters,
        'grouped_orders': grouped_orders,
        'unpaid_grouped': unpaid_grouped,
        'unpaid_total': unpaid_total,
        'unpaid_count': len(unpaid_grouped),
        'total_collection': total_collection,
        'today': today,
        'total_orders_count': today_orders.count(),
        'total_sessions': len(grouped_orders),
        'open_count': open_count,
        'closed_count': closed_count,
        'user_role': get_user_role(request.user),
        'can_view_reports': can_view_reports(request.user),
        'can_manage_users': get_user_role(request.user) == 'admin',
    }
    
    return render(request, 'sales/home.html', context)

@login_required
def submit_orders(request):
    """Handle order submission"""
    if request.method == 'POST':
        try:
            with transaction.atomic():
                session_id = datetime.now().strftime('%Y%m%d%H%M%S') + str(request.user.id)
                
                # Get waiter
                waiter_id = request.POST.get('waiter')
                waiter = Waiter.objects.get(id=waiter_id) if waiter_id else None
                notes = request.POST.get('notes', '')
                
                # Get location
                location_type = request.POST.get('location_type', 'table')
                location_detail = request.POST.get('location_detail', '')
                
                # Process each menu item
                menu_items = MenuItem.objects.filter(is_available=True)
                orders_created = 0
                
                for item in menu_items:
                    is_checked = request.POST.get(f'item_{item.id}')
                    quantity_str = request.POST.get(f'quantity_{item.id}', '0')
                    
                    if is_checked == 'on':
                        quantity = int(quantity_str) if quantity_str else 1
                        
                        if quantity > 0:
                            # Create separate entries for each quantity
                            for _ in range(quantity):
                                Order.objects.create(
                                    menu_item=item,
                                    quantity=1,
                                    unit_price=item.price,
                                    total_price=item.price,
                                    waiter=waiter,
                                    location_type=location_type,
                                    location_detail=location_detail if location_detail else None,
                                    notes=notes,
                                    created_by=request.user,
                                    session_id=session_id
                                )
                                orders_created += 1
                
                if orders_created > 0:
                    location_msg = f"Location: {Order.LOCATION_CHOICES[['table','takeaway','other'].index(location_type)][1]}"
                    if location_detail:
                        location_msg += f" - {location_detail}"
                    messages.success(request, f'{orders_created} orders recorded successfully! {location_msg}')
                else:
                    messages.warning(request, 'No items selected. Please select at least one item.')
                
                return redirect('home')
                
        except Exception as e:
            messages.error(request, f'Error saving orders: {str(e)}')
            return redirect('home')
    
    return redirect('home')


@login_required
def edit_order(request, order_id):
    """Get order data for editing"""
    if request.method == 'GET':
        order = get_object_or_404(Order, id=order_id)
        data = {
            'id': order.id,
            'menu_item_id': order.menu_item.id,
            'menu_item_name': order.menu_item.name,
            'quantity': order.quantity,
            'waiter_id': order.waiter.id if order.waiter else None,
            'waiter_name': order.waiter.display_name if order.waiter else '',
            'location_type': order.location_type,
            'location_detail': order.location_detail or '',
            'notes': order.notes or '',
        }
        return JsonResponse(data)


@login_required
def update_order(request, order_id):
    """Update an existing order (with audit logging)"""
    if request.method == 'POST':
        try:
            order = get_object_or_404(Order, id=order_id)
            
            # Capture BEFORE state
            from .audit import serialize_instance, log_action
            data_before = serialize_instance(order)
            
            menu_item_id = request.POST.get('menu_item')
            quantity = int(request.POST.get('quantity', 1))
            waiter_id = request.POST.get('waiter')
            location_type = request.POST.get('location_type', 'table')
            location_detail = request.POST.get('location_detail', '')
            notes = request.POST.get('notes', '')
            
            order.menu_item = MenuItem.objects.get(id=menu_item_id)
            order.quantity = quantity
            order.total_price = order.menu_item.price * quantity
            order.unit_price = order.menu_item.price
            order.waiter = Waiter.objects.get(id=waiter_id) if waiter_id else None
            order.location_type = location_type
            order.location_detail = location_detail if location_detail else None
            order.notes = notes
            order.save()
            
            # Log the update
            log_action(request, 'update', order, data_before=data_before)
            
            messages.success(request, 'Order updated successfully!')
            return redirect('home')
            
        except Exception as e:
            messages.error(request, f'Error updating order: {str(e)}')
            return redirect('home')
    
    return redirect('home')

@login_required
def delete_order(request, order_id):
    """Delete an order (with audit logging)"""
    if request.method == 'POST':
        order = get_object_or_404(Order, id=order_id)
        
        # Log the delete BEFORE deleting
        from .audit import serialize_instance, log_action
        data_before = serialize_instance(order)
        log_action(request, 'delete', order, data_before=data_before)
        
        order.delete()
        messages.success(request, 'Order deleted successfully!')
    return redirect('home')

@login_required
def delete_group_orders(request, session_id):
    """Delete all orders in a group session (with audit logging)"""
    if request.method == 'POST':
        orders = Order.objects.filter(session_id=session_id)
        count = orders.count()
        
        # Log each order individually
        from .audit import serialize_instance, log_action
        for order in orders:
            data_before = serialize_instance(order)
            log_action(request, 'delete', order, data_before=data_before, 
                       notes=f'Group delete ({count} items)')
        
        orders.delete()
        messages.success(request, f'Deleted {count} orders from the group.')
    return redirect('home')


@login_required
def view_group_orders(request, session_id):
    """Get details of a group order"""
    if request.method == 'GET':
        orders = Order.objects.filter(session_id=session_id).select_related('menu_item', 'waiter')
        
        orders_data = []
        total = 0
        for order in orders:
            orders_data.append({
                'id': order.id,
                'menu_item': order.menu_item.name,
                'category': order.menu_item.category.name,
                'quantity': order.quantity,
                'unit_price': float(order.unit_price),
                'total_price': float(order.total_price),
            })
            total += float(order.total_price)
        
        return JsonResponse({
            'orders': orders_data,
            'total': total,
            'waiter': orders[0].waiter.display_name if orders and orders[0].waiter else 'N/A',
            'notes': orders[0].notes if orders else '',
        })

@reports_required
def export_excel(request):
    """Export orders, expenses, and item sales to Excel with 3 sheets"""
    from datetime import timedelta
    from collections import defaultdict, OrderedDict as OD
    
    today = timezone.now().date()
    date_from = None
    date_to = None
    
    # Preset handling
    preset = request.GET.get('preset')
    
    if preset == 'today':
        date_from = today
        date_to = today
        label = f"Today ({today})"
    elif preset == 'yesterday':
        date_from = today - timedelta(days=1)
        date_to = today - timedelta(days=1)
        label = f"Yesterday ({date_from})"
    elif preset == 'this_week':
        date_from = today - timedelta(days=today.weekday())
        date_to = today
        label = f"This Week ({date_from} to {date_to})"
    elif preset == 'this_month':
        date_from = today.replace(day=1)
        date_to = today
        label = f"This Month ({date_from} to {date_to})"
    elif preset == 'last_month':
        first_of_this_month = today.replace(day=1)
        last_day_of_last_month = first_of_this_month - timedelta(days=1)
        date_from = last_day_of_last_month.replace(day=1)
        date_to = last_day_of_last_month
        label = f"Last Month ({date_from} to {date_to})"
    else:
        date_from_str = request.GET.get('date_from')
        date_to_str = request.GET.get('date_to')
        
        if date_from_str:
            date_from = datetime.strptime(date_from_str, '%Y-%m-%d').date()
        if date_to_str:
            date_to = datetime.strptime(date_to_str, '%Y-%m-%d').date()
        
        if date_from and date_to:
            label = f"{date_from} to {date_to}"
        elif date_from:
            label = f"From {date_from}"
            date_to = today
        elif date_to:
            label = f"Until {date_to}"
            date_from = today - timedelta(days=30)
        else:
            date_from = today
            date_to = today
            label = f"Today ({today})"
    
    # Fetch data
    orders = Order.objects.filter(
        created_at__date__gte=date_from,
        created_at__date__lte=date_to
    ).select_related('menu_item', 'waiter', 'created_by').order_by('-created_at')
    
    expenses = Expense.objects.filter(
        date__gte=date_from,
        date__lte=date_to
    ).select_related('recorded_by').order_by('date', '-created_at')
    
    # Create workbook
    wb = Workbook()
    
    # Styling
    primary_fill = PatternFill(start_color="540b02", end_color="540b02", fill_type="solid")
    secondary_fill = PatternFill(start_color="00b652", end_color="00b652", fill_type="solid")
    accent_fill = PatternFill(start_color="f7941d", end_color="f7941d", fill_type="solid")
    danger_fill = PatternFill(start_color="dc3545", end_color="dc3545", fill_type="solid")
    success_fill = PatternFill(start_color="28a745", end_color="28a745", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # ============================================================
    # SHEET 1: ORDERS
    # ============================================================
    ws = wb.active
    ws.title = "Orders"
    
    ws.merge_cells('A1:K1')
    title_cell = ws['A1']
    title_cell.value = f"DEEM Restaurant - Orders Report ({label})"
    title_cell.font = Font(size=14, bold=True, color="FFFFFF")
    title_cell.fill = primary_fill
    title_cell.alignment = Alignment(horizontal='center')
    
    headers = ['Order #', 'Menu Items', 'Total Items', 'Total Qty', 'Total Amount (TZS)', 
               'Waiter', 'Location', 'Status', 'Notes', 'Created By', 'Time']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col, value=header)
        cell.font = white_font
        cell.fill = secondary_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center')
    
    session_groups = defaultdict(list)
    for order in orders:
        session_key = order.session_id if order.session_id else f"single_{order.id}"
        session_groups[session_key].append(order)
    
    row = 4
    order_num = 1
    grand_total = 0
    
    for session_key, group_orders in session_groups.items():
        group_orders.sort(key=lambda x: x.created_at)
        
        combined = OD()
        for o in group_orders:
            key = o.menu_item.name
            if key in combined:
                combined[key] += o.quantity
            else:
                combined[key] = o.quantity
        
        menu_items_str = ", ".join([f"{name} x{qty}" if qty > 1 else name for name, qty in combined.items()])
        
        total_qty = sum(o.quantity for o in group_orders)
        total_amount = sum(float(o.total_price) for o in group_orders)
        grand_total += total_amount
        
        first_order = group_orders[0]
        
        all_closed = all(o.status == 'closed' for o in group_orders)
        any_closed = any(o.status == 'closed' for o in group_orders)
        if all_closed:
            status = 'Closed'
        elif any_closed:
            status = 'Mixed'
        else:
            status = 'Open'
        
        if first_order.location_type == 'other':
            location = first_order.location_detail or 'Other'
        elif first_order.location_detail:
            location = f"{first_order.get_location_type_display()} #{first_order.location_detail}"
        else:
            location = first_order.get_location_type_display()
        
        waiter_name = first_order.waiter.display_name if first_order.waiter else 'N/A'
        
        ws.cell(row=row, column=1, value=f"#{order_num}").border = border
        ws.cell(row=row, column=2, value=menu_items_str).border = border
        ws.cell(row=row, column=3, value=len(combined)).border = border
        ws.cell(row=row, column=4, value=total_qty).border = border
        ws.cell(row=row, column=5, value=total_amount).border = border
        ws.cell(row=row, column=6, value=waiter_name).border = border
        ws.cell(row=row, column=7, value=location).border = border
        ws.cell(row=row, column=8, value=status).border = border
        ws.cell(row=row, column=9, value=first_order.notes or '').border = border
        ws.cell(row=row, column=10, value=first_order.created_by.username).border = border
        ws.cell(row=row, column=11, value=first_order.created_at.strftime('%H:%M')).border = border
        
        status_cell = ws.cell(row=row, column=8)
        if status == 'Closed':
            status_cell.fill = PatternFill(start_color="c6efce", end_color="c6efce", fill_type="solid")
            status_cell.font = Font(color="006100")
        elif status == 'Open':
            status_cell.fill = PatternFill(start_color="ffc7ce", end_color="ffc7ce", fill_type="solid")
            status_cell.font = Font(color="9c0006")
        else:
            status_cell.fill = PatternFill(start_color="ffeb9c", end_color="ffeb9c", fill_type="solid")
            status_cell.font = Font(color="9c6500")
        
        row += 1
        order_num += 1
    
    # Orders summary
    summary_row = row
    ws.merge_cells(f'A{summary_row}:D{summary_row}')
    summary_cell = ws.cell(row=summary_row, column=1, value="TOTAL SALES")
    summary_cell.font = Font(bold=True, size=12)
    summary_cell.fill = accent_fill
    summary_cell.border = border
    
    total_cell = ws.cell(row=summary_row, column=5, value=grand_total)
    total_cell.font = Font(bold=True, size=12, color="540b02")
    total_cell.fill = accent_fill
    total_cell.border = border
    
    for col in range(1, 12):
        ws.column_dimensions[get_column_letter(col)].width = 20
    ws.column_dimensions['B'].width = 45
    
    # ============================================================
    # SHEET 2: EXPENSES
    # ============================================================
    ws2 = wb.create_sheet("Expenses")
    
    ws2.merge_cells('A1:F1')
    title2 = ws2['A1']
    title2.value = f"DEEM Restaurant - Expenses Report ({label})"
    title2.font = Font(size=14, bold=True, color="FFFFFF")
    title2.fill = primary_fill
    title2.alignment = Alignment(horizontal='center')
    
    headers2 = ['Date', 'Category', 'Description', 'Receipt #', 'Amount (TZS)', 'Recorded By']
    for col, header in enumerate(headers2, 1):
        cell = ws2.cell(row=3, column=col, value=header)
        cell.font = white_font
        cell.fill = danger_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center')
    
    expense_total = 0
    expense_row = 4
    for expense in expenses:
        ws2.cell(row=expense_row, column=1, value=expense.date.strftime('%d/%m/%Y')).border = border
        ws2.cell(row=expense_row, column=2, value=expense.get_category_display()).border = border
        ws2.cell(row=expense_row, column=3, value=expense.description).border = border
        ws2.cell(row=expense_row, column=4, value=expense.receipt_ref or '-').border = border
        ws2.cell(row=expense_row, column=5, value=float(expense.amount)).border = border
        ws2.cell(row=expense_row, column=6, value=expense.recorded_by.username).border = border
        expense_total += float(expense.amount)
        expense_row += 1
    
    # Expenses summary
    if expenses:
        exp_summary_row = expense_row
        ws2.merge_cells(f'A{exp_summary_row}:D{exp_summary_row}')
        exp_summary = ws2.cell(row=exp_summary_row, column=1, value="TOTAL EXPENSES")
        exp_summary.font = Font(bold=True, size=12, color="FFFFFF")
        exp_summary.fill = danger_fill
        exp_summary.border = border
        
        exp_total = ws2.cell(row=exp_summary_row, column=5, value=expense_total)
        exp_total.font = Font(bold=True, size=12, color="FFFFFF")
        exp_total.fill = danger_fill
        exp_total.border = border
    
    for col in range(1, 7):
        ws2.column_dimensions[get_column_letter(col)].width = 22
    ws2.column_dimensions['C'].width = 40
    
    # ============================================================
    # SHEET 3: ITEM SALES + PROFIT/LOSS
    # ============================================================
    ws3 = wb.create_sheet("Item Sales")
    
    ws3.merge_cells('A1:F1')
    title3 = ws3['A1']
    title3.value = f"DEEM Restaurant - Item Sales Summary ({label})"
    title3.font = Font(size=14, bold=True, color="FFFFFF")
    title3.fill = primary_fill
    title3.alignment = Alignment(horizontal='center')
    
    headers3 = ['Menu Item', 'Category', 'Unit Price (TZS)', 'Qty Sold', 'Total Sales (TZS)', '% of Sales']
    for col, header in enumerate(headers3, 1):
        cell = ws3.cell(row=3, column=col, value=header)
        cell.font = white_font
        cell.fill = secondary_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center')
    
    # Aggregate item sales
    item_sales = {}
    for order in orders:
        key = order.menu_item.id
        if key not in item_sales:
            item_sales[key] = {
                'name': order.menu_item.name,
                'category': order.menu_item.category.name,
                'unit_price': float(order.menu_item.price),
                'qty': 0,
                'total': 0,
            }
        item_sales[key]['qty'] += order.quantity
        item_sales[key]['total'] += float(order.total_price)
    
    # Sort by total sales desc
    sorted_items = sorted(item_sales.values(), key=lambda x: x['total'], reverse=True)
    
    item_row = 4
    for item in sorted_items:
        percent = (item['total'] / grand_total * 100) if grand_total > 0 else 0
        ws3.cell(row=item_row, column=1, value=item['name']).border = border
        ws3.cell(row=item_row, column=2, value=item['category']).border = border
        ws3.cell(row=item_row, column=3, value=item['unit_price']).border = border
        ws3.cell(row=item_row, column=4, value=item['qty']).border = border
        ws3.cell(row=item_row, column=5, value=item['total']).border = border
        ws3.cell(row=item_row, column=6, value=f"{percent:.1f}%").border = border
        item_row += 1
    
    # Total row
    total_row = item_row
    ws3.merge_cells(f'A{total_row}:C{total_row}')
    ws3.cell(row=total_row, column=1, value="TOTAL").font = Font(bold=True)
    ws3.cell(row=total_row, column=1).fill = accent_fill
    ws3.cell(row=total_row, column=1).border = border
    ws3.cell(row=total_row, column=4, value=sum(i['qty'] for i in sorted_items)).font = Font(bold=True)
    ws3.cell(row=total_row, column=4).fill = accent_fill
    ws3.cell(row=total_row, column=4).border = border
    ws3.cell(row=total_row, column=5, value=grand_total).font = Font(bold=True)
    ws3.cell(row=total_row, column=5).fill = accent_fill
    ws3.cell(row=total_row, column=5).border = border
    ws3.cell(row=total_row, column=6, value="100%").font = Font(bold=True)
    ws3.cell(row=total_row, column=6).fill = accent_fill
    ws3.cell(row=total_row, column=6).border = border
    
    # ===== PROFIT / LOSS SECTION =====
    profit_start = total_row + 3
    
    ws3.merge_cells(f'A{profit_start}:F{profit_start}')
    pl_title = ws3.cell(row=profit_start, column=1, value="PROFIT / LOSS SUMMARY")
    pl_title.font = Font(size=13, bold=True, color="FFFFFF")
    pl_title.fill = primary_fill
    pl_title.alignment = Alignment(horizontal='center')
    pl_title.border = border
    
    # Total Sales
    ws3.merge_cells(f'A{profit_start+1}:D{profit_start+1}')
    ws3.cell(row=profit_start+1, column=1, value="Total Sales (from Orders sheet)").font = Font(bold=True)
    ws3.cell(row=profit_start+1, column=1).border = border
    ws3.cell(row=profit_start+1, column=1).fill = PatternFill(start_color="e8f5e9", end_color="e8f5e9", fill_type="solid")
    sales_cell = ws3.cell(row=profit_start+1, column=5, value=grand_total)
    sales_cell.font = Font(bold=True, color="006100")
    sales_cell.fill = PatternFill(start_color="e8f5e9", end_color="e8f5e9", fill_type="solid")
    sales_cell.border = border
    
    # Total Expenses
    ws3.merge_cells(f'A{profit_start+2}:D{profit_start+2}')
    ws3.cell(row=profit_start+2, column=1, value="Total Expenses (from Expenses sheet)").font = Font(bold=True)
    ws3.cell(row=profit_start+2, column=1).border = border
    ws3.cell(row=profit_start+2, column=1).fill = PatternFill(start_color="ffebee", end_color="ffebee", fill_type="solid")
    exp_cell = ws3.cell(row=profit_start+2, column=5, value=expense_total)
    exp_cell.font = Font(bold=True, color="9c0006")
    exp_cell.fill = PatternFill(start_color="ffebee", end_color="ffebee", fill_type="solid")
    exp_cell.border = border
    
    # Net Profit/Loss
    net = grand_total - expense_total
    ws3.merge_cells(f'A{profit_start+3}:D{profit_start+3}')
    net_label = ws3.cell(row=profit_start+3, column=1, value="NET PROFIT / LOSS")
    net_label.font = Font(bold=True, size=12, color="FFFFFF")
    net_label.fill = success_fill if net >= 0 else danger_fill
    net_label.border = border
    
    net_cell = ws3.cell(row=profit_start+3, column=5, value=net)
    net_cell.font = Font(bold=True, size=12, color="FFFFFF")
    net_cell.fill = success_fill if net >= 0 else danger_fill
    net_cell.border = border
    
    # Cash on Hand (same as net for this use case)
    ws3.merge_cells(f'A{profit_start+4}:D{profit_start+4}')
    cash_label = ws3.cell(row=profit_start+4, column=1, value="CASH ON HAND (Sales - Expenses)")
    cash_label.font = Font(bold=True, color="FFFFFF")
    cash_label.fill = accent_fill
    cash_label.border = border
    
    cash_cell = ws3.cell(row=profit_start+4, column=5, value=net)
    cash_cell.font = Font(bold=True, color="FFFFFF")
    cash_cell.fill = accent_fill
    cash_cell.border = border
    
    # Column widths
    for col in range(1, 7):
        ws3.column_dimensions[get_column_letter(col)].width = 22
    ws3.column_dimensions['A'].width = 30
    
    # ============================================================
    # 4th Sheet: Waiter Summary (Optional - existing)
    # ============================================================
    ws4 = wb.create_sheet("Waiter Summary")
    waiters_list = Waiter.objects.filter(is_active=True)
    
    ws4.merge_cells('A1:C1')
    title4 = ws4['A1']
    title4.value = f"Waiter Summary - {label}"
    title4.font = Font(size=14, bold=True, color="FFFFFF")
    title4.fill = primary_fill
    title4.alignment = Alignment(horizontal='center')
    
    headers4 = ['Waiter', 'Number of Orders', 'Total Collection (TZS)']
    for col, header in enumerate(headers4, 1):
        cell = ws4.cell(row=3, column=col, value=header)
        cell.font = white_font
        cell.fill = secondary_fill
        cell.border = border
    
    for r, waiter in enumerate(waiters_list, 4):
        waiter_orders = [o for o in orders if o.waiter == waiter]
        ws4.cell(row=r, column=1, value=waiter.full_name).border = border
        ws4.cell(row=r, column=2, value=len(waiter_orders)).border = border
        ws4.cell(row=r, column=3, value=sum(float(o.total_price) for o in waiter_orders)).border = border
    
    for col in range(1, 4):
        ws4.column_dimensions[get_column_letter(col)].width = 25
    
    # Response
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    filename = f"DEEM_Report_{date_from}_to_{date_to}.xlsx" if date_from != date_to else f"DEEM_Report_{date_from}.xlsx"
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    
    wb.save(response)
    return response

@login_required
def get_order_group_info(request, session_id):
    """Get information about an order group for adding items"""
    if request.method == 'GET':
        orders = Order.objects.filter(session_id=session_id)
        
        if not orders.exists():
            return JsonResponse({'error': 'Order not found'}, status=404)
        
        first_order = orders.first()
        
        # Get combined items info
        from collections import OrderedDict
        combined_items = OrderedDict()
        for order in orders:
            item_key = order.menu_item.id
            if item_key in combined_items:
                combined_items[item_key]['quantity'] += order.quantity
            else:
                combined_items[item_key] = {
                    'menu_item_id': order.menu_item.id,
                    'menu_item_name': order.menu_item.name,
                    'quantity': order.quantity,
                    'unit_price': float(order.unit_price),
                }
        
        data = {
            'session_id': session_id,
            'waiter_id': first_order.waiter.id if first_order.waiter else None,
            'waiter_name': first_order.waiter.display_name if first_order.waiter else '',
            'location_type': first_order.location_type,
            'location_detail': first_order.location_detail or '',
            'notes': first_order.notes or '',
            'created_at': first_order.created_at.strftime('%H:%M'),
            'total_items': orders.count(),
            'total_amount': sum(float(o.total_price) for o in orders),
            'existing_items': list(combined_items.values()),
        }
        
        return JsonResponse(data)

@login_required
def add_to_order(request, session_id):
    """Add new items to an existing order session"""
    if request.method == 'POST':
        try:
            with transaction.atomic():
                # Get the existing order to copy waiter, location, notes
                existing_order = Order.objects.filter(session_id=session_id).first()
                
                if not existing_order:
                    messages.error(request, 'Order not found!')
                    return redirect('home')
                
                # Process each menu item
                menu_items = MenuItem.objects.filter(is_available=True)
                items_added = 0
                
                for item in menu_items:
                    is_checked = request.POST.get(f'item_{item.id}')
                    quantity_str = request.POST.get(f'quantity_{item.id}', '0')
                    
                    if is_checked == 'on':
                        quantity = int(quantity_str) if quantity_str else 1
                        
                        if quantity > 0:
                            # Create separate entries for each quantity
                            for _ in range(quantity):
                                Order.objects.create(
                                    menu_item=item,
                                    quantity=1,
                                    unit_price=item.price,
                                    total_price=item.price,
                                    waiter=existing_order.waiter,
                                    location_type=existing_order.location_type,
                                    location_detail=existing_order.location_detail,
                                    notes=existing_order.notes,
                                    created_by=request.user,
                                    session_id=session_id  # Use same session_id
                                )
                                items_added += 1
                
                if items_added > 0:
                    messages.success(request, f'{items_added} items added to order successfully!')
                else:
                    messages.warning(request, 'No items selected to add.')
                
                return redirect('home')
                
        except Exception as e:
            messages.error(request, f'Error adding items: {str(e)}')
            return redirect('home')
    
    return redirect('home')

@login_required
def toggle_group_status(request, session_id):
    """Toggle Open/Closed status for entire order group"""
    if request.method == 'POST':
        orders = Order.objects.filter(session_id=session_id)
        if orders.exists():
            all_closed = all(order.status == 'closed' for order in orders)
            
            if all_closed:
                orders.update(status='open', closed_at=None)
                status_msg = f'Order Reopened! ({orders.count()} items)'
            else:
                orders.update(status='closed', closed_at=timezone.now())
                status_msg = f'Order Closed! ({orders.count()} items)'
            
            # Check if it's an AJAX request
            if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                return JsonResponse({'success': True, 'message': status_msg})
            
            messages.success(request, status_msg)
    
    return redirect('home')

from django.contrib.auth.models import User
from .models import UserProfile
import re


def is_admin(user):
    return user.is_staff or user.is_superuser


def user_management(request):
    """User management dashboard - Admin only"""
    users = User.objects.all().select_related('profile').order_by('-date_joined')
    
    users_data = []
    for u in users:
        try:
            profile = u.profile
        except UserProfile.DoesNotExist:
            profile = UserProfile.objects.create(user=u)
        
        users_data.append({
            'user': u,
            'profile': profile,
        })
    
    context = {
        'users_data': users_data,
        'total_users': users.count(),
        'admin_count': UserProfile.objects.filter(role='admin').count(),
        'manager_count': UserProfile.objects.filter(role='manager').count(),
        'cashier_count': UserProfile.objects.filter(role='cashier').count(),
        'active_count': User.objects.filter(is_active=True).count(),
    }
    return render(request, 'sales/user_management.html', context)

@admin_required
def add_user(request):
    """Add a new user - Admin only"""
    if request.method == 'POST':
        username = request.POST.get('username', '').strip()
        email = request.POST.get('email', '').strip()
        password = request.POST.get('password', '')
        first_name = request.POST.get('first_name', '').strip()
        last_name = request.POST.get('last_name', '').strip()
        role = request.POST.get('role', 'cashier')
        phone = request.POST.get('phone', '').strip()
        
        if not username or not password:
            messages.error(request, 'Username and password are required.')
            return redirect('user_management')
        
        if User.objects.filter(username=username).exists():
            messages.error(request, f'Username "{username}" already exists.')
            return redirect('user_management')
        
        if len(password) < 6:
            messages.error(request, 'Password must be at least 6 characters.')
            return redirect('user_management')
        
        try:
            with transaction.atomic():
                user = User.objects.create_user(
                    username=username,
                    email=email,
                    password=password,
                    first_name=first_name,
                    last_name=last_name,
                )
                
                # Admin role gets staff access (for Django admin)
                if role == 'admin':
                    user.is_staff = True
                
                user.save()
                
                # Update profile
                profile = user.profile
                profile.role = role
                profile.phone = phone
                profile.save()
                
                messages.success(request, f'User "{username}" created as {role.title()}!')
        except Exception as e:
            messages.error(request, f'Error creating user: {str(e)}')
    
    return redirect('user_management')


@admin_required
def edit_user(request, user_id):
    """Edit existing user - Admin only"""
    if request.method == 'POST':
        user = get_object_or_404(User, id=user_id)
        
        # Prevent editing yourself as non-admin
        if user == request.user and request.POST.get('role') != 'admin':
            messages.error(request, 'You cannot change your own role.')
            return redirect('user_management')
        
        try:
            with transaction.atomic():
                user.email = request.POST.get('email', '').strip()
                user.first_name = request.POST.get('first_name', '').strip()
                user.last_name = request.POST.get('last_name', '').strip()
                
                role = request.POST.get('role', 'cashier')
                is_active = request.POST.get('is_active') == 'on'
                
                # Set staff status only for admins
                if role == 'admin':
                    user.is_staff = True
                else:
                    user.is_staff = False
                
                user.is_active = is_active
                # Capture BEFORE state
                from .audit import serialize_instance, log_action
                data_before = serialize_instance(user)
                user.save()
                
                # Update profile
                profile = user.profile
                profile.role = role
                profile.phone = request.POST.get('phone', '').strip()
                profile.is_active = is_active
                profile.save()
                
                new_password = request.POST.get('new_password', '')
                if new_password and len(new_password) >= 6:
                    user.set_password(new_password)
                    user.save()
                    log_action(request, 'update', user, data_before=data_before,
                    notes=f'Role: {role}, Active: {is_active}')
                    messages.success(request, f'User "{user.username}" updated with new password.')
                else:
                    messages.success(request, f'User "{user.username}" updated successfully!')
        except Exception as e:
            messages.error(request, f'Error updating user: {str(e)}')
    
    return redirect('user_management')

@admin_required
def toggle_user_status(request, user_id):
    """Activate/deactivate a user - Admin only"""
    if request.method == 'POST':
        user = get_object_or_404(User, id=user_id)
        
        if user == request.user:
            messages.error(request, 'You cannot deactivate your own account.')
            return redirect('user_management')
        
        user.is_active = not user.is_active
        user.save()
        
        try:
            user.profile.is_active = user.is_active
            user.profile.save()
        except UserProfile.DoesNotExist:
            pass
        
        status = "activated" if user.is_active else "deactivated"
        messages.success(request, f'User "{user.username}" {status}.')
    
    return redirect('user_management')


@admin_required
def delete_user(request, user_id):
    """Delete a user - Admin only"""
    if request.method == 'POST':
        user = get_object_or_404(User, id=user_id)
        
        if user == request.user:
            messages.error(request, 'You cannot delete your own account.')
            return redirect('user_management')
            
        from .audit import serialize_instance, log_action
        data_before = serialize_instance(user)
        log_action(request, 'delete', user, data_before=data_before)
        
        username = user.username
        
        user.delete()
        messages.success(request, f'User "{username}" deleted successfully.')
    
    return redirect('user_management')


@admin_required
def reset_password(request, user_id):
    """Reset a user's password - Admin only"""
    if request.method == 'POST':
        user = get_object_or_404(User, id=user_id)
        new_password = request.POST.get('new_password', '')
        
        if len(new_password) < 6:
            messages.error(request, 'Password must be at least 6 characters.')
            return redirect('user_management')
        
        user.set_password(new_password)
        user.save()
        messages.success(request, f'Password reset for "{user.username}".')
    
    return redirect('user_management')

# ============================================================
# MENU MANAGEMENT (Categories, Items, Waiters)
# ============================================================

@reports_required
def menu_management(request):
    """Menu management dashboard - Admin and Manager"""
    categories = MenuCategory.objects.all().prefetch_related('items')
    menu_items = MenuItem.objects.all().select_related('category').order_by('category__display_order', 'display_order', 'name')
    waiters = Waiter.objects.all().order_by('display_name')
    
    context = {
        'categories': categories,
        'menu_items': menu_items,
        'waiters': waiters,
        'total_categories': categories.count(),
        'total_items': menu_items.count(),
        'total_waiters': waiters.count(),
        'available_items': menu_items.filter(is_available=True).count(),
        'active_waiters': waiters.filter(is_active=True).count(),
    }
    return render(request, 'sales/menu_management.html', context)


# ========== CATEGORIES ==========

@reports_required
def add_category(request):
    """Add a new menu category"""
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        display_order = request.POST.get('display_order', 0)
        
        if not name:
            messages.error(request, 'Category name is required.')
            return redirect('menu_management')
        
        if MenuCategory.objects.filter(name__iexact=name).exists():
            messages.error(request, f'Category "{name}" already exists.')
            return redirect('menu_management')
        
        try:
            MenuCategory.objects.create(
                name=name,
                display_order=int(display_order) if display_order else 0
            )
            messages.success(request, f'Category "{name}" created successfully!')
        except Exception as e:
            messages.error(request, f'Error creating category: {str(e)}')
    
    return redirect('menu_management')


@reports_required
def edit_category(request, category_id):
    """Edit a category"""
    if request.method == 'POST':
        category = get_object_or_404(MenuCategory, id=category_id)
        name = request.POST.get('name', '').strip()
        display_order = request.POST.get('display_order', 0)
        
        if not name:
            messages.error(request, 'Category name is required.')
            return redirect('menu_management')
        
        try:
            category.name = name
            category.display_order = int(display_order) if display_order else 0
            category.save()
            messages.success(request, f'Category updated successfully!')
        except Exception as e:
            messages.error(request, f'Error updating category: {str(e)}')
    
    return redirect('menu_management')


@reports_required
def delete_category(request, category_id):
    """Delete a category"""
    if request.method == 'POST':
        category = get_object_or_404(MenuCategory, id=category_id)
        
        # Check if it has items
        if category.items.exists():
            messages.error(request, f'Cannot delete "{category.name}" — it has {category.items.count()} menu items. Delete or move them first.')
            return redirect('menu_management')
        
        name = category.name
        category.delete()
        messages.success(request, f'Category "{name}" deleted.')
    
    return redirect('menu_management')


# ========== MENU ITEMS ==========

@reports_required
def add_menu_item(request):
    """Add a new menu item"""
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        category_id = request.POST.get('category')
        price = request.POST.get('price', '0')
        display_order = request.POST.get('display_order', 0)
        is_available = request.POST.get('is_available') == 'on'
        
        if not name or not category_id:
            messages.error(request, 'Menu item name and category are required.')
            return redirect('menu_management')
        
        try:
            category = MenuCategory.objects.get(id=category_id)
            MenuItem.objects.create(
                name=name,
                category=category,
                price=float(price),
                display_order=int(display_order) if display_order else 0,
                is_available=is_available
            )
            messages.success(request, f'Menu item "{name}" added successfully!')
        except MenuCategory.DoesNotExist:
            messages.error(request, 'Selected category not found.')
        except Exception as e:
            messages.error(request, f'Error adding menu item: {str(e)}')
    
    return redirect('menu_management')


@reports_required
def edit_menu_item(request, item_id):
    """Edit a menu item (with audit logging)"""
    if request.method == 'POST':
        item = get_object_or_404(MenuItem, id=item_id)
        try:
            from .audit import serialize_instance, log_action
            data_before = serialize_instance(item)
            
            item.name = request.POST.get('name', '').strip()
            item.category = MenuCategory.objects.get(id=request.POST.get('category'))
            item.price = float(request.POST.get('price', '0'))
            item.display_order = int(request.POST.get('display_order', 0) or 0)
            item.is_available = request.POST.get('is_available') == 'on'
            item.save()
            
            log_action(request, 'update', item, data_before=data_before)
            
            messages.success(request, f'Menu item updated successfully!')
        except Exception as e:
            messages.error(request, f'Error updating menu item: {str(e)}')
    
    return redirect('menu_management')


@reports_required
def delete_menu_item(request, item_id):
    """Delete a menu item (with audit logging)"""
    if request.method == 'POST':
        item = get_object_or_404(MenuItem, id=item_id)
        
        from .audit import serialize_instance, log_action
        data_before = serialize_instance(item)
        log_action(request, 'delete', item, data_before=data_before)
        
        name = item.name
        item.delete()
        messages.success(request, f'Menu item "{name}" deleted.')
    
    return redirect('menu_management')


@reports_required
def toggle_item_availability(request, item_id):
    """Toggle item availability"""
    if request.method == 'POST':
        item = get_object_or_404(MenuItem, id=item_id)
        item.is_available = not item.is_available
        item.save()
        status = "available" if item.is_available else "unavailable"
        messages.success(request, f'"{item.name}" marked as {status}.')
    
    return redirect('menu_management')


# ========== WAITERS ==========

@reports_required
def add_waiter(request):
    """Add a new waiter"""
    if request.method == 'POST':
        full_name = request.POST.get('full_name', '').strip()
        display_name = request.POST.get('display_name', '').strip()
        is_active = request.POST.get('is_active') == 'on'
        
        if not full_name or not display_name:
            messages.error(request, 'Full name and display name are required.')
            return redirect('menu_management')
        
        if Waiter.objects.filter(full_name__iexact=full_name).exists():
            messages.error(request, f'Waiter "{full_name}" already exists.')
            return redirect('menu_management')
        
        try:
            Waiter.objects.create(
                full_name=full_name,
                display_name=display_name,
                is_active=is_active
            )
            messages.success(request, f'Waiter "{display_name}" added successfully!')
        except Exception as e:
            messages.error(request, f'Error adding waiter: {str(e)}')
    
    return redirect('menu_management')


@reports_required
def edit_waiter(request, waiter_id):
    """Edit a waiter (with audit logging)"""
    if request.method == 'POST':
        waiter = get_object_or_404(Waiter, id=waiter_id)
        try:
            from .audit import serialize_instance, log_action
            data_before = serialize_instance(waiter)
            
            waiter.full_name = request.POST.get('full_name', '').strip()
            waiter.display_name = request.POST.get('display_name', '').strip()
            waiter.is_active = request.POST.get('is_active') == 'on'
            waiter.save()
            
            log_action(request, 'update', waiter, data_before=data_before)
            
            messages.success(request, f'Waiter updated successfully!')
        except Exception as e:
            messages.error(request, f'Error updating waiter: {str(e)}')
    
    return redirect('menu_management')


@reports_required
def delete_waiter(request, waiter_id):
    """Delete a waiter (with audit logging)"""
    if request.method == 'POST':
        waiter = get_object_or_404(Waiter, id=waiter_id)
        
        from .audit import serialize_instance, log_action
        data_before = serialize_instance(waiter)
        log_action(request, 'delete', waiter, data_before=data_before)
        
        name = waiter.display_name
        waiter.delete()
        messages.success(request, f'Waiter "{name}" deleted.')
    
    return redirect('menu_management')


@reports_required
def toggle_waiter_status(request, waiter_id):
    """Toggle waiter active status"""
    if request.method == 'POST':
        waiter = get_object_or_404(Waiter, id=waiter_id)
        waiter.is_active = not waiter.is_active
        waiter.save()
        status = "activated" if waiter.is_active else "deactivated"
        messages.success(request, f'Waiter "{waiter.display_name}" {status}.')
    
    return redirect('menu_management')

from .models import Expense
from datetime import date, timedelta

@login_required
def expenses_view(request):
    """View and add expenses - all roles can record"""
    today = timezone.now().date()
    selected_date = request.GET.get('date', today.strftime('%Y-%m-%d'))
    
    try:
        selected_date = datetime.strptime(selected_date, '%Y-%m-%d').date()
    except:
        selected_date = today
    
    expenses = Expense.objects.filter(date=selected_date).select_related('recorded_by')
    
    # Daily totals
    today_expenses = Expense.objects.filter(date=today)
    today_expense_total = sum(float(e.amount) for e in today_expenses)
    
    # Today's sales
    today_orders = Order.objects.filter(created_at__date=today)
    today_sales = sum(float(o.total_price) for o in today_orders)
    
    context = {
        'expenses': expenses,
        'selected_date': selected_date,
        'categories': Expense.CATEGORY_CHOICES,
        'today': today,
        'selected_total': sum(float(e.amount) for e in expenses),
        'today_expense_total': today_expense_total,
        'today_sales': today_sales,
        'today_balance': today_sales - today_expense_total,
    }
    return render(request, 'sales/expenses.html', context)


@login_required
def add_expense(request):
    """Add new expense"""
    if request.method == 'POST':
        try:
            category = request.POST.get('category', 'other')
            amount = request.POST.get('amount', '0')
            description = request.POST.get('description', '').strip()
            date_str = request.POST.get('date', '')
            receipt_ref = request.POST.get('receipt_ref', '').strip()
            
            if not description:
                messages.error(request, 'Description is required.')
                return redirect('expenses_view')
            
            expense_date = datetime.strptime(date_str, '%Y-%m-%d').date() if date_str else timezone.now().date()
            
            Expense.objects.create(
                category=category,
                amount=float(amount),
                description=description,
                date=expense_date,
                receipt_ref=receipt_ref if receipt_ref else None,
                recorded_by=request.user
            )
            messages.success(request, f'Expense of {float(amount):,.0f} TZS recorded successfully!')
        except Exception as e:
            messages.error(request, f'Error recording expense: {str(e)}')
    
    return redirect('expenses_view')


@login_required
def edit_expense(request, expense_id):
    """Edit expense (with audit logging)"""
    if request.method == 'POST':
        expense = get_object_or_404(Expense, id=expense_id)
        try:
            from .audit import serialize_instance, log_action
            data_before = serialize_instance(expense)
            
            expense.category = request.POST.get('category', 'other')
            expense.amount = float(request.POST.get('amount', '0'))
            expense.description = request.POST.get('description', '').strip()
            date_str = request.POST.get('date', '')
            if date_str:
                expense.date = datetime.strptime(date_str, '%Y-%m-%d').date()
            expense.receipt_ref = request.POST.get('receipt_ref', '').strip() or None
            expense.save()
            
            log_action(request, 'update', expense, data_before=data_before)
            
            messages.success(request, 'Expense updated successfully!')
        except Exception as e:
            messages.error(request, f'Error updating expense: {str(e)}')
    
    return redirect('expenses_view')


@login_required
def delete_expense(request, expense_id):
    """Delete expense (with audit logging)"""
    if request.method == 'POST':
        expense = get_object_or_404(Expense, id=expense_id)
        
        from .audit import serialize_instance, log_action
        data_before = serialize_instance(expense)
        log_action(request, 'delete', expense, data_before=data_before)
        
        expense.delete()
        messages.success(request, 'Expense deleted.')
    
    return redirect('expenses_view')


@login_required
def get_expense(request, expense_id):
    """Get expense data for editing"""
    if request.method == 'GET':
        expense = get_object_or_404(Expense, id=expense_id)
        data = {
            'id': expense.id,
            'category': expense.category,
            'amount': float(expense.amount),
            'description': expense.description,
            'date': expense.date.strftime('%Y-%m-%d'),
            'receipt_ref': expense.receipt_ref or '',
        }
        return JsonResponse(data)
    
@admin_required
def audit_log_view(request):
    """View all audit logs - Admin only"""
    from .models import AuditLog
    
    # Filters
    action_filter = request.GET.get('action', '')
    model_filter = request.GET.get('model', '')
    user_filter = request.GET.get('user', '')
    date_filter = request.GET.get('date', '')
    
    logs = AuditLog.objects.all().select_related('performed_by')
    
    if action_filter:
        logs = logs.filter(action=action_filter)
    if model_filter:
        logs = logs.filter(model_name=model_filter)
    if user_filter:
        logs = logs.filter(performed_by__username=user_filter)
    if date_filter:
        try:
            filter_date = datetime.strptime(date_filter, '%Y-%m-%d').date()
            logs = logs.filter(performed_at__date=filter_date)
        except:
            pass
    
    logs = logs[:500]  # Limit to latest 500
    
    # Stats
    total_logs = AuditLog.objects.count()
    update_count = AuditLog.objects.filter(action='update').count()
    delete_count = AuditLog.objects.filter(action='delete').count()
    today_logs = AuditLog.objects.filter(performed_at__date=timezone.now().date()).count()
    
    context = {
        'logs': logs,
        'total_logs': total_logs,
        'update_count': update_count,
        'delete_count': delete_count,
        'today_logs': today_logs,
        'action_filter': action_filter,
        'model_filter': model_filter,
        'user_filter': user_filter,
        'date_filter': date_filter,
        'today': timezone.now().date(),
    }
    return render(request, 'sales/audit_log.html', context)

@admin_required
def audit_log_detail(request, log_id):
    """Get audit log details as JSON"""
    from .models import AuditLog
    log = get_object_or_404(AuditLog, id=log_id)
    
    return JsonResponse({
        'id': log.id,
        'action': log.get_action_display(),
        'model_name': log.model_name,
        'object_id': log.object_id,
        'object_repr': log.object_repr,
        'performed_at': log.performed_at.strftime('%d/%m/%Y %H:%M:%S'),
        'performed_by': log.performed_by.username if log.performed_by else 'system',
        'ip_address': log.ip_address or '',
        'notes': log.notes or '',
        'data_before': log.data_before,
        'data_after': log.data_after,
        'changes': log.get_changed_fields(),
    })
