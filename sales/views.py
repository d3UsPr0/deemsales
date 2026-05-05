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


def is_admin(user):
    return user.is_staff or user.is_superuser


def login_view(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        password = request.POST.get('password')
        user = authenticate(request, username=username, password=password)
        
        if user is not None:
            login(request, user)
            messages.success(request, f'Welcome, {user.username}!')
            return redirect('home')
        else:
            messages.error(request, 'Invalid credentials')
    
    return render(request, 'sales/login.html')


def logout_view(request):
    logout(request)
    messages.info(request, 'You have been logged out')
    return redirect('login')


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
    from collections import defaultdict, OrderedDict
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

    context = {
        'categories': categories,
        'menu_items': menu_items,
        'waiters': waiters,
        'grouped_orders': grouped_orders,
        'total_collection': total_collection,
        'today': today,
        'total_orders_count': today_orders.count(),
        'total_sessions': len(grouped_orders),
        'open_count': open_count,
        'closed_count': closed_count,
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
    """Update an existing order"""
    if request.method == 'POST':
        try:
            order = get_object_or_404(Order, id=order_id)
            
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
            
            messages.success(request, 'Order updated successfully!')
            return redirect('home')
            
        except Exception as e:
            messages.error(request, f'Error updating order: {str(e)}')
            return redirect('home')
    
    return redirect('home')

@login_required
def delete_order(request, order_id):
    """Delete an order"""
    if request.method == 'POST':
        order = get_object_or_404(Order, id=order_id)
        order.delete()
        messages.success(request, 'Order deleted successfully!')
    return redirect('home')

@login_required
def delete_group_orders(request, session_id):
    """Delete all orders in a group session"""
    if request.method == 'POST':
        orders = Order.objects.filter(session_id=session_id)
        count = orders.count()
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

@login_required
@user_passes_test(is_admin)
def export_excel(request):
    """Export orders to Excel"""
    # Get date from query params or use today
    date_str = request.GET.get('date')
    if date_str:
        export_date = datetime.strptime(date_str, '%Y-%m-%d').date()
    else:
        export_date = timezone.now().date()
    
    # Get orders for the selected date
    orders = Order.objects.filter(
        created_at__date=export_date
    ).select_related('menu_item', 'waiter', 'created_by').order_by('-created_at')
    
    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = f"Orders - {export_date}"
    
    # DEEM Restaurant colors
    primary_fill = PatternFill(start_color="540b02", end_color="540b02", fill_type="solid")
    secondary_fill = PatternFill(start_color="00b652", end_color="00b652", fill_type="solid")
    accent_fill = PatternFill(start_color="f7941d", end_color="f7941d", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Title
    ws.merge_cells('A1:H1')
    title_cell = ws['A1']
    title_cell.value = f"DEEM Restaurant - Daily Orders Report ({export_date})"
    title_cell.font = Font(size=14, bold=True, color="FFFFFF")
    title_cell.fill = primary_fill
    title_cell.alignment = Alignment(horizontal='center')
    
    # Headers
    headers = ['Order #', 'Menu Items', 'Total Items', 'Total Qty', 'Total Amount (TZS)', 
               'Waiter', 'Location', 'Status', 'Notes', 'Created By', 'Time']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col, value=header)
        cell.font = white_font
        cell.fill = secondary_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center')
    
        # Group orders by session_id for combined display
    from collections import defaultdict
    session_groups = defaultdict(list)
    for order in orders:
        session_key = order.session_id if order.session_id else f"single_{order.id}"
        session_groups[session_key].append(order)
    
    # Data - grouped by order session
    row = 4
    order_num = 1
    grand_total = 0
    
    for session_key, group_orders in session_groups.items():
        # Sort orders within group
        group_orders.sort(key=lambda x: x.created_at)
        
        # Combine menu items into one string
        from collections import OrderedDict as OD
        combined = OD()
        for o in group_orders:
            key = o.menu_item.name
            if key in combined:
                combined[key] += o.quantity
            else:
                combined[key] = o.quantity
        
        menu_items_str = ", ".join([f"{name} ×{qty}" if qty > 1 else name for name, qty in combined.items()])
        
        total_qty = sum(o.quantity for o in group_orders)
        total_amount = sum(float(o.total_price) for o in group_orders)
        grand_total += total_amount
        
        first_order = group_orders[0]
        
        # Status
        all_closed = all(o.status == 'closed' for o in group_orders)
        any_closed = any(o.status == 'closed' for o in group_orders)
        if all_closed:
            status = 'Closed'
        elif any_closed:
            status = 'Mixed'
        else:
            status = 'Open'
        
        # Location
        if first_order.location_type == 'other':
            location = first_order.location_detail or 'Other'
        elif first_order.location_detail:
            location = f"{first_order.get_location_type_display()} #{first_order.location_detail}"
        else:
            location = first_order.get_location_type_display()
        
        # Waiter
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
        
        # Color the status cell
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
    
    # Summary row
    summary_row = row
    ws.merge_cells(f'A{summary_row}:D{summary_row}')
    summary_cell = ws.cell(row=summary_row, column=1, value="TOTAL COLLECTION")
    summary_cell.font = Font(bold=True, size=12)
    summary_cell.fill = accent_fill
    summary_cell.border = border
    
    total_cell = ws.cell(row=summary_row, column=5, value=grand_total)
    total_cell.font = Font(bold=True, size=12, color="540b02")
    total_cell.fill = accent_fill
    total_cell.border = border
    
    # Second sheet with summary by waiter
    ws2 = wb.create_sheet("Waiter Summary")
    waiters = Waiter.objects.filter(is_active=True)
    
    ws2.merge_cells('A1:D1')
    title2 = ws2['A1']
    title2.value = f"Waiter Collection Summary - {export_date}"
    title2.font = Font(size=14, bold=True, color="FFFFFF")
    title2.fill = primary_fill
    title2.alignment = Alignment(horizontal='center')
    
    headers2 = ['Waiter', 'Number of Orders', 'Total Collection (TZS)']
    for col, header in enumerate(headers2, 1):
        cell = ws2.cell(row=3, column=col, value=header)
        cell.font = white_font
        cell.fill = secondary_fill
        cell.border = border
    
    for row, waiter in enumerate(waiters, 4):
        waiter_orders = orders.filter(waiter=waiter)
        ws2.cell(row=row, column=1, value=waiter.full_name).border = border
        ws2.cell(row=row, column=2, value=waiter_orders.count()).border = border
        ws2.cell(row=row, column=3, value=sum(float(o.total_price) for o in waiter_orders)).border = border
    
    # Adjust column widths
    for ws_obj in [ws, ws2]:
        for col in range(1, 12):
            ws_obj.column_dimensions[get_column_letter(col)].width = 20
    # Make menu items column wider
    ws.column_dimensions['B'].width = 45
    
    # Create response
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="DEEM_Orders_{export_date}.xlsx"'
    
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
