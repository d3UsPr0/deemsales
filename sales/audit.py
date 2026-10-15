"""Audit logging helpers — capture UPDATE and DELETE actions"""

from django.utils import timezone
from decimal import Decimal
from datetime import datetime, date


def serialize_instance(instance):
    """Convert a model instance to a JSON-friendly dict"""
    data = {}
    for field in instance._meta.fields:
        value = getattr(instance, field.name)
        
        # Handle special types
        if isinstance(value, Decimal):
            value = float(value)
        elif isinstance(value, datetime):
            value = value.strftime('%Y-%m-%d %H:%M:%S')
        elif isinstance(value, date):
            value = value.strftime('%Y-%m-%d')
        elif hasattr(value, 'pk'):  # ForeignKey
            value = str(value) if value else None
        elif value is None:
            value = None
        else:
            value = str(value) if not isinstance(value, (int, float, bool, str, list, dict)) else value
        
        data[field.name] = value
    return data


def get_client_ip(request):
    """Get client IP address from request"""
    x_forwarded_for = request.META.get('HTTP_X_FORWARDED_FOR')
    if x_forwarded_for:
        ip = x_forwarded_for.split(',')[0]
    else:
        ip = request.META.get('REMOTE_ADDR')
    return ip


def log_action(request, action, instance, data_before=None, notes=None):
    """
    Log an UPDATE or DELETE action
    
    Args:
        request: HTTP request object
        action: 'update' or 'delete'
        instance: the model instance being modified
        data_before: dict snapshot before modification (for update)
        notes: optional extra notes
    """
    from .models import AuditLog
    
    data_after = None
    if action == 'update':
        data_after = serialize_instance(instance)
    
    if data_before is None:
        data_before = {}
    
    AuditLog.objects.create(
        action=action,
        model_name=instance.__class__.__name__,
        object_id=str(instance.pk),
        object_repr=str(instance)[:255],
        data_before=data_before if data_before else None,
        data_after=data_after,
        performed_by=request.user if request.user.is_authenticated else None,
        ip_address=get_client_ip(request),
        notes=notes,
    )