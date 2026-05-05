from django import forms
from .models import Order, MenuItem, Waiter


class OrderForm(forms.Form):
    """Form for creating orders"""
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        
        # Get all available menu items grouped by category
        from .models import MenuCategory, MenuItem, Waiter
        
        categories = MenuCategory.objects.all()
        
        for category in categories:
            items = MenuItem.objects.filter(category=category, is_available=True)
            choices = [(item.id, f"{item.name} ({item.price} TZS)") for item in items]
            
            if choices:
                self.fields[f'category_{category.id}'] = forms.MultipleChoiceField(
                    choices=choices,
                    widget=forms.CheckboxSelectMultiple,
                    required=False,
                    label=category.name
                )
        
        # Add quantity fields for each menu item
        for item in MenuItem.objects.filter(is_available=True):
            self.fields[f'quantity_{item.id}'] = forms.IntegerField(
                min_value=0,
                initial=0,
                required=False,
                widget=forms.NumberInput(attrs={'class': 'quantity-input', 'min': 0})
            )
        
        # Waiter selection
        waiters = Waiter.objects.filter(is_active=True)
        self.fields['waiter'] = forms.ChoiceField(
            choices=[('', 'Select Waiter')] + [(w.id, w.display_name) for w in waiters],
            required=True,
            widget=forms.Select(attrs={'class': 'waiter-select'})
        )
        
        # Notes
        self.fields['notes'] = forms.CharField(
            widget=forms.Textarea(attrs={'rows': 3, 'placeholder': 'Any additional information...'}),
            required=False
        )