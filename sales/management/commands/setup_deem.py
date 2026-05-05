from django.core.management.base import BaseCommand
from django.contrib.auth.models import User, Group, Permission
from sales.models import MenuCategory, MenuItem, Waiter


class Command(BaseCommand):
    help = 'Setup DEEM Restaurant initial data'
    
    def handle(self, *args, **kwargs):
        # Create groups
        admin_group, _ = Group.objects.get_or_create(name='Admin')
        manager_group, _ = Group.objects.get_or_create(name='Order Manager')
        
        # Create users
        admin_user, created = User.objects.get_or_create(username='admin', 
                            defaults={'is_staff': True, 'is_superuser': True})
        if created:
            admin_user.set_password('deemadmin2024')
            admin_user.save()
            self.stdout.write('Admin user created: admin / deemadmin2024')
        
        manager_user, created = User.objects.get_or_create(username='manager')
        if created:
            manager_user.set_password('deemorder2024')
            manager_user.save()
            self.stdout.write('Order Manager created: manager / deemorder2024')
        
        # Create menu categories
        categories = [
            ('Breakfast', 1),
            ('Main', 2),
            ('Drinks', 3),
            ('Choma', 4),
        ]
        
        for cat_name, order in categories:
            MenuCategory.objects.get_or_create(name=cat_name, defaults={'display_order': order})
        
        # Create sample menu items
        menu_items = [
            ('Breakfast', 'Chai', 500),
            ('Breakfast', 'Chapati', 1000),
            ('Breakfast', 'Andazi', 500),
            ('Main', 'Ugali Nyama', 5000),
            ('Main', 'Wali Samaki', 5000),
            ('Main', 'Chips Kuku', 6000),
            ('Drinks', 'Soda', 1500),
            ('Drinks', 'Maji', 1000),
            ('Drinks', 'Juice', 3000),
            ('Choma', 'Nyama Choma 1/4', 5000),
            ('Choma', 'Nyama Choma 1/2', 8000),
            ('Choma', 'Kuku Choma', 10000),
        ]
        
        for cat_name, item_name, price in menu_items:
            category = MenuCategory.objects.get(name=cat_name)
            MenuItem.objects.get_or_create(
                name=item_name, 
                category=category,
                defaults={'price': price}
            )
        
        # Create sample waiters
        waiters = [
            ('Benedicto Joseph', 'Ben'),
            ('Maria Charles', 'Maria'),
            ('John Michael', 'JM'),
            ('Anna Petro', 'Anna'),
        ]
        
        for full_name, display_name in waiters:
            Waiter.objects.get_or_create(
                full_name=full_name,
                defaults={'display_name': display_name}
            )
        
        self.stdout.write(self.style.SUCCESS('Setup completed successfully!'))