from django.core.management.base import BaseCommand
from django.contrib.auth.models import User
from sales.models import MenuCategory, MenuItem, Waiter, UserProfile


class Command(BaseCommand):
    help = 'Setup DEEM Restaurant initial data (categories, menu, waiters, users)'
    
    def handle(self, *args, **kwargs):
        self.stdout.write(self.style.MIGRATE_HEADING('\n=== DEEM Sales Setup ===\n'))
        
        # ===== 1. CREATE USERS =====
        self.stdout.write(self.style.HTTP_INFO('1. Creating users...'))
        
        users_config = [
            {
                'username': 'admin',
                'password': 'admin123',
                'email': 'admin@deemrestaurant.co.tz',
                'first_name': 'System',
                'last_name': 'Administrator',
                'role': 'admin',
                'is_staff': True,
            },
            {
                'username': 'manager',
                'password': 'manager123',
                'email': 'manager@deemrestaurant.co.tz',
                'first_name': 'Order',
                'last_name': 'Manager',
                'role': 'manager',
                'is_staff': False,
            },
            {
                'username': 'cashier',
                'password': 'cashier123',
                'email': 'cashier@deemrestaurant.co.tz',
                'first_name': 'Sales',
                'last_name': 'Cashier',
                'role': 'cashier',
                'is_staff': False,
            },
        ]
        
        for cfg in users_config:
            user, created = User.objects.get_or_create(
                username=cfg['username'],
                defaults={
                    'email': cfg['email'],
                    'first_name': cfg['first_name'],
                    'last_name': cfg['last_name'],
                    'is_staff': cfg['is_staff'],
                }
            )
            
            if created:
                user.set_password(cfg['password'])
                user.save()
                self.stdout.write(self.style.SUCCESS(f'   ✓ Created: {cfg["username"]} ({cfg["role"]})'))
            else:
                self.stdout.write(self.style.WARNING(f'   • Exists: {cfg["username"]}'))
            
            # Ensure profile
            profile, _ = UserProfile.objects.get_or_create(user=user)
            profile.role = cfg['role']
            profile.is_active = True
            profile.save()
        
        # ===== 2. CREATE MENU CATEGORIES =====
        self.stdout.write(self.style.HTTP_INFO('\n2. Creating menu categories...'))
        
        categories = [
            ('Breakfast', 1),
            ('Main', 2),
            ('Drinks', 3),
            ('Choma', 4),
        ]
        
        for cat_name, order in categories:
            cat, created = MenuCategory.objects.get_or_create(
                name=cat_name,
                defaults={'display_order': order}
            )
            status = '✓ Created' if created else '• Exists'
            self.stdout.write(self.style.SUCCESS(f'   {status}: {cat_name}'))
        
        # ===== 3. CREATE MENU ITEMS =====
        self.stdout.write(self.style.HTTP_INFO('\n3. Creating menu items...'))
        
        menu_items = [
            # Breakfast
            ('Breakfast', 'Chai', 500),
            ('Breakfast', 'Chapati', 1000),
            # Main
            ('Main', 'Ugali Nyama', 5000),
            ('Main', 'Wali Samaki', 5500),
            # Drinks
            ('Drinks', 'Soda', 1500),
            ('Drinks', 'Maji', 1000),
            # Choma
            ('Choma', 'Nyama Choma', 5000),
            ('Choma', 'Ndizi', 8000),
        ]
        
        for cat_name, item_name, price in menu_items:
            try:
                category = MenuCategory.objects.get(name=cat_name)
                item, created = MenuItem.objects.get_or_create(
                    name=item_name,
                    category=category,
                    defaults={'price': price, 'is_available': True}
                )
                status = '✓' if created else '•'
                self.stdout.write(f'   {status} {item_name} - {price} TZS')
            except MenuCategory.DoesNotExist:
                self.stdout.write(self.style.ERROR(f'   ✗ Category "{cat_name}" not found!'))
        
        # ===== 4. CREATE WAITERS =====
        self.stdout.write(self.style.HTTP_INFO('\n4. Creating waiters...'))
        
        waiters = [
            ('Benedicto Joseph', 'Ben'),
            ('Mary Charles', 'Mary'),
            ('John Michael', 'John'),
        ]
        
        for full_name, display_name in waiters:
            waiter, created = Waiter.objects.get_or_create(
                full_name=full_name,
                defaults={'display_name': display_name, 'is_active': True}
            )
            status = '✓ Created' if created else '• Exists'
            self.stdout.write(self.style.SUCCESS(f'   {status}: {display_name} ({full_name})'))
        
        # ===== SUMMARY =====
        self.stdout.write(self.style.MIGRATE_HEADING('\n=== SETUP COMPLETE ===\n'))
        self.stdout.write(self.style.SUCCESS('Login Credentials:'))
        self.stdout.write(self.style.SUCCESS('  🔴 Admin   : admin   / admin123'))
        self.stdout.write(self.style.SUCCESS('  🟠 Manager : manager / manager123'))
        self.stdout.write(self.style.SUCCESS('  🟢 Cashier : cashier / cashier123'))
        self.stdout.write('')
        self.stdout.write(self.style.WARNING('⚠️  IMPORTANT: Change these passwords after first login!\n'))