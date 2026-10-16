from django.contrib.auth import logout
from django.shortcuts import redirect
from django.urls import reverse
from django.utils import timezone


class SessionIdleTimeout:
    """
    Middleware that logs out users after a period of inactivity.
    Works alongside SESSION_COOKIE_AGE for double protection.
    """
    
    def __init__(self, get_response):
        self.get_response = get_response
        # 5 minutes idle timeout
        self.idle_timeout = 300
    
    def __call__(self, request):
        if request.user.is_authenticated:
            last_activity = request.session.get('last_activity')
            now = timezone.now().timestamp()
            
            if last_activity:
                elapsed = now - last_activity
                if elapsed > self.idle_timeout:
                    logout(request)
                    return redirect(reverse('login') + '?timeout=1')
            
            # Update last activity
            request.session['last_activity'] = now
        
        response = self.get_response(request)
        return response