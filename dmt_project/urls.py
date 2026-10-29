"""
URL configuration for dmt_project project.
"""

from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static
from django.http import JsonResponse
from django.db import connection


def health_check(request):
    """
    Lightweight healthcheck endpoint for Docker, Coolify, and load balancers.
    Returns HTTP 200 with JSON payload when application and database are operational.
    """
    db_status = "ok"
    try:
        with connection.cursor() as cursor:
            cursor.execute("SELECT 1")
    except Exception as exc:
        db_status = f"unhealthy: {exc}"

    is_healthy = (db_status == "ok")
    status_code = 200 if is_healthy else 503
    return JsonResponse({
        "status": "healthy" if is_healthy else "degraded",
        "database": db_status,
        "app": "dmt",
    }, status=status_code)


urlpatterns = [
    path('health/', health_check, name='health_check'),
    path('healthz', health_check, name='healthz'),
    path('admin/', admin.site.urls),
    path('', include('hub.urls', namespace='hub')),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
    urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)
