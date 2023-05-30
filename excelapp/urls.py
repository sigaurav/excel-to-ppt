from django.urls import path
from . import views
from django.conf import settings
from django.conf.urls.static import static

app_name = 'excelapp'

urlpatterns = [
    path('upload/', views.upload_file, name='upload_file'),
    path('generate-ppt/', views.generate_ppt, name='generate_ppt'),
    path('download-ppt/<str:file_name>/', views.download_ppt, name='download_ppt'),
]
if settings.DEBUG:
    urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)