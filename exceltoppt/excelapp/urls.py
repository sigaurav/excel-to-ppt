from django.urls import path
from . import views

app_name = 'excelapp'

urlpatterns = [
    path('upload/', views.upload_file, name='upload_file'),
    path('generate-ppt/', views.generate_ppt, name='generate_ppt'),
    path('download-ppt/<str:file_name>/', views.download_ppt, name='download_ppt'),
]
