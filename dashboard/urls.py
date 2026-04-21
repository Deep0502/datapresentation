from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('upload/', views.upload_file, name='upload_file'),
    path('filter/', views.filter_data, name='filter_data'),
    path('export/', views.export_excel, name='export_excel'),
]