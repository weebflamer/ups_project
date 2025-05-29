from django.urls import path
from . import views

from .views import download_template


urlpatterns = [
    path('branch/<int:pk>/', views.branch_detail, name='branch_detail'),
    path('branches/', views.branch_list, name='branch_list'),
    path('branches/upload/', views.upload_excel, name='upload_excel'),
    path('branches/export/', views.export_to_excel, name='export_to_excel'),
    path('branches/<int:pk>/edit/', views.edit_branch, name='edit_branch'),
    path('branches/<int:pk>/delete/', views.delete_branch, name='delete_branch'),
    path('branches/add/', views.add_branch, name='add_branch'),
    path('branches/download-template/', download_template, name='download_template'),
]