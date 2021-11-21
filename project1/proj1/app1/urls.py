from django.contrib import admin
from django.urls import path
from app1 import views

urlpatterns = [
    path('submit', views.submit, name='submit'),
    path('', views.index, name='home')
    
]
