from django.contrib import admin
from django.urls import path
from Users.views import *

urlpatterns = [
    path('register/', register, name='register'),
    path('login/', Login, name='login'),

    path('logout/', logout_view, name='logout'),
    
]