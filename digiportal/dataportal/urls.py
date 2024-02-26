from django.urls import path
from . import views

urlpatterns = [
    path('', views.home),
    path('iff', views.iff),
    path('ym', views.ym),
    path('mm', views.mm),
    path('iffstats', views.iffstats),
    path('iffquality', views.iffquality)
]