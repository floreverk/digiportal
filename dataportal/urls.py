from django.urls import path
from . import views

urlpatterns = [
    path('', views.home),
    path('iff', views.iff),
    path('ym', views.ym),
    path('mm', views.mm),
    path('iffstats', views.iffstats),
    path('ymstats', views.ymstats),
    path('mmstats', views.mmstats),
    path('iffq001', views.iffq001),
    path('ymq001', views.ymq001),
    path('mmq001', views.mmq001),
    path('t001', views.t001),
]