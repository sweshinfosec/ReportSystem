from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('reporting-dashboard/', views.test_reporting_dashboard, name='reporting_dashboard'),
    path('interim-pre/', views.interim_pre, name='interim_pre'),
    path('interim-post/', views.interim_post, name='interim_post'),
    path('word-report/', views.word_report, name='word_report'),
]