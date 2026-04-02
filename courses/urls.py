from django.urls import path

from . import views


urlpatterns = [
    path('', views.course_list, name='course_list'),
    path('<str:course_code>/', views.course_detail, name='course_detail'),
    path('<str:course_code>/workbook/', views.course_workbook_detail, name='course_workbook_detail'),
]
