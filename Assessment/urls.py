from django.urls import path
from . import views

urlpatterns = [
    path('list/', views.assessment_list, name='assessment_list'),
    path('take/<int:assessment_id>/', views.take_assessment, name='take_assessment'),
    path('result/<int:assessment_id>/', views.assessment_result, name='assessment_result'),
]
