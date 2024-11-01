from django.urls import path
from .views import ReportUploadApi

urlpatterns = [
    path('uploader/',ReportUploadApi.as_view(),name = "upload-report"),
]