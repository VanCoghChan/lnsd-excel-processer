"""
URL configuration for lnsd_excel_processer project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/4.2/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
# from django.contrib import admin
from django.urls import path
from . import views

urlpatterns = [
    #    path('admin/', admin.site.urls),
    path('', views.index, name='index'),
    path('upload_metadata/', views.upload_and_get_metadata_view, name='upload_and_get_metadata'),
    path('analyze/', views.trigger_final_analysis_view, name='trigger_final_analysis'),
    path('download/<str:filename>/', views.download_result, name='download_result'),
]
