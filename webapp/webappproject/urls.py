"""webappproject URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/3.0/topics/http/urls/
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
from django.contrib import admin
from django.urls import path, include
from pdfscanner import views
from django.conf.urls import url
from django.views.static import serve
from webappproject import settings
import django

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', views.index, name="index"),
    url(r'^index$', views.index, name="index"),
    url(r'^aboutus$', views.aboutus, name="aboutus"),
    path('analysepdf.html', views.analysepdf, name="analysepdf"),
    path('signup.html', views.signup, name="signup"),
    path('uploaddocument.html', views.uploaddocument, name="uploaddocumet"),
    path('pdfscanner/uploaddocument.html', views.uploaddocument, name='uploaddocumet'),
    path('register.html', views.register, name='register'),
    url('pdfscanner/register.html', views.register, name='register'),
    url(r'pdfscanner/user_login', views.user_login, name='user_login'),
    url(r'login.html', views.user_login, name='user_login'),
    url(r'logout', views.user_logout, name='logout'),
    path('analysedocument.html', views.documentsview, name='documentsview'),
    path('exportdetailstoexcel', views.exportdetailstoexcel, name='exportdetailstoexcel'),
    path('exporttopdf', views.exporttopdf, name='exporttopdf'),
    path('exportdicttoexceluvo', views.exportdicttoexceluvo, name='exportdicttoexceluvo'),
    path('deletehighlights', views.deletehighlights, name='deletehighlights'),
    path('extracthighlights', views.extracthighlights, name='extracthighlights'),
    url(r'^test.html/$', views.test, name='test'),
    path('<int:pk>', views.deletedocument, name='deletedocument'),
    path('highlightreview', views.highlightreview, name='highlightreview'),
    path('activate/<uidb64>/<token>/', views.activate, name='activate'),
    path(r'celery-progress/', include('celery_progress.urls')),
    url(r'analysisresult.html', views.analysisresult, name='analysisresult'),
    url(r'^settingspage$', views.settingspage, name='settingspage'),
    url('^downloadfilemeta', views.downloadfilemeta, name='downloadfilemeta'),
    url('^extractimages', views.extractimages, name='extractimages'),
    url('^pdftoimage', views.pdftoimage, name='pdftoimage'),
    path('privacypolicy.html', views.privacypolicy, name='privacypolicy'),
    path('enduseragreement.html', views.enduseragreement, name='enduseragreement'),
    path('contactus.html', views.contactus, name='contactus'),
    ## TODO look into this url creating error while migrating
    url(r'^(.*?)media/(?P<path>.*)$', serve,
        {'document_root': settings.MEDIA_ROOT}),
]
