from django.contrib import admin
#from pdfscanner.models import User
from pdfscanner.models import FileDocument, UserProfileInfo, UserSettingsDocument
# Register your models here.
#admin.site.register(User)
admin.site.register(FileDocument)
admin.site.register(UserProfileInfo)
admin.site.register(UserSettingsDocument)
