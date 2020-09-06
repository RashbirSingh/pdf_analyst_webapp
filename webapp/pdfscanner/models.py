from django.db import models
from django.contrib.auth.models import User

class FileDocument(models.Model):
    # Create relationship (don't inherit from User!)
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    description = models.CharField(max_length=255, blank=True)
    file_field = models.FileField(upload_to='documents/')
    uploaded_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        # Built-in attribute of django.contrib.auth.models.User !
        return self.user.username
    def delete(self):
        self.file_field.delete()
        super().delete()

class UserSettingsDocument(models.Model):

    user = models.ForeignKey(User, on_delete=models.CASCADE)
    setting_field = models.FileField(upload_to='usersettings/')

    def __str__(self):
        # Built-in attribute of django.contrib.auth.models.User !
        return self.user.username
    def delete(self):
        self.setting_field.delete()
        super().delete()

class UserProfileInfo(models.Model):
    # Create relationship (don't inherit from User!)
    user = models.OneToOneField(User, on_delete=models.CASCADE)

    # Add any additional attributes you want
    #portfolio_site = models.URLField(blank=True)
    profile_pic = models.ImageField(upload_to='profile_pics/', blank=True)

    def __str__(self):
        # Built-in attribute of django.contrib.auth.models.User !
        return self.user.username
