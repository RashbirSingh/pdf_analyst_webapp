from django import forms
from django.contrib.auth.models import User
from pdfscanner.models import UserProfileInfo, FileDocument, UserSettingsDocument

class UserForm(forms.ModelForm):
    password = forms.CharField(widget=forms.PasswordInput())
    confirm_password = forms.CharField(widget=forms.PasswordInput())
    email = forms.CharField(max_length=75, required=True)
    class Meta():
        model = User
        fields = ('username', 'first_name', 'last_name', 'email', 'password')

    def clean(self):
        cleaned_data = super(UserForm, self).clean()
        password = cleaned_data.get("password")
        confirm_password = cleaned_data.get("confirm_password")

        if password != confirm_password:
            raise forms.ValidationError(
                "password and confirm_password does not match"
            )

class UserProfileInfoForm(forms.ModelForm):
    class Meta():
        model = UserProfileInfo
        fields = ('profile_pic',)


class documentform(forms.ModelForm):
    file_field = forms.FileField(widget=forms.ClearableFileInput(attrs={'multiple': True}))
    class Meta():
        model = FileDocument
        fields = ('description', 'file_field',)

class usersettingsform(forms.ModelForm):
    setting_field = forms.FileField(widget=forms.ClearableFileInput(attrs={'multiple': True}))
    class Meta():
        model = UserSettingsDocument
        fields = ('setting_field',)
