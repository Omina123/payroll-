from django import forms
from django.contrib.auth.forms import UserCreationForm
from .models import CustomUser, Profile, validate_kenyan_phone

class RegistrationForm(UserCreationForm):
    first_name = forms.CharField(max_length=30, required=True)
    last_name = forms.CharField(max_length=30, required=True)
    email = forms.EmailField(required=True)
    phone_number = forms.CharField(max_length=12, validators=[validate_kenyan_phone], required=True)
    id_number = forms.CharField(max_length=20, required=False)

    class Meta(UserCreationForm.Meta):
        model = CustomUser
        # MUST include all fields you want to use in save()
        fields = ('email', 'first_name', 'last_name', 'phone_number')  

    def save(self, commit=True):
        user = super().save(commit=False)
        user.email = self.cleaned_data['email']
        user.first_name = self.cleaned_data['first_name']
        user.last_name = self.cleaned_data['last_name']
        
        # Now this will work because phone_number is in 'fields'
        user.username = self.cleaned_data['phone_number']
        user.user_type = '1'  

        if commit:
            user.save()
            # Check if profile already exists or create it
            Profile.objects.update_or_create(
                user=user,
                defaults={
                    'phone_number': self.cleaned_data.get('phone_number'),
                    'id_number': self.cleaned_data.get('id_number')
                }
            )
        return user
class LoginForm(forms.Form):
    username = forms.CharField(label="Email or Phone Number", widget=forms.TextInput(attrs={'class': 'form-control'}))
    password = forms.CharField(widget=forms.PasswordInput(attrs={'class': 'form-control'}))