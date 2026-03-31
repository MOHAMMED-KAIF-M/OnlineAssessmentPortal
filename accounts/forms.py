from django import forms
from django.contrib.auth.forms import UserCreationForm

from .models import User


INPUT_CLASS = (
    'block w-full rounded-2xl border border-slate-200 bg-slate-50/80 px-4 py-3.5 '
    'text-base text-slate-900 placeholder:text-slate-400 shadow-sm transition '
    'duration-200 focus:border-amber-400 focus:bg-white focus:ring-4 '
    'focus:ring-amber-100'
)


class UserRegistrationForm(UserCreationForm):
    email = forms.EmailField(required=True)

    class Meta(UserCreationForm.Meta):
        model = User
        fields = UserCreationForm.Meta.fields + ('email',)

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        field_config = {
            'username': {
                'placeholder': 'Choose a username',
                'autocomplete': 'username',
            },
            'email': {
                'placeholder': 'Enter your email address',
                'autocomplete': 'email',
            },
            'password1': {
                'placeholder': 'Create a password',
                'autocomplete': 'new-password',
            },
            'password2': {
                'placeholder': 'Confirm your password',
                'autocomplete': 'new-password',
            },
        }

        for name, field in self.fields.items():
            field.widget.attrs['class'] = INPUT_CLASS
            field.widget.attrs.update(field_config.get(name, {}))

    def save(self, commit=True):
        user = super().save(commit=False)
        # Public registration must never be able to elevate privileges.
        user.role = User.IS_STUDENT
        user.is_staff = False
        user.is_superuser = False
        if commit:
            user.save()
        return user


class UserLoginForm(forms.Form):
    username = forms.CharField(
        widget=forms.TextInput(
            attrs={
                'class': INPUT_CLASS,
                'placeholder': 'Enter your username',
                'autocomplete': 'username',
            }
        )
    )
    password = forms.CharField(
        widget=forms.PasswordInput(
            attrs={
                'class': INPUT_CLASS,
                'placeholder': 'Enter your password',
                'autocomplete': 'current-password',
            }
        )
    )
