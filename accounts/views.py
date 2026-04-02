from django.contrib import messages
from django.shortcuts import render, redirect
from django.contrib.auth import login, authenticate, logout
from django.contrib.auth.decorators import login_required
from .forms import UserRegistrationForm, UserLoginForm
from .models import User


def _redirect_if_not_admin(request):
    if not request.user.is_portal_admin:
        return redirect('student_dashboard')
    return None


def _admin_context(admin_page, **extra_context):
    return {'admin_page': admin_page, **extra_context}


def register_view(request):
    if request.user.is_authenticated:
        return redirect('dashboard')

    if request.method == 'POST':
        form = UserRegistrationForm(request.POST)
        if form.is_valid():
            form.save()
            messages.success(request, 'Account created successfully. Please log in.')
            return redirect('login')
    else:
        form = UserRegistrationForm()
    return render(request, 'accounts/register.html', {'form': form})


def login_view(request):
    if request.user.is_authenticated:
        return redirect('dashboard')

    if request.method == 'POST':
        form = UserLoginForm(request.POST)
        if form.is_valid():
            username = form.cleaned_data.get('username')
            password = form.cleaned_data.get('password')
            user = authenticate(username=username, password=password)
            if user:
                login(request, user)
                messages.success(request, 'Login successful.')
                if user.is_portal_admin:
                    return redirect('admin_dashboard')
                else:
                    return redirect('student_dashboard')
            form.add_error(None, 'Invalid username or password.')
    else:
        form = UserLoginForm()
    return render(request, 'accounts/login.html', {'form': form})


@login_required
def dashboard_redirect_view(request):
    if request.user.is_portal_admin:
        return redirect('admin_dashboard')
    return redirect('student_dashboard')


@login_required
def student_dashboard(request):
    if request.user.is_portal_admin:
        return redirect('admin_dashboard')
    return render(request, 'accounts/student_dashboard.html')


@login_required
def admin_dashboard(request):
    redirect_response = _redirect_if_not_admin(request)
    if redirect_response:
        return redirect_response
    return render(
        request,
        'accounts/admin_dashboard.html',
        _admin_context('dashboard'),
    )


@login_required
def admin_profile(request):
    redirect_response = _redirect_if_not_admin(request)
    if redirect_response:
        return redirect_response
    return render(
        request,
        'accounts/admin_profile.html',
        _admin_context('profile'),
    )


@login_required
def admin_settings(request):
    redirect_response = _redirect_if_not_admin(request)
    if redirect_response:
        return redirect_response
    return render(
        request,
        'accounts/admin_settings.html',
        _admin_context('settings'),
    )


def logout_view(request):
    logout(request)
    return redirect('login')
