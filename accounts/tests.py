from django.test import TestCase
from django.urls import reverse

from .forms import UserRegistrationForm
from .models import User


class RegistrationSecurityTests(TestCase):
    def test_login_page_renders(self):
        response = self.client.get(reverse('login'))

        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, 'accounts/login.html')
        self.assertContains(response, 'Login')
        self.assertContains(response, 'datamites.png')

    def test_register_page_renders(self):
        response = self.client.get(reverse('register'))

        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, 'accounts/register.html')
        self.assertContains(response, 'Register')
        self.assertContains(response, 'datamites.png')
        self.assertContains(response, 'Email')

    def test_registration_form_does_not_expose_role_field(self):
        form = UserRegistrationForm()
        self.assertNotIn('role', form.fields)
        self.assertIn('email', form.fields)
        self.assertTrue(form.fields['email'].required)

    def test_register_view_forces_student_role_even_if_post_is_tampered(self):
        response = self.client.post(
            reverse('register'),
            {
                'username': 'tampered-user',
                'email': 'tampered@example.com',
                'password1': 'StrongPass123!',
                'password2': 'StrongPass123!',
                'role': User.IS_ADMIN,
            },
        )

        self.assertRedirects(
            response,
            reverse('login'),
            fetch_redirect_response=False,
        )
        user = User.objects.get(username='tampered-user')
        self.assertEqual('tampered@example.com', user.email)
        self.assertEqual(User.IS_STUDENT, user.role)
        self.assertFalse(user.is_staff)
        self.assertFalse(user.is_superuser)
        self.assertNotIn('_auth_user_id', self.client.session)

    def test_register_redirects_to_login_with_success_message(self):
        response = self.client.post(
            reverse('register'),
            {
                'username': 'new-student',
                'email': 'new-student@example.com',
                'password1': 'StrongPass123!',
                'password2': 'StrongPass123!',
            },
            follow=True,
        )

        self.assertRedirects(response, reverse('login'))
        self.assertContains(response, 'Account created successfully. Please log in.')
        self.assertNotIn('_auth_user_id', self.client.session)
        self.assertEqual(
            'new-student@example.com',
            User.objects.get(username='new-student').email,
        )

    def test_student_cannot_access_admin_dashboard(self):
        user = User.objects.create_user(
            username='student-user',
            password='StrongPass123!',
            role=User.IS_STUDENT,
        )
        self.client.force_login(user)

        response = self.client.get(reverse('admin_dashboard'))

        self.assertRedirects(
            response,
            reverse('student_dashboard'),
            fetch_redirect_response=False,
        )

    def test_student_login_redirects_with_success_message(self):
        User.objects.create_user(
            username='student-user',
            password='StrongPass123!',
            role=User.IS_STUDENT,
        )

        response = self.client.post(
            reverse('login'),
            {
                'username': 'student-user',
                'password': 'StrongPass123!',
            },
            follow=True,
        )

        self.assertRedirects(response, reverse('student_dashboard'))
        self.assertContains(response, 'Login successful.')

    def test_create_superuser_sets_admin_role(self):
        user = User.objects.create_superuser(
            username='admin-user',
            password='StrongPass123!',
        )

        self.assertEqual(User.IS_ADMIN, user.role)
        self.assertTrue(user.is_portal_admin)

    def test_superuser_login_redirects_to_admin_dashboard(self):
        User.objects.create_superuser(
            username='admin-user',
            password='StrongPass123!',
        )

        response = self.client.post(
            reverse('login'),
            {
                'username': 'admin-user',
                'password': 'StrongPass123!',
            },
        )

        self.assertRedirects(
            response,
            reverse('admin_dashboard'),
            fetch_redirect_response=False,
        )
    
    def test_superuser_login_redirects_with_success_message(self):
        User.objects.create_superuser(
            username='admin-user',
            password='StrongPass123!',
        )

        response = self.client.post(
            reverse('login'),
            {
                'username': 'admin-user',
                'password': 'StrongPass123!',
            },
            follow=True,
        )

        self.assertRedirects(response, reverse('admin_dashboard'))
        self.assertContains(response, 'Login successful.')

    def test_superuser_is_redirected_away_from_student_dashboard(self):
        user = User.objects.create_superuser(
            username='admin-user',
            password='StrongPass123!',
        )
        self.client.force_login(user)

        response = self.client.get(reverse('student_dashboard'))

        self.assertRedirects(
            response,
            reverse('admin_dashboard'),
            fetch_redirect_response=False,
        )

    def test_admin_dashboard_shows_profile_settings_and_logout_links(self):
        user = User.objects.create_superuser(
            username='admin-user',
            password='StrongPass123!',
        )
        self.client.force_login(user)

        response = self.client.get(reverse('admin_dashboard'))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, reverse('admin_profile'))
        self.assertContains(response, reverse('admin_settings'))
        self.assertContains(response, reverse('logout'))

    def test_admin_profile_page_renders_for_admin(self):
        user = User.objects.create_superuser(
            username='admin-user',
            email='admin@example.com',
            password='StrongPass123!',
        )
        self.client.force_login(user)

        response = self.client.get(reverse('admin_profile'))

        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, 'accounts/admin_profile.html')
        self.assertContains(response, 'admin@example.com')
        self.assertContains(response, 'Admin Profile')

    def test_admin_settings_page_renders_for_admin(self):
        user = User.objects.create_superuser(
            username='admin-user',
            password='StrongPass123!',
        )
        self.client.force_login(user)

        response = self.client.get(reverse('admin_settings'))

        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, 'accounts/admin_settings.html')
        self.assertContains(response, 'Settings')
        self.assertContains(response, 'Notification Preferences')

    def test_student_cannot_access_admin_profile_or_settings(self):
        user = User.objects.create_user(
            username='student-user',
            password='StrongPass123!',
            role=User.IS_STUDENT,
        )
        self.client.force_login(user)

        profile_response = self.client.get(reverse('admin_profile'))
        settings_response = self.client.get(reverse('admin_settings'))

        self.assertRedirects(
            profile_response,
            reverse('student_dashboard'),
            fetch_redirect_response=False,
        )
        self.assertRedirects(
            settings_response,
            reverse('student_dashboard'),
            fetch_redirect_response=False,
        )

    def test_invalid_login_shows_form_error(self):
        User.objects.create_user(
            username='registered-user',
            password='StrongPass123!',
            role=User.IS_STUDENT,
        )

        response = self.client.post(
            reverse('login'),
            {
                'username': 'registered-user',
                'password': 'wrong-password',
            },
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Invalid username or password.')
