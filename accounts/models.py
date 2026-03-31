from django.contrib.auth.models import AbstractUser, UserManager as DjangoUserManager
from django.db import models


class PortalUserManager(DjangoUserManager):
    def create_user(self, username, email=None, password=None, **extra_fields):
        extra_fields.setdefault('role', self.model.IS_STUDENT)
        return super().create_user(username, email=email, password=password, **extra_fields)

    def create_superuser(self, username, email=None, password=None, **extra_fields):
        extra_fields.setdefault('is_staff', True)
        extra_fields.setdefault('is_superuser', True)
        extra_fields['role'] = self.model.IS_ADMIN
        return super().create_superuser(username, email=email, password=password, **extra_fields)


class User(AbstractUser):
    IS_STUDENT = 'student'
    IS_ADMIN = 'admin'

    ROLE_CHOICES = [
        (IS_STUDENT, 'Student'),
        (IS_ADMIN, 'Admin'),
    ]

    role = models.CharField(max_length=10, choices=ROLE_CHOICES, default=IS_STUDENT)
    objects = PortalUserManager()

    @property
    def is_portal_admin(self):
        return self.is_superuser or self.is_staff or self.role == self.IS_ADMIN

    @property
    def portal_role_display(self):
        return 'Admin' if self.is_portal_admin else self.get_role_display()

    def __str__(self):
        return f"{self.username} ({self.role})"
