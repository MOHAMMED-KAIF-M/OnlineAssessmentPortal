from django.db import models
from django.conf import settings


class Assessment(models.Model):
    title = models.CharField(max_length=200)
    description = models.TextField(blank=True)
    duration_minutes = models.PositiveIntegerField(default=30)
    total_marks = models.PositiveIntegerField(default=100)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.title


class Question(models.Model):
    assessment = models.ForeignKey(Assessment, related_name='questions', on_delete=models.CASCADE)
    text = models.TextField()
    marks = models.PositiveIntegerField(default=1)

    def __str__(self):
        return f"{self.assessment.title} - Question {self.id}"


class Choice(models.Model):
    question = models.ForeignKey(Question, related_name='choices', on_delete=models.CASCADE)
    text = models.CharField(max_length=500)
    is_correct = models.BooleanField(default=False)

    def __str__(self):
        return self.text


class AssessmentAttempt(models.Model):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE)
    assessment = models.ForeignKey(Assessment, on_delete=models.CASCADE)
    score = models.IntegerField(null=True, blank=True)
    start_time = models.DateTimeField(auto_now_add=True)
    end_time = models.DateTimeField(null=True, blank=True)
    completed = models.BooleanField(default=False)

    def __str__(self):
        return f"{self.user.username} - {self.assessment.title}"
