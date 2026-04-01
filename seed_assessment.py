import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'OnlineAssessmentPortal.settings')
django.setup()

from Assessment.models import Assessment, Question, Choice

def seed_assessment():
    # Create Assessment
    assessment, created = Assessment.objects.get_or_create(
        title="Python Fundamentals",
        description="A basic assessment covering Python basics, data structures, and functions.",
        duration_minutes=15,
        total_marks=30
    )
    
    if created:
        # Question 1
        q1 = Question.objects.create(assessment=assessment, text="What is the correct way to declare a list in Python?", marks=10)
        Choice.objects.create(question=q1, text="[]", is_correct=True)
        Choice.objects.create(question=q1, text="{}", is_correct=False)
        Choice.objects.create(question=q1, text="()", is_correct=False)
        Choice.objects.create(question=q1, text="<>", is_correct=False)

        # Question 2
        q2 = Question.objects.create(assessment=assessment, text="Which keyword is used to define a function in Python?", marks=10)
        Choice.objects.create(question=q2, text="function", is_correct=False)
        Choice.objects.create(question=q2, text="def", is_correct=True)
        Choice.objects.create(question=q2, text="func", is_correct=False)
        Choice.objects.create(question=q2, text="method", is_correct=False)

        # Question 3
        q3 = Question.objects.create(assessment=assessment, text="What is the output of 3 * 3?", marks=10)
        Choice.objects.create(question=q3, text="6", is_correct=False)
        Choice.objects.create(question=q3, text="9", is_correct=True)
        Choice.objects.create(question=q3, text="12", is_correct=False)
        Choice.objects.create(question=q3, text="27", is_correct=False)

        print("Sample assessment created successfully.")
    else:
        print("Sample assessment already exists.")

if __name__ == "__main__":
    seed_assessment()
