from django.shortcuts import render, get_object_or_404, redirect
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.utils import timezone
from .models import Assessment, Question, Choice, AssessmentAttempt


@login_required
def assessment_list(request):
    """ List all available assessments, with optional course filtering. """
    course_filter = request.GET.get('course')
    assessments = Assessment.objects.all().order_by('-created_at')
    
    if course_filter:
        assessments = assessments.filter(course=course_filter)
        
    return render(request, 'Assessment/assessment_list.html', {
        'assessments': assessments,
        'course_filter': course_filter
    })


@login_required
def take_assessment(request, assessment_id):
    """
    Core page for taking an assessment.
    Handles rendering the questions and the final submission.
    """
    assessment = get_object_or_404(Assessment, id=assessment_id)
    questions = assessment.questions.all().prefetch_related('choices')

    # Optional: Check if student has already completed this assessment
    previous_attempt = AssessmentAttempt.objects.filter(
        user=request.user, assessment=assessment, completed=True
    ).exists()
    if previous_attempt:
        messages.info(request, "You have already completed this assessment.")
        return redirect('assessment_list')

    if request.method == 'POST':
        # Simple scoring logic: 1 point per correct answer for now.
        total_questions = questions.count()
        obtained_score = 0
        
        # Iterate over questions, check POST data
        for question in questions:
            selected_choice_id = request.POST.get(f'question_{question.id}')
            if selected_choice_id:
                try:
                    choice = Choice.objects.get(id=selected_choice_id, question=question)
                    if choice.is_correct:
                        obtained_score += question.marks
                except Choice.DoesNotExist:
                    pass

        # Create/Update attempt
        attempt, created = AssessmentAttempt.objects.get_or_create(
            user=request.user, 
            assessment=assessment,
            defaults={'score': obtained_score, 'completed': True, 'end_time': timezone.now()}
        )
        if not created:
            attempt.score = obtained_score
            attempt.completed = True
            attempt.end_time = timezone.now()
            attempt.save()

        messages.success(request, f"Assessment submitted! You scored {obtained_score} marks.")
        return redirect('assessment_result', assessment_id=assessment.id)

    return render(request, 'Assessment/take_assessment.html', {
        'assessment': assessment,
        'questions': questions
    })


@login_required
def assessment_result(request, assessment_id):
    """ Show the results of the student's attempt. """
    assessment = get_object_or_404(Assessment, id=assessment_id)
    attempt = get_object_or_404(AssessmentAttempt, user=request.user, assessment=assessment, completed=True)
    return render(request, 'Assessment/assessment_result.html', {
        'assessment': assessment,
        'attempt': attempt
    })


@login_required
def excel_viewer(request):
    """ View for reading Excel files in the browser. """
    return render(request, 'Assessment/excel_viewer.html')
