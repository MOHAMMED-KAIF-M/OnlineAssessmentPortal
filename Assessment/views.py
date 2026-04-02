import random

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone

from questions.models import QuestionBankEntry
from questions.question_types import ALLOWED_QUESTION_TYPES

from .forms import AssessmentGenerationForm
from .generation import (
    build_question_bank_hierarchy,
    normalize_text,
    split_module_path,
)
from .models import Assessment, AssessmentAttempt, Choice, Question


CHOICE_LABELS = ('A', 'B', 'C', 'D')


def _redirect_if_not_admin(request):
    if not request.user.is_portal_admin:
        return redirect('student_dashboard')
    return None


def _serialize_generation_filters(cleaned_data):
    return {
        'title': cleaned_data['title'],
        'description': cleaned_data['description'],
        'course': cleaned_data['course'],
        'subcourses': list(cleaned_data['subcourses']),
        'modules': list(cleaned_data['modules']),
        'workbooks': list(cleaned_data['workbooks']),
        'topics': list(cleaned_data['topics']),
        'question_types': list(cleaned_data['question_types']),
        'difficulties': list(cleaned_data['difficulties']),
        'question_count': cleaned_data['question_count'],
        'duration_minutes': cleaned_data['duration_minutes'],
        'marks_per_question': cleaned_data['marks_per_question'],
        'randomize': cleaned_data['randomize'],
    }


def _entry_has_complete_options(entry):
    return all(
        normalize_text(option_value)
        for option_value in (entry.option_a, entry.option_b, entry.option_c, entry.option_d)
    )


def _eligible_question_bank_entries(cleaned_data):
    return _scoped_question_bank_entries(cleaned_data, include_question_types=True)


def _scoped_question_bank_entries(cleaned_data, include_question_types=True):
    queryset = QuestionBankEntry.objects.filter(
        course_code=cleaned_data['course'],
    ).exclude(question_text='')
    if include_question_types:
        queryset = queryset.filter(question_type__in=cleaned_data['question_types'])

    if cleaned_data['difficulties']:
        queryset = queryset.filter(difficulty__in=cleaned_data['difficulties'])
    if cleaned_data['topics']:
        queryset = queryset.filter(topic__in=cleaned_data['topics'])

    selected_subcourses = set(cleaned_data['subcourses'])
    selected_modules = set(cleaned_data['modules'])
    selected_workbooks = set(cleaned_data['workbooks'])
    entries = []
    for entry in queryset.order_by('source_file', 'row_number'):
        if not normalize_text(entry.question_text) or not _entry_has_complete_options(entry):
            continue
        subcourse, module, workbook_from_path = split_module_path(entry.module_path)
        workbook = normalize_text(entry.workbook_title) or workbook_from_path
        if selected_subcourses and subcourse not in selected_subcourses:
            continue
        if selected_modules and module not in selected_modules:
            continue
        if selected_workbooks and workbook not in selected_workbooks:
            continue
        entries.append(entry)
    return entries


def _available_question_type_counts(cleaned_data):
    counts = {}
    for entry in _scoped_question_bank_entries(cleaned_data, include_question_types=False):
        question_type = normalize_text(entry.question_type)
        if question_type not in ALLOWED_QUESTION_TYPES:
            continue
        counts[question_type] = counts.get(question_type, 0) + 1
    return counts


def _format_type_counts(type_counts):
    ordered = []
    for question_type in ALLOWED_QUESTION_TYPES:
        count = type_counts.get(question_type)
        if count:
            ordered.append(f'{question_type}: {count}')
    return ', '.join(ordered)


def _select_entries_for_generation(cleaned_data):
    entries = _eligible_question_bank_entries(cleaned_data)
    requested_count = cleaned_data['question_count']
    if len(entries) < requested_count:
        return entries, []
    if cleaned_data['randomize']:
        return entries, random.sample(entries, requested_count)
    return entries, entries[:requested_count]


def _parse_selected_entry_ids(raw_value):
    selected_ids = []
    seen_ids = set()
    for chunk in str(raw_value or '').split(','):
        chunk = chunk.strip()
        if not chunk:
            continue
        try:
            entry_id = int(chunk)
        except ValueError:
            continue
        if entry_id <= 0 or entry_id in seen_ids:
            continue
        seen_ids.add(entry_id)
        selected_ids.append(entry_id)
    return selected_ids


def _selected_entries_by_ids(course_code, selected_ids):
    entries_by_id = {
        entry.id: entry
        for entry in QuestionBankEntry.objects.filter(
            course_code=course_code,
            id__in=selected_ids,
        ).order_by('source_file', 'row_number')
    }
    return [entries_by_id[entry_id] for entry_id in selected_ids if entry_id in entries_by_id]


def _preview_entry(entry, selected_correct=''):
    subcourse, module, workbook_from_path = split_module_path(entry.module_path)
    options = []
    for label, value in zip(
        CHOICE_LABELS,
        (entry.option_a, entry.option_b, entry.option_c, entry.option_d),
    ):
        options.append({'label': label, 'value': value})

    return {
        'id': entry.id,
        'question_text': entry.question_text,
        'question_type': entry.question_type or 'Unspecified',
        'difficulty': entry.difficulty or 'Unspecified',
        'topic': entry.topic or 'Unspecified',
        'subcourse': subcourse or 'Unspecified',
        'module': module or 'Unspecified',
        'workbook_title': normalize_text(entry.workbook_title) or workbook_from_path or 'Unspecified',
        'context': entry.context,
        'details': entry.details,
        'options': options,
        'selected_correct': selected_correct,
    }


def _build_assessment_question_text(entry):
    parts = []
    if normalize_text(entry.context):
        parts.append(f"Scenario/Context: {normalize_text(entry.context)}")
    if normalize_text(entry.details):
        parts.append(f"Skills Tested: {normalize_text(entry.details)}")
    parts.append(normalize_text(entry.question_text))
    return "\n\n".join(parts)


def _create_assessment_from_entries(cleaned_data, selected_entries, selected_correct_map):
    with transaction.atomic():
        assessment = Assessment.objects.create(
            title=cleaned_data['title'],
            description=cleaned_data['description'],
            course=cleaned_data['course'],
            duration_minutes=cleaned_data['duration_minutes'],
            total_marks=cleaned_data['question_count'] * cleaned_data['marks_per_question'],
        )
        choice_batch = []
        for entry in selected_entries:
            question = Question.objects.create(
                assessment=assessment,
                text=_build_assessment_question_text(entry),
                marks=cleaned_data['marks_per_question'],
            )
            correct_label = selected_correct_map[entry.id]
            for label, value in zip(
                CHOICE_LABELS,
                (entry.option_a, entry.option_b, entry.option_c, entry.option_d),
            ):
                choice_batch.append(
                    Choice(
                        question=question,
                        text=value,
                        is_correct=label == correct_label,
                    )
                )
        Choice.objects.bulk_create(choice_batch)
    return assessment


@login_required
def generate_assessment(request):
    redirect_response = _redirect_if_not_admin(request)
    if redirect_response:
        return redirect_response

    hierarchy = build_question_bank_hierarchy()
    form = AssessmentGenerationForm(
        request.POST or None,
        hierarchy=hierarchy,
        initial={
            'course': Assessment.COURSE_CHOICES[0][0],
            'question_types': ['MCQ'],
            'question_count': 10,
            'duration_minutes': 30,
            'marks_per_question': 1,
        },
    )
    preview_entries = []
    persisted_filters = None
    selected_entry_ids = ''
    matched_question_count = 0

    if request.method == 'POST':
        action = request.POST.get('action', 'preview')
        if form.is_valid():
            persisted_filters = _serialize_generation_filters(form.cleaned_data)
            if action == 'create':
                selected_ids = _parse_selected_entry_ids(request.POST.get('selected_entry_ids'))
                selected_entries = _selected_entries_by_ids(form.cleaned_data['course'], selected_ids)
                eligible_entry_ids = {
                    entry.id for entry in _eligible_question_bank_entries(form.cleaned_data)
                }
                if len(selected_entries) != len(selected_ids):
                    form.add_error(None, 'The question preview expired. Preview the assessment again and retry.')
                elif any(entry.id not in eligible_entry_ids for entry in selected_entries):
                    form.add_error(None, 'The selected questions no longer match the chosen hierarchy filters. Preview the assessment again and retry.')
                elif len(selected_entries) != form.cleaned_data['question_count']:
                    form.add_error(None, 'The selected question count no longer matches the requested assessment size.')
                else:
                    selected_correct_map = {}
                    has_correct_choice_error = False
                    for entry in selected_entries:
                        selected_label = normalize_text(
                            request.POST.get(f'correct_choice_{entry.id}')
                        ).upper()
                        if selected_label not in CHOICE_LABELS:
                            has_correct_choice_error = True
                        selected_correct_map[entry.id] = selected_label

                    if has_correct_choice_error:
                        form.add_error(None, 'Choose the correct option for every previewed question before creating the assessment.')
                    else:
                        assessment = _create_assessment_from_entries(
                            form.cleaned_data,
                            selected_entries,
                            selected_correct_map,
                        )
                        messages.success(
                            request,
                            f'Assessment "{assessment.title}" created with {len(selected_entries)} questions.',
                        )
                        return redirect('generate_assessment')

                selected_entry_ids = ','.join(str(entry_id) for entry_id in selected_ids)
                matched_question_count = len(_eligible_question_bank_entries(form.cleaned_data))
                preview_entries = [
                    _preview_entry(
                        entry,
                        selected_correct=normalize_text(
                            request.POST.get(f'correct_choice_{entry.id}')
                        ).upper(),
                    )
                    for entry in selected_entries
                ]
            else:
                eligible_entries, selected_entries = _select_entries_for_generation(form.cleaned_data)
                matched_question_count = len(eligible_entries)
                if len(selected_entries) < form.cleaned_data['question_count']:
                    form.add_error(
                        None,
                        (
                            f'Only {len(eligible_entries)} question-bank rows match these filters. '
                            'Adjust the options or reduce the question count.'
                        ),
                    )
                else:
                    preview_entries = [_preview_entry(entry) for entry in selected_entries]
                    selected_entry_ids = ','.join(str(entry.id) for entry in selected_entries)

    return render(
        request,
        'Assessment/generate_assessment.html',
        {
            'admin_page': 'assessments',
            'form': form,
            'hierarchy_panels': [
                {
                    'name': 'subcourses',
                    'label': 'Subcourses',
                    'hint': 'Choose one or more subcourses inside the selected course.',
                    'field': form['subcourses'],
                },
                {
                    'name': 'modules',
                    'label': 'Modules',
                    'hint': 'Modules narrow down automatically from the chosen subcourses.',
                    'field': form['modules'],
                },
                {
                    'name': 'workbooks',
                    'label': 'Workbooks',
                    'hint': 'Workbooks inherit from the selected modules to keep the list manageable.',
                    'field': form['workbooks'],
                },
                {
                    'name': 'topics',
                    'label': 'Topics',
                    'hint': 'Use topics for the final cut when you need a very specific slice.',
                    'field': form['topics'],
                },
            ],
            'preview_entries': preview_entries,
            'persisted_filters': persisted_filters,
            'selected_entry_ids': selected_entry_ids,
            'matched_question_count': matched_question_count,
            'hierarchy_data': hierarchy,
        },
    )


@login_required
def assessment_list(request):
    """ List all available assessments, with optional course filtering. """
    course_filter = request.GET.get('course')
    assessments = Assessment.objects.all().order_by('-created_at')

    if course_filter:
        assessments = assessments.filter(course=course_filter)

    return render(
        request,
        'Assessment/assessment_list.html',
        {
            'assessments': assessments,
            'course_filter': course_filter,
        },
    )


@login_required
def take_assessment(request, assessment_id):
    """
    Core page for taking an assessment.
    Handles rendering the questions and the final submission.
    """
    assessment = get_object_or_404(Assessment, id=assessment_id)
    questions = assessment.questions.all().prefetch_related('choices')

    previous_attempt = AssessmentAttempt.objects.filter(
        user=request.user,
        assessment=assessment,
        completed=True,
    ).exists()
    if previous_attempt:
        messages.info(request, "You have already completed this assessment.")
        return redirect('assessment_list')

    if request.method == 'POST':
        obtained_score = 0

        for question in questions:
            selected_choice_id = request.POST.get(f'question_{question.id}')
            if not selected_choice_id:
                continue
            try:
                choice = Choice.objects.get(id=selected_choice_id, question=question)
            except Choice.DoesNotExist:
                continue
            if choice.is_correct:
                obtained_score += question.marks

        attempt, created = AssessmentAttempt.objects.get_or_create(
            user=request.user,
            assessment=assessment,
            defaults={
                'score': obtained_score,
                'completed': True,
                'end_time': timezone.now(),
            },
        )
        if not created:
            attempt.score = obtained_score
            attempt.completed = True
            attempt.end_time = timezone.now()
            attempt.save()

        messages.success(request, f"Assessment submitted! You scored {obtained_score} marks.")
        return redirect('assessment_result', assessment_id=assessment.id)

    return render(
        request,
        'Assessment/take_assessment.html',
        {
            'assessment': assessment,
            'questions': questions,
        },
    )


@login_required
def assessment_result(request, assessment_id):
    """ Show the results of the student's attempt. """
    assessment = get_object_or_404(Assessment, id=assessment_id)
    attempt = get_object_or_404(
        AssessmentAttempt,
        user=request.user,
        assessment=assessment,
        completed=True,
    )
    return render(
        request,
        'Assessment/assessment_result.html',
        {
            'assessment': assessment,
            'attempt': attempt,
        },
    )


@login_required
def excel_viewer(request):
    """ View for reading Excel files in the browser. """
    return render(request, 'Assessment/excel_viewer.html')
