from django.test import TestCase
from django.urls import reverse

from accounts.models import User
from questions.models import QuestionBankEntry

from .models import Assessment


class AssessmentGenerationTests(TestCase):
    def setUp(self):
        self.admin_user = User.objects.create_superuser(
            username='admin-user',
            password='StrongPass123!',
        )
        self.student_user = User.objects.create_user(
            username='student-user',
            password='StrongPass123!',
            role=User.IS_STUDENT,
        )

    def _create_question_bank_entry(
        self,
        row_number,
        suffix,
        *,
        subcourse='ADA-134',
        module='DATA ANALYTICS FOUNDATION',
        workbook='Regression Workbook',
        topic='Regression',
    ):
        return QuestionBankEntry.objects.create(
            source_workbook_id=1,
            source_file='0001.xlsx',
            source_path=f'Assessment Questions/CDA/{subcourse}/{module}/{workbook}.xlsx',
            course_code='CDA',
            module_path=f'{subcourse}/{module}/{workbook}',
            workbook_title=workbook,
            sheet_name='questions',
            row_number=row_number,
            topic=topic,
            question_type='MCQ',
            difficulty='Easy',
            question_text=f'Question {suffix}?',
            context='A short business scenario',
            details='Interpret the regression output',
            option_a='Option A',
            option_b='Option B',
            option_c='Option C',
            option_d='Option D',
        )

    def test_generate_assessment_requires_login(self):
        response = self.client.get(reverse('generate_assessment'))

        self.assertRedirects(
            response,
            f"{reverse('login')}?next={reverse('generate_assessment')}",
            fetch_redirect_response=False,
        )

    def test_student_cannot_access_generate_assessment(self):
        self.client.force_login(self.student_user)

        response = self.client.get(reverse('generate_assessment'))

        self.assertRedirects(
            response,
            reverse('student_dashboard'),
            fetch_redirect_response=False,
        )

    def test_admin_can_preview_and_create_assessment(self):
        self.client.force_login(self.admin_user)
        first_entry = self._create_question_bank_entry(1, 'One')
        second_entry = self._create_question_bank_entry(2, 'Two')

        preview_response = self.client.post(
            reverse('generate_assessment'),
            {
                'action': 'preview',
                'title': 'Generated CDA Assessment',
                'description': 'Built from workbook rows',
                'course': 'CDA',
                'subcourses': ['ADA-134'],
                'modules': ['DATA ANALYTICS FOUNDATION'],
                'workbooks': ['Regression Workbook'],
                'topics': ['Regression'],
                'question_types': ['MCQ'],
                'difficulties': ['Easy'],
                'question_count': 2,
                'duration_minutes': 25,
                'marks_per_question': 2,
                'randomize': 'on',
            },
        )

        self.assertEqual(preview_response.status_code, 200)
        self.assertContains(preview_response, 'Choose the Correct Option for Each Question')
        self.assertEqual(len(preview_response.context['preview_entries']), 2)

        create_response = self.client.post(
            reverse('generate_assessment'),
            {
                'action': 'create',
                'title': 'Generated CDA Assessment',
                'description': 'Built from workbook rows',
                'course': 'CDA',
                'subcourses': ['ADA-134'],
                'modules': ['DATA ANALYTICS FOUNDATION'],
                'workbooks': ['Regression Workbook'],
                'topics': ['Regression'],
                'question_types': ['MCQ'],
                'difficulties': ['Easy'],
                'question_count': 2,
                'duration_minutes': 25,
                'marks_per_question': 2,
                'randomize': 'on',
                'selected_entry_ids': f'{first_entry.id},{second_entry.id}',
                f'correct_choice_{first_entry.id}': 'B',
                f'correct_choice_{second_entry.id}': 'D',
            },
            follow=True,
        )

        self.assertRedirects(create_response, reverse('generate_assessment'))
        self.assertContains(create_response, 'Generated CDA Assessment')
        self.assertContains(create_response, 'created with 2 questions')

        assessment = Assessment.objects.get(title='Generated CDA Assessment')
        self.assertEqual(assessment.course, 'CDA')
        self.assertEqual(assessment.total_marks, 4)
        self.assertEqual(assessment.duration_minutes, 25)
        self.assertEqual(assessment.questions.count(), 2)

        created_questions = list(assessment.questions.order_by('id'))
        self.assertIn('Scenario/Context:', created_questions[0].text)
        self.assertEqual(created_questions[0].marks, 2)
        self.assertEqual(created_questions[0].choices.count(), 4)
        self.assertTrue(created_questions[0].choices.get(text='Option B').is_correct)
        self.assertTrue(created_questions[1].choices.get(text='Option D').is_correct)

    def test_preview_shows_error_when_matching_question_count_is_too_low(self):
        self.client.force_login(self.admin_user)
        self._create_question_bank_entry(1, 'Only')

        response = self.client.post(
            reverse('generate_assessment'),
            {
                'action': 'preview',
                'title': 'Short Assessment',
                'course': 'CDA',
                'question_types': ['MCQ'],
                'question_count': 2,
                'duration_minutes': 20,
                'marks_per_question': 1,
            },
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Only 1 question-bank rows match these filters.')
        self.assertEqual(response.context['preview_entries'], [])

    def test_preview_can_filter_by_subcourse_module_workbook_and_topic(self):
        self.client.force_login(self.admin_user)
        self._create_question_bank_entry(
            1,
            'Target',
            subcourse='ADA-134',
            module='DATA ANALYTICS FOUNDATION',
            workbook='Regression Workbook',
            topic='Regression',
        )
        self._create_question_bank_entry(
            2,
            'Other',
            subcourse='DAA-131',
            module='TABLEAU',
            workbook='Dashboard Basics',
            topic='Visualization',
        )

        response = self.client.post(
            reverse('generate_assessment'),
            {
                'action': 'preview',
                'title': 'Filtered Assessment',
                'course': 'CDA',
                'subcourses': ['ADA-134'],
                'modules': ['DATA ANALYTICS FOUNDATION'],
                'workbooks': ['Regression Workbook'],
                'topics': ['Regression'],
                'question_types': ['MCQ'],
                'question_count': 1,
                'duration_minutes': 20,
                'marks_per_question': 1,
            },
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(len(response.context['preview_entries']), 1)
        preview_entry = response.context['preview_entries'][0]
        self.assertEqual(preview_entry['subcourse'], 'ADA-134')
        self.assertEqual(preview_entry['module'], 'DATA ANALYTICS FOUNDATION')
        self.assertEqual(preview_entry['workbook_title'], 'Regression Workbook')
        self.assertEqual(preview_entry['topic'], 'Regression')
