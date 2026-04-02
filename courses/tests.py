from django.test import TestCase
from django.urls import reverse

from accounts.models import User
from courses.views import _normalize_packed_question_rows, _strip_answer_columns
from questions.models import QuestionBankEntry


class CourseListTests(TestCase):
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

    def test_course_list_requires_login(self):
        response = self.client.get(reverse('course_list'))

        self.assertRedirects(
            response,
            f"{reverse('login')}?next={reverse('course_list')}",
            fetch_redirect_response=False,
        )

    def test_admin_can_view_course_list(self):
        self.client.force_login(self.admin_user)

        response = self.client.get(reverse('course_list'))

        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, 'courses/course_list.html')
        for course_code in ('AIE', 'CDA', 'CDS', 'CDE'):
            self.assertContains(response, course_code)
            self.assertContains(response, reverse('course_detail', args=[course_code]))

    def test_admin_dashboard_shows_courses_button(self):
        self.client.force_login(self.admin_user)

        response = self.client.get(reverse('admin_dashboard'))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, reverse('course_list'))
        self.assertContains(response, 'Courses')

    def test_course_list_shows_admin_corner_actions(self):
        self.client.force_login(self.admin_user)

        response = self.client.get(reverse('course_list'))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, reverse('admin_profile'))
        self.assertContains(response, reverse('admin_settings'))
        self.assertContains(response, reverse('logout'))

    def test_admin_can_view_course_hierarchy(self):
        self.client.force_login(self.admin_user)

        response = self.client.get(reverse('course_detail', args=['AIE']))

        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, 'courses/course_detail.html')
        self.assertContains(response, 'AIE')
        self.assertContains(response, 'ADS-114')

    def test_admin_can_view_workbook_details(self):
        self.client.force_login(self.admin_user)
        relative_path = 'AIE/ADS-114/ADVANCED DATA ANALYSIS WITH MS EXCEL/Advanced Functions (VLOOKUP, INDIRECT..)/Advanced Functions (VLOOKUP, INDIRECT..).xlsx'
        QuestionBankEntry.objects.create(
            source_workbook_id=1,
            source_file='0001.xlsx',
            source_path=f'Assessment Questions/{relative_path}',
            course_code='AIE',
            module_path='ADS-114/ADVANCED DATA ANALYSIS WITH MS EXCEL/Advanced Functions (VLOOKUP, INDIRECT..)',
            workbook_title='Advanced Functions (VLOOKUP, INDIRECT..)',
            sheet_name='questions',
            row_number=1,
            question_text='What does VLOOKUP return?',
            question_type='MCQ',
            difficulty='Easy',
            option_a='A row label',
            option_b='A matching value',
            correct_answer='A matching value',
        )

        response = self.client.get(
            reverse('course_workbook_detail', args=['AIE']),
            {'path': relative_path},
        )

        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, 'courses/workbook_detail.html')
        self.assertContains(response, 'Workbook Sheets')
        self.assertContains(response, 'Question No.')
        self.assertContains(response, 'MCQ: 1')
        self.assertContains(response, 'Sheet 1 View')

    def test_admin_can_switch_between_workbook_sheets(self):
        self.client.force_login(self.admin_user)
        relative_path = 'CDA/BDF-117/BIG DATA INTRODUCTION/Big Data Analytics Introduction/Big Data Analytics Introduction.xlsx'

        response = self.client.get(
            reverse('course_workbook_detail', args=['CDA']),
            {'path': relative_path, 'sheet': 'Sheet2'},
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Sheet 1 View')
        self.assertContains(response, 'Sheet 2 View')
        self.assertContains(response, 'Selected Sheet')
        self.assertContains(response, 'Question No.')

    def test_bdf_dataset_sheet_is_expanded_into_columns(self):
        self.client.force_login(self.admin_user)
        relative_path = 'CDA/BDF-117/BIG DATA INTRODUCTION/Big Data Analytics Introduction/Big Data Analytics Introduction.xlsx'

        response = self.client.get(
            reverse('course_workbook_detail', args=['CDA']),
            {'path': relative_path, 'sheet': 'dataset'},
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response.context['workbook']['active_sheet']['header'][:5],
            ['Analytics_ID', 'Timestamp', 'Analytics_Type', 'Industry', 'Data_Source'],
        )
        self.assertEqual(
            response.context['workbook']['active_sheet']['body_rows'][0][:5],
            ['BA-001', '2024-03-01 00:00:01', 'Descriptive', 'Retail', 'CRM'],
        )
        self.assertEqual(
            response.context['workbook']['active_sheet']['header'][-1],
            'Question Type',
        )

    def test_bia_dataset_sheet_is_expanded_into_columns(self):
        self.client.force_login(self.admin_user)
        relative_path = 'AIE/BIA - 119/POWER-BI Basics/Basic Data Cleaning/Basic Data Cleaning.xlsx'

        response = self.client.get(
            reverse('course_workbook_detail', args=['AIE']),
            {'path': relative_path, 'sheet': 'dataset'},
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response.context['workbook']['active_sheet']['header'][:5],
            ['Row_ID', 'Order_ID', 'Order_Date', 'Product_Name', 'Category'],
        )
        self.assertEqual(
            response.context['workbook']['active_sheet']['body_rows'][0][:5],
            ['1', 'ORD-1001', '2024-01-05', 'MacBook Pro', 'Electronics'],
        )
        self.assertEqual(
            response.context['workbook']['active_sheet']['header'][-1],
            'Question Type',
        )

    def test_missing_workbook_path_returns_404(self):
        self.client.force_login(self.admin_user)

        response = self.client.get(reverse('course_workbook_detail', args=['AIE']))

        self.assertEqual(response.status_code, 404)

    def test_unknown_course_returns_404(self):
        self.client.force_login(self.admin_user)

        response = self.client.get(reverse('course_detail', args=['UNKNOWN']))

        self.assertEqual(response.status_code, 404)

    def test_student_is_redirected_away_from_course_hierarchy(self):
        self.client.force_login(self.student_user)

        response = self.client.get(reverse('course_detail', args=['AIE']))

        self.assertRedirects(
            response,
            reverse('student_dashboard'),
            fetch_redirect_response=False,
        )

    def test_student_is_redirected_away_from_workbook_detail(self):
        self.client.force_login(self.student_user)

        response = self.client.get(
            reverse('course_workbook_detail', args=['AIE']),
            {'path': 'AIE/ADS-114/ADVANCED DATA ANALYSIS WITH MS EXCEL/Advanced Functions (VLOOKUP, INDIRECT..)/Advanced Functions (VLOOKUP, INDIRECT..).xlsx'},
        )

        self.assertRedirects(
            response,
            reverse('student_dashboard'),
            fetch_redirect_response=False,
        )

    def test_student_is_redirected_away_from_course_list(self):
        self.client.force_login(self.student_user)

        response = self.client.get(reverse('course_list'))

        self.assertRedirects(
            response,
            reverse('student_dashboard'),
            fetch_redirect_response=False,
        )

    def test_strip_answer_columns_removes_answer_headers_from_sheet_rows(self):
        rows = [
            ['Question', 'Option A', 'Correct Answer', 'Question Type'],
            ['What is SQL?', 'Structured Query Language', 'A', 'MCQ'],
        ]

        stripped = _strip_answer_columns(rows)

        self.assertEqual(
            stripped,
            [
                ['Question', 'Option A', 'Question Type'],
                ['What is SQL?', 'Structured Query Language', 'MCQ'],
            ],
        )

    def test_normalize_packed_question_rows_unpacks_embedded_mcq_values(self):
        rows = [
            ['Question No.', 'Topic', 'Question Type', 'Question', 'Scenario/Context', 'Skills Tested', 'Option A', 'Option B', 'Option C', 'Option D'],
            [1, 'Business Analytics Overview', 'MCQ', 'Section,Type,Q.No,Scenario,Question,Option A,Option B,Option C,Option D,Correct Answer', '', '', '', '', '', ''],
            [2, 'Business Analytics Overview', 'MCQ', 'SECTION A: Medium Level MCQs,MCQ,1,,What is Business Analytics?,Raw data only,Using data for decision-making,Presentation design,Database storage,B', '', '', '', '', '', ''],
        ]

        normalized = _normalize_packed_question_rows(rows)

        self.assertEqual(len(normalized), 2)
        self.assertEqual(normalized[1][2], 'MCQ')
        self.assertEqual(normalized[1][3], 'What is Business Analytics?')
        self.assertEqual(normalized[1][6], 'Raw data only')
        self.assertEqual(normalized[1][9], 'Database storage')
