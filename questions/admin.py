from django.contrib import admin

from .models import QuestionBankEntry


@admin.register(QuestionBankEntry)
class QuestionBankEntryAdmin(admin.ModelAdmin):
    exclude = (
        'correct_answer',
        'raw_payload',
    )
    list_display = (
        'id',
        'course_code',
        'source_file',
        'workbook_title',
        'question_type',
        'difficulty',
    )
    list_filter = ('course_code', 'question_type', 'difficulty')
    search_fields = (
        'source_file',
        'workbook_title',
        'question_text',
        'topic',
        'source_path',
    )
