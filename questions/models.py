from django.db import models

from .question_types import ALLOWED_QUESTION_TYPES, QUESTION_TYPE_CHOICES


class QuestionBankEntry(models.Model):
    source_workbook_id = models.PositiveIntegerField(db_index=True, null=True, blank=True)
    source_file = models.CharField(max_length=64, db_index=True)
    source_path = models.TextField(blank=True)
    course_code = models.CharField(max_length=3, blank=True, db_index=True)
    module_path = models.TextField(blank=True)
    workbook_title = models.CharField(max_length=255)
    sheet_name = models.CharField(max_length=128, blank=True)
    row_number = models.PositiveIntegerField()
    topic = models.CharField(max_length=255, blank=True)
    question_type = models.CharField(max_length=64, blank=True, choices=QUESTION_TYPE_CHOICES)
    difficulty = models.CharField(max_length=64, blank=True)
    question_text = models.TextField()
    context = models.TextField(blank=True)
    details = models.TextField(blank=True)
    option_a = models.TextField(blank=True)
    option_b = models.TextField(blank=True)
    option_c = models.TextField(blank=True)
    option_d = models.TextField(blank=True)
    correct_answer = models.TextField(blank=True)
    raw_payload = models.TextField(blank=True)
    imported_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['course_code', 'source_file', 'row_number']
        constraints = [
            models.UniqueConstraint(
                fields=['source_file', 'sheet_name', 'row_number'],
                name='unique_question_bank_entry_row',
            ),
            models.CheckConstraint(
                check=models.Q(question_type='') | models.Q(question_type__in=ALLOWED_QUESTION_TYPES),
                name='question_bank_entry_valid_type',
            ),
        ]

    def __str__(self):
        return f"{self.course_code or 'UNK'} | {self.workbook_title} | Row {self.row_number}"
