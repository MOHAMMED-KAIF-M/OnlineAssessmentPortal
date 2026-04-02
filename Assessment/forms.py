from django import forms

from questions.models import QuestionBankEntry
from questions.question_types import QUESTION_TYPE_CHOICES

from .generation import build_question_bank_hierarchy, hierarchy_choices
from .models import Assessment


MULTISELECT_WIDGET_ATTRS = {
    'class': 'mt-2 w-full rounded-2xl border-slate-200 bg-white px-4 py-3 text-sm text-slate-900 shadow-sm focus:border-sky-500 focus:ring-sky-500',
    'size': 8,
}


class AssessmentGenerationForm(forms.Form):
    title = forms.CharField(
        max_length=200,
        widget=forms.TextInput(
            attrs={
                'class': 'mt-2 w-full rounded-2xl border-slate-200 bg-white px-4 py-3 text-sm text-slate-900 shadow-sm focus:border-sky-500 focus:ring-sky-500',
                'placeholder': 'Example: CDA Regression Assessment',
            }
        ),
    )
    description = forms.CharField(
        required=False,
        widget=forms.Textarea(
            attrs={
                'class': 'mt-2 w-full rounded-2xl border-slate-200 bg-white px-4 py-3 text-sm text-slate-900 shadow-sm focus:border-sky-500 focus:ring-sky-500',
                'rows': 4,
                'placeholder': 'Short description for students',
            }
        ),
    )
    course = forms.ChoiceField(
        choices=Assessment.COURSE_CHOICES,
        widget=forms.Select(
            attrs={
                'class': 'mt-2 w-full rounded-2xl border-slate-200 bg-white px-4 py-3 text-sm text-slate-900 shadow-sm focus:border-sky-500 focus:ring-sky-500',
            }
        ),
    )
    subcourses = forms.MultipleChoiceField(
        required=False,
        choices=(),
        widget=forms.SelectMultiple(attrs=MULTISELECT_WIDGET_ATTRS),
    )
    modules = forms.MultipleChoiceField(
        required=False,
        choices=(),
        widget=forms.SelectMultiple(attrs=MULTISELECT_WIDGET_ATTRS),
    )
    workbooks = forms.MultipleChoiceField(
        required=False,
        choices=(),
        widget=forms.SelectMultiple(attrs=MULTISELECT_WIDGET_ATTRS),
    )
    topics = forms.MultipleChoiceField(
        required=False,
        choices=(),
        widget=forms.SelectMultiple(attrs=MULTISELECT_WIDGET_ATTRS),
    )
    question_types = forms.MultipleChoiceField(
        choices=QUESTION_TYPE_CHOICES,
        widget=forms.CheckboxSelectMultiple,
    )
    difficulties = forms.MultipleChoiceField(
        required=False,
        choices=(),
        widget=forms.CheckboxSelectMultiple,
    )
    question_count = forms.IntegerField(
        min_value=1,
        max_value=100,
        widget=forms.NumberInput(
            attrs={
                'class': 'mt-2 w-full rounded-2xl border-slate-200 bg-white px-4 py-3 text-sm text-slate-900 shadow-sm focus:border-sky-500 focus:ring-sky-500',
            }
        ),
    )
    duration_minutes = forms.IntegerField(
        min_value=1,
        max_value=300,
        initial=30,
        widget=forms.NumberInput(
            attrs={
                'class': 'mt-2 w-full rounded-2xl border-slate-200 bg-white px-4 py-3 text-sm text-slate-900 shadow-sm focus:border-sky-500 focus:ring-sky-500',
            }
        ),
    )
    marks_per_question = forms.IntegerField(
        min_value=1,
        max_value=20,
        initial=1,
        widget=forms.NumberInput(
            attrs={
                'class': 'mt-2 w-full rounded-2xl border-slate-200 bg-white px-4 py-3 text-sm text-slate-900 shadow-sm focus:border-sky-500 focus:ring-sky-500',
            }
        ),
    )
    randomize = forms.BooleanField(required=False, initial=True)

    def __init__(self, *args, hierarchy=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.hierarchy = hierarchy or build_question_bank_hierarchy()
        selected_course = self._selected_course()
        selected_subcourses = self._selected_values('subcourses')
        selected_modules = self._selected_values('modules')
        selected_workbooks = self._selected_values('workbooks')
        options = hierarchy_choices(
            self.hierarchy,
            selected_course,
            selected_subcourses=selected_subcourses,
            selected_modules=selected_modules,
            selected_workbooks=selected_workbooks,
        )

        self.fields['subcourses'].choices = [(value, value) for value in options['subcourses']]
        self.fields['modules'].choices = [(value, value) for value in options['modules']]
        self.fields['workbooks'].choices = [(value, value) for value in options['workbooks']]
        self.fields['topics'].choices = [(value, value) for value in options['topics']]

        difficulties = sorted(
            {
                value.strip()
                for value in QuestionBankEntry.objects.values_list('difficulty', flat=True)
                if value and value.strip()
            },
            key=str.lower,
        )
        self.fields['difficulties'].choices = [(value, value) for value in difficulties]
        self.fields['question_types'].initial = ['MCQ']
        self.fields['question_count'].initial = 10

    def _selected_course(self):
        if self.is_bound:
            return self.data.get('course') or Assessment.COURSE_CHOICES[0][0]
        return (
            self.initial.get('course')
            or self.fields['course'].initial
            or Assessment.COURSE_CHOICES[0][0]
        )

    def _selected_values(self, field_name):
        if self.is_bound:
            return [value for value in self.data.getlist(field_name) if value]
        initial_value = self.initial.get(field_name, [])
        if isinstance(initial_value, (list, tuple)):
            return [value for value in initial_value if value]
        return [initial_value] if initial_value else []
