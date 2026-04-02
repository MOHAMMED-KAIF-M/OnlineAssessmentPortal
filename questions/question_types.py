from __future__ import annotations

import re


ALLOWED_QUESTION_TYPES = (
    "Theory",
    "MCQ",
    "Coding",
    "SQL",
    "Practical",
    "Scenario-Based",
    "Short Answer",
)

QUESTION_TYPE_CHOICES = [(value, value) for value in ALLOWED_QUESTION_TYPES]

QUESTION_TYPE_ALIASES = {
    "application": "Practical",
    "coding": "Coding",
    "command": "Practical",
    "concept": "Theory",
    "conceptual": "Theory",
    "design": "Practical",
    "mcq": "MCQ",
    "one word": "Short Answer",
    "one word/command": "Short Answer",
    "one-word": "Short Answer",
    "practical": "Practical",
    "project": "Practical",
    "scenario": "Scenario-Based",
    "scenario based": "Scenario-Based",
    "scenario-based": "Scenario-Based",
    "short answer": "Short Answer",
    "short-answer": "Short Answer",
    "size": "Short Answer",
    "sql": "SQL",
    "technical": "Theory",
    "theory": "Theory",
    "two words": "Short Answer",
    "word": "Short Answer",
    "write yaml": "Coding",
}


def normalize_question_type_text(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def normalize_question_type_key(value: object) -> str:
    return normalize_question_type_text(value).lower()


def normalize_question_type(value: object) -> str:
    key = normalize_question_type_key(value)
    if not key:
        return ""
    return QUESTION_TYPE_ALIASES.get(key, "")


def is_allowed_question_type(value: object) -> bool:
    return normalize_question_type_text(value) in ALLOWED_QUESTION_TYPES
