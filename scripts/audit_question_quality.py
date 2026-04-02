from __future__ import annotations

import argparse
import csv
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from questions.importers import load_manifest, normalize_key, normalize_text
from questions.question_types import normalize_question_type
from scripts.rebuild_question_sheets import extract_existing_questions, needs_rebuild


QUESTION_ROW_LIMIT = 5


@dataclass
class WorkbookAudit:
    source_file: str
    source_path: str
    workbook_title: str
    mode: str
    question_count: int
    is_proper: bool
    issue_counts: Counter
    issue_examples: list[str]


def question_issue_keys(question_text: str, question_type: str, options: list[str]) -> list[str]:
    issues: list[str] = []
    normalized_question = normalize_text(question_text)
    normalized_type = normalize_text(question_type)
    option_count = sum(bool(normalize_text(option)) for option in options)

    if not normalized_question:
        issues.append("blank_question")
    if normalized_type and not normalize_question_type(normalized_type):
        issues.append("unsupported_type_row")
    if normalized_question.startswith("Section,Type,Q.No"):
        issues.append("header_leaked_into_question_rows")
    if normalized_question.startswith("SECTION "):
        issues.append("packed_question_row")
    if normalized_question.startswith("SECTION C: Dataset,Dataset"):
        issues.append("dataset_row_in_questions_sheet")
    if normalized_type == "MCQ" and option_count < 4:
        issues.append("mcq_missing_split_options")
    if normalized_type == "Scenario-Based" and option_count not in {0, 4}:
        issues.append("scenario_partial_options")
    return issues


def audit_workbook(path: Path, manifest_meta) -> WorkbookAudit:
    mode, rows = extract_existing_questions(path)
    rebuild_required, rebuild_reasons = needs_rebuild(mode, rows, path)

    issue_counts: Counter[str] = Counter(rebuild_reasons)
    issue_examples: list[str] = []

    duplicate_counter = Counter(
        normalize_key(row.question_text)
        for row in rows
        if normalize_text(row.question_text)
    )
    duplicate_questions = {
        key for key, count in duplicate_counter.items() if key and count > 1
    }
    if duplicate_questions:
        issue_counts["duplicate_question_text"] += len(duplicate_questions)
        for row in rows:
            question_key = normalize_key(row.question_text)
            if question_key in duplicate_questions:
                issue_examples.append(
                    f"duplicate_question_text: {normalize_text(row.question_text)[:160]}"
                )
                if len(issue_examples) >= QUESTION_ROW_LIMIT:
                    break

    for row in rows:
        row_issues = question_issue_keys(
            row.question_text,
            row.question_type,
            [row.option_a, row.option_b, row.option_c, row.option_d],
        )
        for issue in row_issues:
            issue_counts[issue] += 1
            if len(issue_examples) < QUESTION_ROW_LIMIT:
                issue_examples.append(
                    f"{issue}: {normalize_text(row.question_text)[:160]}"
                )

    is_proper = not issue_counts
    return WorkbookAudit(
        source_file=path.name,
        source_path=manifest_meta.source_path if manifest_meta else str(path),
        workbook_title=manifest_meta.workbook_title if manifest_meta else path.stem,
        mode=mode,
        question_count=len(rows),
        is_proper=is_proper,
        issue_counts=issue_counts,
        issue_examples=issue_examples,
    )


def write_csv_report(path: Path, audits: list[WorkbookAudit]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            [
                "source_file",
                "source_path",
                "workbook_title",
                "mode",
                "question_count",
                "status",
                "issue_summary",
                "issue_examples",
            ]
        )
        for audit in audits:
            issue_summary = "; ".join(
                f"{key}={value}" for key, value in audit.issue_counts.most_common()
            )
            writer.writerow(
                [
                    audit.source_file,
                    audit.source_path,
                    audit.workbook_title,
                    audit.mode,
                    audit.question_count,
                    "proper" if audit.is_proper else "not_proper",
                    issue_summary,
                    " | ".join(audit.issue_examples),
                ]
            )


def write_markdown_report(path: Path, audits: list[WorkbookAudit], totals: Counter, top_issues: Counter) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    proper = sum(1 for audit in audits if audit.is_proper)
    not_proper = len(audits) - proper
    worst = sorted(
        (audit for audit in audits if not audit.is_proper),
        key=lambda audit: (-sum(audit.issue_counts.values()), audit.source_file),
    )[:25]

    lines = [
        "# Question Quality Report",
        "",
        f"- Workbooks scanned: {len(audits)}",
        f"- Proper workbooks: {proper}",
        f"- Not proper workbooks: {not_proper}",
        "",
        "## Audit Rules",
        "",
        "- `needs_rebuild(...)` flags: wrong question count, wrong raw row count, non-structured sheet, blank/unsupported/nonstandard type, blank question",
        "- Question-level flags: MCQ rows without 4 split options, packed question rows, dataset rows inside question sheets, leaked header rows, duplicate question text",
        "",
        "## Top Issues",
        "",
    ]

    for issue, count in top_issues.most_common():
        workbook_count = totals[issue]
        lines.append(f"- `{issue}`: {count} occurrences across {workbook_count} workbooks")

    lines.extend(
        [
            "",
            "## Worst Affected Workbooks",
            "",
            "| Status | Source File | Workbook Title | Question Count | Issues |",
            "|---|---|---|---:|---|",
        ]
    )
    for audit in worst:
        summary = ", ".join(
            f"{key}={value}" for key, value in audit.issue_counts.most_common(5)
        )
        lines.append(
            f"| not_proper | {audit.source_file} | {audit.workbook_title} | {audit.question_count} | {summary} |"
        )

    lines.extend(["", "## Sample Problem Rows", ""])
    for audit in worst[:15]:
        if not audit.issue_examples:
            continue
        lines.append(f"### {audit.source_file} - {audit.workbook_title}")
        for example in audit.issue_examples[:QUESTION_ROW_LIMIT]:
            lines.append(f"- {example}")
        lines.append("")

    path.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Audit question quality across workbook files.")
    parser.add_argument("--root", default="aq_files")
    parser.add_argument("--manifest", default="aq_manifest.csv")
    parser.add_argument("--csv-out", default="reports/question_quality_report.csv")
    parser.add_argument("--md-out", default="reports/question_quality_report.md")
    args = parser.parse_args()

    root = Path(args.root)
    manifest = load_manifest(Path(args.manifest))

    audits: list[WorkbookAudit] = []
    issue_workbook_totals: Counter[str] = Counter()
    issue_occurrence_totals: Counter[str] = Counter()

    for path in sorted(root.glob("*.xlsx")):
        audit = audit_workbook(path, manifest.get(path.name))
        audits.append(audit)
        for issue, count in audit.issue_counts.items():
            issue_occurrence_totals[issue] += count
            issue_workbook_totals[issue] += 1

    write_csv_report(Path(args.csv_out), audits)
    write_markdown_report(
        Path(args.md_out),
        audits,
        issue_workbook_totals,
        issue_occurrence_totals,
    )

    proper = sum(1 for audit in audits if audit.is_proper)
    not_proper = len(audits) - proper
    print(f"Scanned: {len(audits)}")
    print(f"Proper: {proper}")
    print(f"Not proper: {not_proper}")
    print("Top issues:")
    for issue, count in issue_occurrence_totals.most_common(10):
        print(f"- {issue}: {count}")
    print(f"CSV report: {args.csv_out}")
    print(f"Markdown report: {args.md_out}")


if __name__ == "__main__":
    main()
