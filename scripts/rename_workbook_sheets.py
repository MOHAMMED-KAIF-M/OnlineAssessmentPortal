from __future__ import annotations

import argparse
import csv
from pathlib import Path

from openpyxl import load_workbook


def to_long_path(path: Path) -> str:
    resolved = path.resolve()
    return f"\\\\?\\{resolved}"


def workbook_targets_from_manifest(manifest_path: Path) -> list[Path]:
    targets: list[Path] = []
    seen: set[str] = set()

    with manifest_path.open(newline='', encoding='utf-8') as handle:
        for row in csv.DictReader(handle):
            original_path = Path(row['original_path'])
            extracted_path = Path(row['extracted_file'])
            for path in (original_path, extracted_path):
                normalized = str(path).replace('/', '\\').lower()
                if normalized not in seen:
                    seen.add(normalized)
                    targets.append(path)

    return targets


def desired_titles(sheet_count: int) -> list[str] | None:
    if sheet_count == 1:
        return ['questions']
    if sheet_count == 2:
        return ['dataset', 'questions']
    return None


def rename_workbook_sheets(path: Path, dry_run: bool) -> tuple[str, int]:
    workbook = load_workbook(to_long_path(path))
    try:
        target_titles = desired_titles(len(workbook.worksheets))
        if target_titles is None:
            return 'skipped', 0

        current_titles = [sheet.title for sheet in workbook.worksheets]
        if current_titles == target_titles:
            return 'unchanged', 0

        # Avoid duplicate-title collisions while reassigning sheet names.
        for index, worksheet in enumerate(workbook.worksheets, start=1):
            worksheet.title = f'__tmp_sheet_{index}__'

        for worksheet, title in zip(workbook.worksheets, target_titles):
            worksheet.title = title

        if not dry_run:
            workbook.save(to_long_path(path))
        return 'changed', len(target_titles)
    finally:
        workbook.close()


def main() -> None:
    parser = argparse.ArgumentParser(description='Rename workbook sheets to dataset/questions based on sheet count.')
    parser.add_argument('--manifest', default='aq_manifest.csv')
    parser.add_argument('--dry-run', action='store_true')
    args = parser.parse_args()

    manifest_path = Path(args.manifest)
    targets = workbook_targets_from_manifest(manifest_path)

    processed = 0
    changed = 0
    renamed_sheets = 0
    skipped = 0
    missing = 0
    failures: list[tuple[Path, str]] = []

    for path in targets:
        try:
            status, sheet_count = rename_workbook_sheets(path, dry_run=args.dry_run)
            processed += 1
            if status == 'changed':
                changed += 1
                renamed_sheets += sheet_count
            elif status == 'skipped':
                skipped += 1
        except FileNotFoundError:
            missing += 1
        except Exception as exc:  # noqa: BLE001
            failures.append((path, str(exc)))

    print(f'Processed targets: {processed}')
    print(f'Changed workbooks: {changed}')
    print(f'Renamed sheets: {renamed_sheets}')
    print(f'Skipped workbooks: {skipped}')
    print(f'Missing targets: {missing}')
    print(f'Failures: {len(failures)}')

    if failures:
        print('Failure details:')
        for path, message in failures[:20]:
            print(f'- {path}: {message}')


if __name__ == '__main__':
    main()
