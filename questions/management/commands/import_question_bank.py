from __future__ import annotations

from pathlib import Path

from django.core.management.base import BaseCommand

from questions.importers import load_manifest, parse_workbook, workbook_matches
from questions.models import QuestionBankEntry


class Command(BaseCommand):
    help = 'Import assessment questions from Excel workbooks into the question bank table.'

    def add_arguments(self, parser):
        parser.add_argument('--root', default='aq_files')
        parser.add_argument('--manifest', default='aq_manifest.csv')
        parser.add_argument('--contains', nargs='*', default=[])
        parser.add_argument('--limit', type=int, default=None)
        parser.add_argument('--clear', action='store_true')

    def handle(self, *args, **options):
        root = Path(options['root'])
        manifest_path = Path(options['manifest'])
        manifest = load_manifest(manifest_path)

        files = sorted(root.glob('*.xlsx'))
        if options['contains']:
            files = [path for path in files if workbook_matches(path, options['contains'])]
        if options['limit'] is not None:
            files = files[: options['limit']]

        if options['clear']:
            deleted_count, _ = QuestionBankEntry.objects.all().delete()
            self.stdout.write(f'Cleared existing rows: {deleted_count}')

        records = []
        failures: list[tuple[str, str]] = []
        processed = 0

        for path in files:
            meta = manifest.get(path.name)
            try:
                records.extend(parse_workbook(path, meta))
                processed += 1
            except Exception as exc:  # noqa: BLE001
                failures.append((path.name, str(exc)))

        before_count = QuestionBankEntry.objects.count()
        batch = [QuestionBankEntry(**record) for record in records]
        QuestionBankEntry.objects.bulk_create(batch, batch_size=500, ignore_conflicts=True)
        after_count = QuestionBankEntry.objects.count()

        self.stdout.write(f'Processed workbooks: {processed}')
        self.stdout.write(f'Prepared question rows: {len(records)}')
        self.stdout.write(f'Question bank rows before import: {before_count}')
        self.stdout.write(f'Question bank rows after import: {after_count}')
        self.stdout.write(f'Inserted new rows: {after_count - before_count}')
        self.stdout.write(f'Failures: {len(failures)}')

        if failures:
            self.stdout.write('Failure details:')
            for name, message in failures[:20]:
                self.stdout.write(f'- {name}: {message}')
