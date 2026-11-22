"""Проверки бухгалтерского экспорта без Qt и сторонних библиотек.

Запуск: python -m unittest discover -s tests -p 'test_accounting_export.py' -v
Опционально: REINDEER_TEST_XLSX=... для сверки с реальной ведомостью.
"""
import hashlib
import os
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from accounting_export import (
    AccountingExportError, GROUPS, SHEET_NAMES, TEMPLATE_PATH, _sheet_paths,
    build_accounting_summary, compare_with_template, export_accounting_xlsx,
    page_boundaries_from_text, read_template_reference,
)
from ledger_core import Animal, Ledger, read_ledger


OLD_MARKED = {8,11,12,14,21,25,28,30,32,37,43,47,50,54,55,
              60,61,62,64,65,73,75,89,98,100,102}
TEMPLATE_MARKED = (OLD_MARKED - {25}) | {24}


def synthetic_ledger():
    return Ledger(Path('/tmp/input.xlsx'), 'тест', '01.01.2026', [
        Animal(1,'01','', '',None,10.,None,'тел.самец','Телёнок-самец',2026,5),
        Animal(2,'02','', '',None,None,20.,'важ','Важенка',2021,6),
        Animal(3,'03','', '',None,30.,None,'бык','Бык',2023,7),
        Animal(4,'04','', '',None,None,12.,'тел.самка','Телёнок-самка',2026,8),
    ], {}, [])


class TestAccountingExport(TestCase):
    def test_page_boundaries(self):
        self.assertEqual(page_boundaries_from_text('44, 94', 3), (44, 94))
        self.assertEqual(page_boundaries_from_text('2', 2), (2,))
        self.assertEqual(page_boundaries_from_text('', 1), ())
        for text, count in [('44',3), ('44,44',3), ('94,44',3), ('abc',2),('',2)]:
            with self.subTest(text=text,count=count), self.assertRaises(AccountingExportError):
                page_boundaries_from_text(text,count)

    def test_unique_assignment_and_category(self):
        l = synthetic_ledger()
        s = build_accounting_summary(l, {2,4}, (2,))
        self.assertEqual(s.count,4)
        self.assertEqual(s.purchased_count,2)
        self.assertEqual(s.purchased_weight,32)
        self.assertEqual(s.totals['Итого'], (2,40,2,32))
        self.assertEqual(s.totals['Страница 1'], (1,10,1,20))
        self.assertEqual(s.totals['Страница 2'], (1,30,1,12))
        self.assertEqual(s.sheets['Итого'][GROUPS.index('З')], [0,0,2,32])
        self.assertEqual(s.totals['Страница 3'], (0,0,0,0))
        with self.assertRaises(AccountingExportError):
            build_accounting_summary(l,{5},(2,))
        with self.assertRaises(AccountingExportError):
            build_accounting_summary(l,{1},(4,))

    def test_export_preserves_formulas_and_original_bytes(self):
        l=synthetic_ledger()
        summary=build_accounting_summary(l,{2,4},(2,))
        original_digest=hashlib.sha256(TEMPLATE_PATH.read_bytes()).hexdigest()
        with TemporaryDirectory() as td:
            output=Path(td)/'report.xlsx'
            export_accounting_xlsx(summary, output, source=l.path)
            self.assertTrue(output.is_file())
            with ZipFile(TEMPLATE_PATH) as zin, ZipFile(output) as zout:
                paths=_sheet_paths(zout)
                self.assertEqual(tuple(paths), SHEET_NAMES)
                for member in zin.namelist():
                    if member not in paths.values():
                        self.assertEqual(zin.read(member),zout.read(member))
                ns={'x':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                for name, path in paths.items():
                    root=ET.fromstring(zout.read(path))
                    cells={c.attrib['r']:c for c in root.findall('.//x:c',ns)}
                    for col,total in zip(('B','C','D','E'),summary.totals[name]):
                        cell=cells[f'{col}9']
                        self.assertEqual(float(cell.find('x:v',ns).text),float(total))
                        self.assertTrue(cell.find('x:f',ns).text.startswith('SUM('))
                    self.assertEqual(float(cells['B7'].find('x:v',ns).text),
                                     float(summary.sheets[name][4][0]))
            self.assertEqual(hashlib.sha256(TEMPLATE_PATH.read_bytes()).hexdigest(),original_digest)
            with self.assertRaises(AccountingExportError):
                export_accounting_xlsx(summary,l.path,source=l.path)
            with self.assertRaises(AccountingExportError):
                export_accounting_xlsx(summary,TEMPLATE_PATH)

    def test_real_18_11_if_available(self):
        source=os.environ.get('REINDEER_TEST_XLSX_18')
        if not source:
            self.skipTest('REINDEER_TEST_XLSX_18 не задан')
        ledger=read_ledger(Path(source))
        summary=build_accounting_summary(ledger,set(),())
        self.assertEqual((summary.count, summary.weight), (70, 1893))
        self.assertEqual(summary.totals['Итого'], (60, 1645, 10, 248))
        self.assertEqual(summary.totals['Страница 2'], (0, 0, 0, 0))

    def test_real_17_11_if_available(self):
        source=os.environ.get('REINDEER_TEST_XLSX')
        if not source:
            self.skipTest('REINDEER_TEST_XLSX не задан')
        ledger=read_ledger(Path(source))
        original=build_accounting_summary(ledger,OLD_MARKED,(44,94))
        self.assertEqual((original.count, original.weight, original.purchased_count,
                          original.purchased_weight), (102,2542,26,852))
        self.assertEqual(original.totals['Итого'],(62,1700,40,842))
        self.assertEqual(original.totals['Страница 1'],(25,644,19,378))
        self.assertEqual(original.totals['Страница 2'],(30,873,20,440))
        self.assertEqual(original.totals['Страница 3'],(7,183,1,24))
        corrected=build_accounting_summary(ledger,TEMPLATE_MARKED,(44,94))
        self.assertEqual((corrected.purchased_count,corrected.purchased_weight),(26,867))
        self.assertEqual(compare_with_template(corrected),[])
        self.assertGreater(len(compare_with_template(original)),0)
