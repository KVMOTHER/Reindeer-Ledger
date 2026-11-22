"""Reindeer Ledger 0.5.3: бухгалтерская сводка в исходном Excel-шаблоне.

Только стандартная библиотека. Исходный шаблон и XLSX пользователя не меняются.
Копируются все ZIP-части шаблона; числовые ячейки B3:E7 и кэш формул
B9:E9 обновляются. Стили, объединения и формулы остаются как в образце.
"""
from __future__ import annotations

from bisect import bisect_left
from dataclasses import dataclass
from math import isfinite, isclose
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Iterable
from xml.etree import ElementTree as ET
from zipfile import ZipFile
import os
import re

from ledger_core import Ledger, short_weight, summarize

TEMPLATE_PATH = Path(__file__).resolve().parent / 'assets' / 'accounting_template.xlsx'
SHEET_NAMES = ('Итого', 'Страница 1', 'Страница 2', 'Страница 3')
GROUPS = ('ТС', 'ТВ', 'В', 'Б', 'З')
GROUP_CATEGORY = {
    'Телёнок-самец': 'ТС',
    'Телёнок-самка': 'ТВ',
    'Важенка': 'В',
    'Бык': 'Б',
}
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
RELSNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
ET.register_namespace('', NS)
ET.register_namespace('r', RNS)
ROW_START = 3
DATA_COLS = ('B', 'C', 'D', 'E')


class AccountingExportError(ValueError):
    """Источник, разметка, страницы или шаблон не прошли проверку."""


@dataclass(frozen=True)
class AccountingSummary:
    """Матрицы [группа][I кол-во, I вес, II кол-во, II вес]."""
    sheets: dict[str, list[list[int | float]]]
    totals: dict[str, tuple[int | float, int | float, int | float, int | float]]
    page_limits: tuple[int, ...]
    count: int
    weight: float
    purchased_count: int
    purchased_weight: float


def page_boundaries_from_text(text: str, page_count: int) -> tuple[int, ...]:
    if page_count not in (1, 2, 3):
        raise AccountingExportError('Шаблон допускает от 1 до 3 страниц ведомости.')
    cleaned = text.strip()
    if not cleaned:
        if page_count == 1:
            return ()
        raise AccountingExportError(
            'Укажите последний № п/п на каждой странице, кроме последней. '
            'Пример для трёх страниц: 44, 94.')
    tokens = re.split(r'[,;\s]+', cleaned)
    if len(tokens) != page_count - 1 or not all(x.isdigit() for x in tokens):
        raise AccountingExportError(
            f'Для {page_count} стр. нужно указать {page_count-1} границ(ы), '
            'целыми номерами через запятую.')
    result = tuple(int(x) for x in tokens)
    if list(result) != sorted(set(result)):
        raise AccountingExportError('Границы страниц должны идти строго по возрастанию.')
    return result


def _zero_matrix() -> list[list[int | float]]:
    return [[0, 0, 0, 0] for _ in GROUPS]


def _subtotal(matrix: list[list[int | float]]) -> tuple[int | float, ...]:
    return tuple(sum(row[col] for row in matrix) for col in range(4))


def build_accounting_summary(
    ledger: Ledger, purchased: set[int], page_limits: Iterable[int]
) -> AccountingSummary:
    """Без догадок и дублирования: каждая строка — ровно одна ячейка веса."""
    animals = ledger.animals
    if not animals:
        raise AccountingExportError('Ведомость не содержит животных.')
    nums = {a.number for a in animals}
    if len(nums) != len(animals):
        raise AccountingExportError('В исходной ведомости повторяются № п/п.')
    if not purchased <= nums:
        raise AccountingExportError(f'Неизвестные закупные №: {sorted(purchased-nums)}')
    cutoffs = tuple(page_limits)
    if len(cutoffs) > 2 or tuple(sorted(set(cutoffs))) != cutoffs:
        raise AccountingExportError('Нужно не более двух возрастающих границ страниц.')
    if cutoffs and (min(cutoffs) <= min(nums) or max(cutoffs) >= max(nums)):
        raise AccountingExportError('Границы страниц должны быть внутри диапазона № п/п.')
    # Граница должна заканчиваться существующим животным, а не служебной строкой Excel.
    if any(limit not in nums for limit in cutoffs):
        raise AccountingExportError('Граница страницы должна совпадать с № п/п животного.')

    matrices = {name: _zero_matrix() for name in SHEET_NAMES}
    pages_used = [0] * (len(cutoffs) + 1)
    for a in animals:
        if a.category not in GROUP_CATEGORY:
            raise AccountingExportError(f'№ {a.number}: неизвестная категория «{a.category}».')
        has_free, has_industrial = a.free_weight is not None, a.industrial_weight is not None
        if has_free == has_industrial:
            raise AccountingExportError(f'№ {a.number}: должно быть одно направление и один вес.')
        mass = a.free_weight if has_free else a.industrial_weight
        if mass is None or not isfinite(mass) or mass <= 0:
            raise AccountingExportError(f'№ {a.number}: недопустимый вес {mass!r}.')
        group = 'З' if a.number in purchased else GROUP_CATEGORY[a.category]
        row = GROUPS.index(group)
        weight_col = 1 if has_free else 3
        count_col = weight_col - 1
        index = bisect_left(cutoffs, a.number)
        pages_used[index] += 1
        page_matrix = matrices[SHEET_NAMES[1 + index]]
        for matrix in (matrices['Итого'], page_matrix):
            matrix[row][count_col] += 1
            matrix[row][weight_col] += mass

    if any(c == 0 for c in pages_used):
        raise AccountingExportError('Одна из выбранных страниц пустая. Проверьте границы.')

    totals = {name: _subtotal(matrix) for name, matrix in matrices.items()}
    all_totals = totals['Итого']
    if all_totals[0] + all_totals[2] != len(animals):
        raise AccountingExportError('Проверка числа голов не пройдена.')
    actual = summarize(animals)
    for idx, key in ((0, 'free_count'), (1, 'free_weight'),
                     (2, 'industrial_count'), (3, 'industrial_weight')):
        if not isclose(all_totals[idx], actual[key], abs_tol=1e-7):
            raise AccountingExportError(f'Не совпадает общий итог: {key}.')
        if not isclose(sum(totals[name][idx] for name in SHEET_NAMES[1:]), all_totals[idx], abs_tol=1e-7):
            raise AccountingExportError('Суммы страниц не сходятся с листом «Итого».')
    for key, count_key, weight_key in (('free', 'free_count', 'free_weight'),
                                        ('industrial', 'industrial_count', 'industrial_weight')):
        if key in ledger.footer:
            expected_count, expected_weight = ledger.footer[key]
            if expected_count != actual[count_key] or not isclose(expected_weight, actual[weight_key], abs_tol=1e-7):
                raise AccountingExportError('Итоги строк отличаются от контрольных итогов исходного XLSX.')
    bought = [a for a in animals if a.number in purchased]
    purchased_sum = summarize(bought)
    bought_table = matrices['Итого'][GROUPS.index('З')]
    if ((bought_table[0] + bought_table[2] != len(purchased)) or
            not isclose(bought_table[1] + bought_table[3], purchased_sum['slaughter_weight'], abs_tol=1e-7)):
        raise AccountingExportError('Проверка количества или веса закупных не пройдена.')
    return AccountingSummary(matrices, totals, cutoffs, len(animals),
                             actual['slaughter_weight'], len(purchased),
                             purchased_sum['slaughter_weight'])


def _sheet_paths(zf: ZipFile) -> dict[str, str]:
    try:
        root = ET.fromstring(zf.read('xl/workbook.xml'))
        relroot = ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
        by_id = {r.get('Id'): r.get('Target') for r in relroot.findall(f'{{{RELSNS}}}Relationship')}
        result: dict[str, str] = {}
        from posixpath import normpath
        sheets = root.find(f'{{{NS}}}sheets')
        if sheets is None:
            raise AccountingExportError('В шаблоне нет листов.')
        for sheet in sheets:
            name = sheet.get('name')
            target = by_id.get(sheet.get(f'{{{RNS}}}id'))
            if not target:
                raise AccountingExportError(f'Нет XML для листа {name}.')
            target = target.lstrip('/')
            result[name] = normpath(target if target.startswith('xl/') else 'xl/' + target)
        if tuple(result) != SHEET_NAMES:
            raise AccountingExportError('Листы шаблона изменены или отсутствуют.')
        return result
    except (KeyError, ET.ParseError, TypeError, AttributeError) as exc:
        raise AccountingExportError('Повреждён Excel-шаблон бухгалтерской ведомости.') from exc


def _cell_text(cell: ET.Element) -> str:
    v = cell.find(f'{{{NS}}}v')
    return v.text or '' if v is not None else ''


def _format_number(value: float | int) -> str:
    return str(int(value)) if isclose(value, round(value), abs_tol=1e-8) else format(value, '.12g')


def _patch_sheet(original: bytes, matrix: list[list[int | float]],
                 totals: tuple[int | float, ...]) -> bytes:
    try:
        root = ET.fromstring(original)
        cells = {c.get('r'): c for c in root.iter(f'{{{NS}}}c')}
        for row_no, group in enumerate(GROUPS, ROW_START):
            label_cell = cells.get(f'A{row_no}')
            if label_cell is None or _cell_text(label_cell) != group:
                raise AccountingExportError('В шаблоне изменены строки групп ТС, ТВ, В, Б, З.')
        for idx, row in enumerate(matrix, ROW_START):
            for col, value in zip(DATA_COLS, row):
                address = f'{col}{idx}'
                cell = cells.get(address)
                if cell is None or cell.get('t','n') not in ('n',''):
                    raise AccountingExportError(f'В шаблоне нет числовой ячейки {address}.')
                v = cell.find(f'{{{NS}}}v')
                if v is None:
                    v = ET.SubElement(cell, f'{{{NS}}}v')
                v.text = _format_number(value)
        for col, total in zip(DATA_COLS, totals):
            address = f'{col}9'
            cell = cells.get(address)
            if cell is None or cell.find(f'{{{NS}}}f') is None:
                raise AccountingExportError(f'В шаблоне отсутствует формула {address}.')
            v = cell.find(f'{{{NS}}}v')
            if v is None:
                v = ET.SubElement(cell, f'{{{NS}}}v')
            v.text = _format_number(total)
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    except ET.ParseError as exc:
        raise AccountingExportError('Повреждена XML-разметка шаблона.') from exc


def read_template_reference(template: Path = TEMPLATE_PATH) -> dict[str, list[list[int | float]]]:
    """Числовые значения образца для *сравнения*, а не автоматической подмены."""
    matrices = {}
    with ZipFile(template, 'r') as zf:
        paths = _sheet_paths(zf)
        for name, path in paths.items():
            root = ET.fromstring(zf.read(path))
            cells = {c.get('r'): c for c in root.iter(f'{{{NS}}}c')}
            matrix = []
            for row in range(3,8):
                matrix.append([float(_cell_text(cells[f'{col}{row}'])) for col in DATA_COLS])
            matrices[name] = matrix
    return matrices


def compare_with_template(summary: AccountingSummary, template: Path = TEMPLATE_PATH) -> list[str]:
    """Возвращает все отличающиеся ячейки; шаблон — образец, не эталон закупа."""
    reference = read_template_reference(template)
    changes = []
    for name, matrix in summary.sheets.items():
        for idx, (actual, target) in enumerate(zip(matrix, reference[name])):
            for colidx, (current, expected) in enumerate(zip(actual, target)):
                if not isclose(current, expected, abs_tol=1e-7):
                    changes.append(f'{name}: {GROUPS[idx]} {DATA_COLS[colidx]}{idx+3}: '
                                   f'{_format_number(expected)} → {_format_number(current)}')
    return changes


def export_accounting_xlsx(summary: AccountingSummary, destination: Path,
                           template: Path = TEMPLATE_PATH, source: Path | None = None) -> Path:
    """Экспорт с атомарной заменой. В оригинальном шаблоне ничего не правим."""
    destination, template = Path(destination), Path(template)
    if destination.suffix.lower() != '.xlsx':
        raise AccountingExportError('Файл отчёта должен иметь расширение .xlsx.')
    output = destination.resolve()
    if output == template.resolve() or (source is not None and output == Path(source).resolve()):
        raise AccountingExportError('Нельзя перезаписывать исходный XLSX или файл шаблона.')
    if not template.is_file():
        raise AccountingExportError(f'Не найден бухгалтерский шаблон: {template}')
    destination.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = None
    try:
        with NamedTemporaryFile(prefix='.reindeer-', suffix='.xlsx', dir=destination.parent, delete=False) as temp:
            tmp_path = Path(temp.name)
        with ZipFile(template, 'r') as zin, ZipFile(tmp_path, 'w') as zout:
            if zin.testzip() is not None:
                raise AccountingExportError('Бухгалтерский шаблон повреждён.')
            paths = _sheet_paths(zin)
            replacements = {paths[name]: _patch_sheet(zin.read(paths[name]), summary.sheets[name], summary.totals[name])
                            for name in SHEET_NAMES}
            for item in zin.infolist():
                content = replacements.get(item.filename)
                zout.writestr(item, content if content is not None else zin.read(item.filename))
        with ZipFile(tmp_path, 'r') as verify:
            if verify.testzip() is not None:
                raise AccountingExportError('Сформированный XLSX повреждён.')
        os.replace(tmp_path, destination)
        return destination
    except (OSError, ET.ParseError) as exc:
        raise AccountingExportError(f'Не удалось сохранить Excel: {exc}') from exc
    finally:
        if tmp_path is not None and tmp_path.exists():
            tmp_path.unlink()
