#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Reindeer Ledger: локальный подсчёт XLSX, ручная разметка закупных, PDF и JSONL.

Требуется Python 3.10+. Для экспорта PDF: pip install reportlab.
Чтение .xlsx построено на стандартных библиотеках Python, исходник не изменяется.
"""
from __future__ import annotations

import argparse
import json
import math
import re
import sys
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZipFile

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
XREL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
PKG = '{http://schemas.openxmlformats.org/package/2006/relationships}'
CELL = re.compile(r'([A-Za-z]+)(\d+)$')
NUM = re.compile(r'^\d+(?:\.\d+)?$')
DATE = re.compile(r'\b(\d{1,2})\.(\d{1,2})\.(\d{2,4})\b')
WEIGHT_FOOTER = re.compile(r'(\d+)\s*гол[^\d]*?(\d+)', re.I)


def col_number(address: str) -> int:
    result = 0
    for ch in address.upper():
        if not 'A' <= ch <= 'Z':
            break
        result = result * 26 + ord(ch) - ord('A') + 1
    return result


def parse_cell_value(cell: ET.Element, shared: list[str]) -> Any:
    kind = cell.get('t', 'n')
    if kind == 'inlineStr':
        return ''.join(t.text or '' for t in cell.iter(NS + 't'))
    value = cell.find(NS + 'v')
    if kind == 's':
        return shared[int(value.text)] if value is not None and value.text else ''
    if kind in ('str', 'e'):
        return value.text if value is not None else None
    if kind == 'b':
        return (value is not None and value.text == '1')
    if value is None or value.text is None:
        return None
    text = value.text.strip()
    if not NUM.fullmatch(text):
        return text
    if text.isdigit():
        return int(text)  # сохраняем длинные федеральные номера без округления
    number = float(text)
    return int(number) if number.is_integer() else number


def read_xlsx_cells(path: Path) -> tuple[str, dict[int, dict[int, Any]]]:
    """Read first worksheet using ZIP + SpreadsheetML (no third-party XLSX libs)."""
    with ZipFile(path) as zf:
        shared: list[str] = []
        if 'xl/sharedStrings.xml' in zf.namelist():
            root = ET.fromstring(zf.read('xl/sharedStrings.xml'))
            for si in root.findall(NS + 'si'):
                shared.append(''.join(t.text or '' for t in si.iter(NS + 't')))
        workbook = ET.fromstring(zf.read('xl/workbook.xml'))
        sheets = workbook.find(NS + 'sheets')
        if sheets is None or not list(sheets):
            raise ValueError('В книге нет листов')
        first = list(sheets)[0]
        sheet_name = first.get('name', 'Лист 1')
        relationship_id = first.get(XREL + 'id')
        relroot = ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
        target = None
        for rel in relroot.findall(PKG + 'Relationship'):
            if rel.get('Id') == relationship_id:
                target = rel.get('Target')
                break
        if not target:
            raise ValueError('Не удалось определить XML первого листа')
        target = target.lstrip('/')
        if not target.startswith('xl/'):
            target = 'xl/' + target
        from posixpath import normpath
        target = normpath(target)
        root = ET.fromstring(zf.read(target))
        rows: dict[int, dict[int, Any]] = {}
        for r in root.iter(NS + 'row'):
            row_id = int(r.get('r', '0'))
            cells = {}
            for c in r.findall(NS + 'c'):
                address = c.get('r', '')
                match = CELL.match(address)
                if not match:
                    continue
                value = parse_cell_value(c, shared)
                if value is not None and value != '':
                    cells[col_number(match.group(1))] = value
            rows[row_id] = cells
        return sheet_name, rows


def string(value: Any) -> str:
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def weight(value: Any) -> float | None:
    if value in (None, ''):
        return None
    if isinstance(value, (float, int)):
        return float(value)
    s = str(value).strip().replace(',', '.')
    if NUM.fullmatch(s):
        return float(s)
    raise ValueError(f'Нечисловой вес: {value!r}')


def short_weight(value: float | None) -> str:
    if value is None:
        return '—'
    return str(int(value)) if value.is_integer() else f'{value:g}'


def normalize_category(raw: str) -> str:
    cleaned = raw.lower().replace('ё','е').replace(' ', '').strip('. ')
    mapping = {
        'тв': 'Телёнок-самка', 'тел.самка': 'Телёнок-самка',
        'тел.самец': 'Телёнок-самец', 'тб': 'Телёнок-самец',
        'важ': 'Важенка', 'важенка': 'Важенка',
        'бык': 'Бык',
    }
    return mapping.get(cleaned, raw or 'Не указана')


@dataclass(frozen=True)
class Animal:
    number: int
    fed_number: str
    tag: str
    internal_number: str
    live_weight: float | None
    free_weight: float | None
    industrial_weight: float | None
    category_raw: str
    category: str
    birth_year: int | None
    excel_row: int

    @property
    def slaughter_weight(self) -> float:
        return (self.free_weight or 0) + (self.industrial_weight or 0)

    @property
    def operation(self) -> str:
        if self.free_weight is not None and self.industrial_weight is not None:
            return 'ОШИБКА: оба направления'
        if self.free_weight is not None:
            return 'Свободная реализация'
        if self.industrial_weight is not None:
            return 'Промышленная переработка'
        return 'Нет направления'


@dataclass
class Ledger:
    path: Path
    sheet: str
    date: str
    animals: list[Animal]
    footer: dict[str, tuple[int, float]]
    warnings: list[str]
    source_pdf: str | None = None


def read_ledger(path: Path) -> Ledger:
    sheet, rows = read_xlsx_cells(path)
    headers = rows.get(4, {})
    if 'фед' not in string(headers.get(2)).lower() or 'свобод' not in string(headers.get(6)).lower():
        raise ValueError('Не распознана структура ведомости (нужны колонки A:I и заголовки в строке 4).')
    m = DATE.search(string(rows.get(2, {}).get(1)))
    date = ''
    if m:
        d, mon, yy = map(int, m.groups())
        yy = yy + 2000 if yy < 100 else yy
        date = f'{d:02d}.{mon:02d}.{yy}'
    else:
        date = 'неизвестная дата'
    warnings = []
    animals = []
    footer = {}
    for rownum, cells in sorted(rows.items()):
        serial = cells.get(1)
        if rownum >= 5 and isinstance(serial, int) and serial > 0:
            try:
                free = weight(cells.get(6))
                industrial = weight(cells.get(7))
                live = weight(cells.get(5))
            except ValueError as exc:
                raise ValueError(f'Строка Excel {rownum}: {exc}') from exc
            raw = string(cells.get(8))
            try:
                year = int(cells[9]) if cells.get(9) is not None else None
            except (ValueError, TypeError):
                year = None
                warnings.append(f'№ {serial}: некорректный год рождения')
            item = Animal(
                number=serial,
                fed_number=string(cells.get(2)), tag=string(cells.get(3)),
                internal_number=string(cells.get(4)), live_weight=live,
                free_weight=free, industrial_weight=industrial,
                category_raw=raw, category=normalize_category(raw),
                birth_year=year, excel_row=rownum,
            )
            if free is not None and industrial is not None:
                warnings.append(f'№ {serial}: указаны оба направления!')
            if free is None and industrial is None:
                warnings.append(f'№ {serial}: не указан убойный вес')
            animals.append(item)
        if rownum >= 5:
            descriptor = string(cells.get(5)).lower()
            summary = string(cells.get(8))
            fm = WEIGHT_FOOTER.search(summary)
            if fm:
                if 'свобод' in descriptor:
                    footer['free'] = (int(fm.group(1)), float(fm.group(2)))
                if 'промыш' in descriptor:
                    footer['industrial'] = (int(fm.group(1)), float(fm.group(2)))
    if not animals:
        raise ValueError('Ни одного животного в таблице не найдено.')
    nums = [a.number for a in animals]
    if len(set(nums)) != len(nums):
        raise ValueError('Повторяются номера строк животных: нужна ручная проверка файла.')
    if nums != list(range(1, len(nums)+1)):
        warnings.append('Непоследовательные порядковые номера животных.')
    if not any(a.live_weight is not None for a in animals):
        warnings.append('В XLSX нет живого веса; сумма живого веса не рассчитывается.')
    if date != 'неизвестная дата' and date not in sheet:
        warnings.append(f'Название листа «{sheet}» не совпадает с датой документа {date}. Используется дата из шапки.')
    ledger = Ledger(path, sheet, date, animals, footer, warnings)
    count_free = sum(a.free_weight is not None for a in animals)
    count_ind = sum(a.industrial_weight is not None for a in animals)
    mass_free = sum(a.free_weight or 0 for a in animals)
    mass_ind = sum(a.industrial_weight or 0 for a in animals)
    for label, count, mass, key in [('Свободная реализация', count_free, mass_free, 'free'), ('Промышленная переработка', count_ind, mass_ind, 'industrial')]:
        if key in footer:
            exp_count, exp_mass = footer[key]
            if count != exp_count or abs(mass-exp_mass)>0.001:
                ledger.warnings.append(f'{label}: строки {count} гол./{mass:g} кг; напечатано {exp_count} гол./{exp_mass:g} кг!')
    return ledger


def summarize(animals: list[Animal]) -> dict[str, Any]:
    free = [a for a in animals if a.free_weight is not None]
    industrial = [a for a in animals if a.industrial_weight is not None]
    known_live = [a.live_weight for a in animals if a.live_weight is not None]
    return {
        'count': len(animals),
        'free_count': len(free), 'free_weight': sum(a.free_weight or 0 for a in free),
        'industrial_count': len(industrial), 'industrial_weight': sum(a.industrial_weight or 0 for a in industrial),
        'slaughter_weight': sum(a.slaughter_weight for a in animals),
        'live_known_count': len(known_live),
        'live_known_sum': sum(known_live),
        'by_category': dict(sorted(Counter(a.category for a in animals).items())),
    }


def print_summary(title: str, animals: list[Animal]) -> None:
    s = summarize(animals)
    print(f'\n{title}: {s["count"]} голов')
    print(f'  Свободная реализация: {s["free_count"]} гол., {short_weight(float(s["free_weight"]))} кг')
    print(f'  Промышленная переработка: {s["industrial_count"]} гол., {short_weight(float(s["industrial_weight"]))} кг')
    print(f'  Убойный вес: {short_weight(float(s["slaughter_weight"]))} кг')
    if s['live_known_count'] == s['count']:
        print(f'  Живой вес: {short_weight(float(s["live_known_sum"]))} кг')
    else:
        print(f'  Живой вес: нет полного набора значений ({s["live_known_count"]}/{s["count"]})')
    print('  Категории: ' + ', '.join(f'{name} — {count}' for name, count in s['by_category'].items()))


def print_animals(animals: list[Animal]) -> None:
    print('\n№ | Федеральный №           | Бирка     | Категория       | Направление | Вес, кг')
    print('-'*92)
    for a in animals:
        ch = 'Р' if a.free_weight is not None else ('П' if a.industrial_weight is not None else '?')
        print(f'{a.number:>3} | {a.fed_number[:23]:<23} | {a.tag[:9]:<9} | {a.category[:16]:<16} | {ch:^11} | {short_weight(a.slaughter_weight):>7}')
    print('Р = свободная реализация, П = промпереработка')


def parse_numbers(text: str, available: set[int]) -> set[int]:
    values = set()
    if not text.strip():
        return values
    for item in re.split(r'[,;\s]+', text.strip()):
        if not item:
            continue
        m = re.fullmatch(r'(\d+)\s*[-–]\s*(\d+)', item)
        if m:
            begin, end = map(int, m.groups())
            if begin > end:
                raise ValueError(f'Диапазон {item}: начало больше конца')
            if end-begin > 100000:
                raise ValueError('Слишком большой диапазон')
            values.update(range(begin, end+1))
        elif item.isdigit():
            values.add(int(item))
        else:
            raise ValueError(f'Некорректный номер: {item!r}. Пример: 1,3,7-10')
    unknown = values - available
    if unknown:
        raise ValueError(f'Таких № п/п нет в ведомости: {", ".join(map(str, sorted(unknown)))}')
    return values


def yes(prompt: str, default: bool = False) -> bool:
    while True:
        answer = input(prompt).strip().lower()
        if not answer:
            return default
        if answer in ('д','да','y','yes'):
            return True
        if answer in ('н','нет','n','no'):
            return False
        print('Ответь «да» или «нет».')


def record_to_json(ledger: Ledger, a: Animal, purchased: set[int] | None) -> dict[str, Any]:
    result = {
        'schema_version': 1,
        'document_date': ledger.date,
        'source_xlsx': ledger.path.name,
        'source_pdf': ledger.source_pdf,
        'source_sheet': ledger.sheet,
        'excel_row': a.excel_row,
        'row_number': a.number,
        'federal_number': a.fed_number,
        'tag': a.tag,
        'internal_number': a.internal_number,
        'birth_year': a.birth_year,
        'category_raw': a.category_raw,
        'category_normalized': a.category,
        'live_weight_kg': a.live_weight,
        'free_sale_weight_kg': a.free_weight,
        'industrial_weight_kg': a.industrial_weight,
        'slaughter_weight_kg': a.slaughter_weight,
        'operation': a.operation,
        'purchased': None if purchased is None else a.number in purchased,
        'purchase_label_source': 'unknown' if purchased is None else 'manually_confirmed',
    }
    return result


def write_jsonl(ledger: Ledger, target: Path, purchased: set[int] | None) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open('w', encoding='utf-8') as fp:
        for a in ledger.animals:
            fp.write(json.dumps(record_to_json(ledger,a,purchased), ensure_ascii=False) + '\n')
    print('JSONL сохранён:', target)



class MarkupError(ValueError):
    """Разметка не принадлежит этой ведомости или повреждена."""


def read_purchased_jsonl(ledger: Ledger, source: Path) -> set[int]:
    """Строгое восстановление разметки без угадывания по повторяющимся номерам.

    Проверяются ВСЕ строки и значения исходной ведомости. Частичный JSONL,
    неподтверждённая разметка, чужой документ и изменённые веса отклоняются.
    """
    source = Path(source)
    expected = {a.number: a for a in ledger.animals}
    found: dict[int, bool] = {}
    try:
        with source.open('r', encoding='utf-8-sig') as fp:
            for lineno, line in enumerate(fp, 1):
                if not line.strip():
                    continue
                try:
                    item = json.loads(line)
                except json.JSONDecodeError as exc:
                    raise MarkupError(f'JSONL, строка {lineno}: неверный JSON.') from exc
                if not isinstance(item, dict):
                    raise MarkupError(f'JSONL, строка {lineno}: ожидалась запись животного.')
                number = item.get('row_number')
                if type(number) is not int or number not in expected:
                    raise MarkupError(f'JSONL, строка {lineno}: № п/п {number!r} отсутствует в текущей ведомости.')
                if number in found:
                    raise MarkupError(f'JSONL, строка {lineno}: повтор № п/п {number}.')
                if item.get('document_date') != ledger.date:
                    raise MarkupError(f'JSONL, строка {lineno}: дата ведомости отличается.')
                if item.get('purchased') is None or type(item.get('purchased')) is not bool:
                    raise MarkupError(f'JSONL, строка {lineno}: разметка закупных не подтверждена.')
                animal = expected[number]
                fields = (
                    ('excel_row', animal.excel_row),
                    ('federal_number', animal.fed_number),
                    ('tag', animal.tag),
                    ('category_normalized', animal.category),
                    ('free_sale_weight_kg', animal.free_weight),
                    ('industrial_weight_kg', animal.industrial_weight),
                    ('slaughter_weight_kg', animal.slaughter_weight),
                )
                for name, actual in fields:
                    incoming = item.get(name)
                    if isinstance(actual, (float, int)) and not isinstance(actual, bool):
                        valid = (type(incoming) in (float, int)
                                 and math.isfinite(incoming)
                                 and math.isclose(float(incoming), float(actual), rel_tol=0, abs_tol=1e-6))
                    else:
                        valid = incoming == actual
                    if not valid:
                        raise MarkupError(f'JSONL, строка {lineno}, животное № {number}: поле {name} не совпадает с XLSX.')
                found[number] = item['purchased']
    except UnicodeDecodeError as exc:
        raise MarkupError('Разметка должна быть текстовым JSONL в UTF-8.') from exc
    if set(found) != set(expected):
        missing = sorted(set(expected) - set(found))
        raise MarkupError(f'Неполная разметка: записей {len(found)} вместо {len(expected)}. Отсутствуют № {missing[:10]}.')
    return {number for number, is_purchased in found.items() if is_purchased}


def compare_purchased(ledger: Ledger, current: set[int], incoming: set[int]) -> dict[str, Any]:
    """Разница между двумя разметками, а не между пересчитанными итогами."""
    valid = {a.number for a in ledger.animals}
    if not current <= valid or not incoming <= valid:
        raise ValueError('Разметка содержит отсутствующие номера животных.')
    get_weight = {a.number: a.slaughter_weight for a in ledger.animals}
    old_weight = sum(get_weight[n] for n in current)
    new_weight = sum(get_weight[n] for n in incoming)
    return {
        'added': sorted(incoming - current),
        'removed': sorted(current - incoming),
        'current_count': len(current), 'incoming_count': len(incoming),
        'current_weight': old_weight, 'incoming_weight': new_weight,
        'weight_delta': new_weight - old_weight,
    }

def pick_font() -> tuple[str, str]:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    candidates = [
        ('/usr/share/fonts/TTF/DejaVuSans.ttf','/usr/share/fonts/TTF/DejaVuSans-Bold.ttf'),
        ('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf','/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'),
        ('/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf','/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf'),
        ('C:/Windows/Fonts/arial.ttf','C:/Windows/Fonts/arialbd.ttf'),
    ]
    for regular, bold in candidates:
        if Path(regular).exists() and Path(bold).exists():
            pdfmetrics.registerFont(TTFont('RLRegular', regular))
            pdfmetrics.registerFont(TTFont('RLBold', bold))
            return 'RLRegular','RLBold'
    raise RuntimeError('Не найден шрифт с кириллицей. Установи DejaVu Sans (на CachyOS: sudo pacman -S ttf-dejavu).')


def export_pdf(ledger: Ledger, purchased: set[int], destination: Path) -> None:
    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_LEFT
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, LongTable, Table, TableStyle, KeepTogether
    except ImportError as exc:
        raise RuntimeError('Для экспорта PDF установи reportlab: python -m pip install reportlab') from exc
    regular, bold = pick_font()
    destination.parent.mkdir(parents=True, exist_ok=True)
    all_summary = summarize(ledger.animals)
    owned = [a for a in ledger.animals if a.number not in purchased]
    bought = [a for a in ledger.animals if a.number in purchased]
    doc = SimpleDocTemplate(
        str(destination), pagesize=landscape(A4), leftMargin=26, rightMargin=26,
        topMargin=26, bottomMargin=24, title=f'Ведомость {ledger.date}', author='Reindeer Ledger',
    )
    page_width = landscape(A4)[0] - 52
    styles = {
        'h1': ParagraphStyle('h1',fontName=bold,fontSize=15,leading=19,spaceAfter=5,textColor=colors.HexColor('#151515')),
        'h2': ParagraphStyle('h2',fontName=bold,fontSize=10,leading=14,spaceBefore=13,spaceAfter=5),
        'p': ParagraphStyle('p',fontName=regular,fontSize=9,leading=13,spaceAfter=4),
        'small': ParagraphStyle('small',fontName=regular,fontSize=7,leading=9),
        'cell': ParagraphStyle('cell',fontName=regular,fontSize=7.4,leading=10),
        'th': ParagraphStyle('th',fontName=bold,fontSize=7,leading=9),
    }
    def para(text: str, style='cell'):
        from xml.sax.saxutils import escape
        return Paragraph(escape(str(text)), styles[style])
    elements = [
        para(f'Ведомость оленей — {ledger.date}', 'h1'),
        para('Источник: ' + ledger.path.name + (f'; скан: {ledger.source_pdf}' if ledger.source_pdf else '') + '. Закупные отмечены вручную по № п/п.', 'p'),
    ]
    overview = [
        ['Показатель','Всего','Закупные','Прочие (не отмечены закупными)'],
        ['Голов',len(ledger.animals),len(bought),len(owned)],
        ['Свободная реализация, голов', all_summary['free_count'],sum(a.free_weight is not None for a in bought),sum(a.free_weight is not None for a in owned)],
        ['Свободная реализация, кг',short_weight(all_summary['free_weight']),short_weight(sum(a.free_weight or 0 for a in bought)),short_weight(sum(a.free_weight or 0 for a in owned))],
        ['Промпереработка, голов',all_summary['industrial_count'],sum(a.industrial_weight is not None for a in bought),sum(a.industrial_weight is not None for a in owned)],
        ['Промпереработка, кг',short_weight(all_summary['industrial_weight']),short_weight(sum(a.industrial_weight or 0 for a in bought)),short_weight(sum(a.industrial_weight or 0 for a in owned))],
        ['Убойный вес, кг',short_weight(all_summary['slaughter_weight']),short_weight(sum(a.slaughter_weight for a in bought)),short_weight(sum(a.slaughter_weight for a in owned))],
    ]
    summary_table = Table([[para(x, 'th' if ridx==0 else 'cell') for x in row] for ridx,row in enumerate(overview)], colWidths=[225,100,140,page_width-465], hAlign='LEFT')
    summary_table.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#F0F0F0')),
        ('LINEBELOW',(0,0),(-1,0),.6,colors.HexColor('#555555')),
        ('LINEBELOW',(0,1),(-1,-2),.2,colors.HexColor('#DFDFDF')),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),7),('RIGHTPADDING',(0,0),(-1,-1),6),
        ('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),
    ]))
    elements.append(summary_table)
    live_count = all_summary['live_known_count']
    if live_count != len(ledger.animals):
        elements.append(Spacer(1,7))
        elements.append(para(f'Живой вес: в XLSX заполнено {live_count} из {len(ledger.animals)} строк. Полный итог не рассчитывается.', 'p'))
    else:
        elements.append(para(f'Живой вес: {short_weight(all_summary["live_known_sum"])} кг.', 'p'))
    elements.append(para('Разбивка по половозрастным группам', 'h2'))
    cat_table = [['Категория','Всего, гол.','Закупные, гол.','Убойный вес, кг']]
    for category in sorted(all_summary['by_category']):
        subset = [a for a in ledger.animals if a.category == category]
        cat_table.append([category,len(subset),sum(a.number in purchased for a in subset),short_weight(sum(a.slaughter_weight for a in subset))])
    ct = Table([[para(x,'th' if i==0 else 'cell') for x in row] for i,row in enumerate(cat_table)],colWidths=[225,130,155,page_width-510],hAlign='LEFT')
    ct.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#F0F0F0')),('LINEBELOW',(0,0),(-1,0),.6,colors.HexColor('#555555')),('LINEBELOW',(0,1),(-1,-2),.2,colors.HexColor('#DFDFDF')),('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),('LEFTPADDING',(0,0),(-1,-1),7)]))
    elements.append(ct)
    elements.append(para('Реестр животных', 'h2'))
    detail = [['№','Федеральный №','Бирка','Внутр. №','Категория','Год','Реал., кг','Пром., кг','Закупной']]
    for a in ledger.animals:
        detail.append([a.number,a.fed_number or '—',a.tag or '—',a.internal_number or '—',a.category,a.birth_year or '—',short_weight(a.free_weight),short_weight(a.industrial_weight),'ДА' if a.number in purchased else 'нет'])
    widths = [30,145,65,65,116,40,72,72,page_width-605]
    dt = LongTable([[para(x,'th' if index==0 else 'cell') for x in row] for index,row in enumerate(detail)],colWidths=widths,repeatRows=1,hAlign='LEFT')
    dt.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#EDEDED')),
        ('LINEBELOW',(0,0),(-1,0),.6,colors.HexColor('#444444')),
        ('LINEBELOW',(0,1),(-1,-1),.15,colors.HexColor('#E3E3E3')),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),
        ('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),4),
    ]))
    elements.append(dt)
    def footer(canvas, document):
        canvas.saveState()
        canvas.setFont(regular,7)
        canvas.drawString(26,13, f'Сформировано Reindeer Ledger • {ledger.date} • исходный XLSX не изменён')
        canvas.drawRightString(landscape(A4)[0]-26,13,f'Стр. {document.page}')
        canvas.restoreState()
    doc.build(elements, onFirstPage=footer, onLaterPages=footer)
    print('PDF сохранён:',destination)


def main() -> int:
    parser = argparse.ArgumentParser(description='Подсчёт XLSX ведомостей, закупные олени, JSONL и PDF.')
    parser.add_argument('xlsx',type=Path,help='Исходная ведомость XLSX')
    parser.add_argument('--source-pdf',type=Path,help='Связанный скан PDF (для происхождения данных и ML)')
    parser.add_argument('--inspect',action='store_true',help='Только проверить/показать итоги; без вопросов и маркировки')
    parser.add_argument('--unlabeled-jsonl',type=Path,help='Экспортировать JSONL без догадок о закупных (purchased=null)')
    parser.add_argument('--output-dir',type=Path,default=None,help='Папка результата (по умолчанию ./reindeer_output)')
    args = parser.parse_args()
    if not args.xlsx.is_file():
        parser.error(f'Не найден файл: {args.xlsx}')
    try:
        ledger = read_ledger(args.xlsx)
    except (OSError, ValueError, KeyError, ET.ParseError) as exc:
        print('Ошибка при чтении XLSX:',exc,file=sys.stderr)
        return 2
    if args.source_pdf:
        if not args.source_pdf.is_file():
            parser.error(f'Не найден исходный PDF: {args.source_pdf}')
        ledger.source_pdf = args.source_pdf.name
    print(f'Ведомость: {ledger.date}; XLSX: {ledger.path.name}; лист: {ledger.sheet}')
    print_summary('Исходная ведомость',ledger.animals)
    if ledger.warnings:
        print('\nПроверка данных:')
        for msg in ledger.warnings:
            print('  !',msg)
    if args.unlabeled_jsonl:
        write_jsonl(ledger,args.unlabeled_jsonl,None)
    if args.inspect:
        return 0
    print('\nЗАПОЛНЕНИЕ ЗАКУПНЫХ. Вводи № п/п, а не федеральные номера.')
    print('Пример: 2,5,11-14. Повторяющиеся федеральные номера не являются уникальным ID.')
    if yes('Показать список всех животных? [д/Н]: ', default=False):
        print_animals(ledger.animals)
    available = {a.number for a in ledger.animals}
    while True:
        try:
            marked = parse_numbers(input('\nНомера ЗАКУПНЫХ (№ п/п): '), available)
        except ValueError as exc:
            print('Ошибка:',exc)
            continue
        if not marked:
            print('Ни одного закупного не выбрано. Это НЕ означает, что подтверждено отсутствие закупных.')
            if not yes('Подтвердить, что закупных действительно нет? [д/Н]: ', default=False):
                continue
        print('\nОтмеченные закупные:')
        for a in ledger.animals:
            if a.number in marked:
                print(f'  № {a.number}: фед. {a.fed_number or "—"}, {a.category}, {short_weight(a.slaughter_weight)} кг')
        print_summary('Закупные',[a for a in ledger.animals if a.number in marked])
        print_summary('Прочие (не закупные)',[a for a in ledger.animals if a.number not in marked])
        if yes('Подтверждаешь эту разметку? [д/Н]: ', default=False):
            break
    outdir = args.output_dir or Path('reindeer_output')
    base = f'{ledger.date.replace(".", "-")}_{ledger.path.stem}'
    write_jsonl(ledger,outdir/(base+'_marked.jsonl'),marked)
    if yes('Экспортировать сводку и реестр в PDF? [Д/н]: ',default=True):
        filename = outdir/(base+'_report.pdf')
        try:
            export_pdf(ledger,marked,filename)
        except (RuntimeError,OSError,ValueError) as exc:
            print('Не удалось создать PDF:',exc,file=sys.stderr)
            return 3
    print('Готово. Исходный XLSX остался без изменений.')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
