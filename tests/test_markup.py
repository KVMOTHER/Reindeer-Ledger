"""Изолированные регрессионные тесты: не требуют PySide6 или внешних XLSX."""
from pathlib import Path
import json
import sys

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from ledger_core import (Animal, Ledger, MarkupError, compare_purchased,
                         read_purchased_jsonl, write_jsonl)


def create_ledger(tmp_path):
    animals = [
        Animal(1, 'RU123', 'к12', '', None, 12.0, None, 'тб', 'Телёнок-самец', 2026, 5),
        Animal(2, 'RU123', 'к12', '', None, None, 18.0, 'тв', 'Телёнок-самка', 2026, 6),
        Animal(3, '643099013043460', '', '', None, 45.0, None, 'бык', 'Бык', 2021, 7),
    ]
    return Ledger(tmp_path / 'original.xlsx', '17.11.2026', '17.11.2026', animals, {}, [])


def test_roundtrip_duplicate_federal_number(tmp_path):
    ledger = create_ledger(tmp_path)
    file = tmp_path / 'selected.jsonl'
    write_jsonl(ledger, file, {2, 3})
    result = read_purchased_jsonl(ledger, file)
    assert result == {2, 3}
    assert compare_purchased(ledger, {1, 2}, result) == {
        'added': [3], 'removed': [1], 'current_count': 2,
        'incoming_count': 2, 'current_weight': 30.0,
        'incoming_weight': 63.0, 'weight_delta': 33.0,
    }


def test_reject_changed_weight(tmp_path):
    ledger = create_ledger(tmp_path)
    file = tmp_path / 'selected.jsonl'
    write_jsonl(ledger, file, {1})
    records = [json.loads(line) for line in file.read_text().splitlines()]
    records[1]['slaughter_weight_kg'] = 99.0
    file.write_text(''.join(json.dumps(record) + '\n' for record in records))
    with pytest.raises(MarkupError, match='slaughter_weight_kg'):
        read_purchased_jsonl(ledger, file)


def test_reject_wrong_document_and_partial(tmp_path):
    ledger = create_ledger(tmp_path)
    file = tmp_path / 'selected.jsonl'
    write_jsonl(ledger, file, {1})
    records = [json.loads(line) for line in file.read_text().splitlines()]
    records[0]['document_date'] = '18.11.2026'
    file.write_text(''.join(json.dumps(record) + '\n' for record in records))
    with pytest.raises(MarkupError, match='дата'):
        read_purchased_jsonl(ledger, file)
    file.write_text(json.dumps(records[2]) + '\n')
    with pytest.raises(MarkupError, match='Неполная'):
        read_purchased_jsonl(ledger, file)


def test_reject_unlabeled_and_repeated_id(tmp_path):
    ledger = create_ledger(tmp_path)
    file = tmp_path / 'selected.jsonl'
    write_jsonl(ledger, file, None)
    with pytest.raises(MarkupError, match='не подтверждена'):
        read_purchased_jsonl(ledger, file)
    write_jsonl(ledger, file, {2})
    file.write_text(file.read_text() + file.read_text().splitlines()[0] + '\n')
    with pytest.raises(MarkupError, match='повтор'):
        read_purchased_jsonl(ledger, file)
