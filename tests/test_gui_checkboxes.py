"""Интеграционные проверки GUI; запускать там, где установлен PySide6.

Linux без дисплея: QT_QPA_PLATFORM=offscreen python -m pytest -q
"""
from pathlib import Path
import sys

import pytest

pytest.importorskip('PySide6')
from PySide6.QtCore import Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from ledger_core import Animal, Ledger, read_purchased_jsonl, write_jsonl
from main import AnimalModel, Window


def _fixture_ledger():
    # Синтетические данные; настоящие бухгалтерские XLSX в исходники не включаем.
    animals = [
        Animal(1, 'RU001', 'бирка 1', '', None, 20.0, None, 'бык', 'Бык', 2022, 5),
        Animal(2, 'RU002', 'бирка 2', '', None, None, 30.0, 'важ', 'Важенка', 2023, 6),
        Animal(3, 'RU003', '', '', None, 15.0, None, 'тб', 'Телёнок-самец', 2026, 7),
    ]
    return Ledger(Path('synthetic.xlsx'), '17.11.2026', '17.11.2026', animals, {}, [])


@pytest.fixture(scope='module')
def qtapp():
    app = QApplication.instance() or QApplication([])
    yield app


def test_model_toggle_and_persistence(qtapp, tmp_path):
    ledger = _fixture_ledger()
    model = AnimalModel()
    selected = set()
    model.reset(ledger, selected)
    idx = model.index(0, 0)
    first = ledger.animals[0].number
    assert model.setData(idx, Qt.CheckState.Checked, Qt.ItemDataRole.CheckStateRole)
    assert first in selected
    assert model.data(idx, Qt.ItemDataRole.CheckStateRole) == Qt.CheckState.Checked
    assert model.setData(idx, 0, Qt.ItemDataRole.CheckStateRole)
    assert first not in selected
    assert not model.setData(idx, 1, Qt.ItemDataRole.CheckStateRole)
    assert model.setData(idx, 2, Qt.ItemDataRole.CheckStateRole)
    file = tmp_path / 'marked.jsonl'
    write_jsonl(ledger, file, selected)
    assert read_purchased_jsonl(ledger, file) == {first}


def test_gui_single_click_toggle_through_proxy(qtapp):
    ledger = _fixture_ledger()
    window = Window()
    window.ledger = ledger
    window.purchased = set()
    window.animal_model.reset(ledger, window.purchased)
    window.show()
    window.navigate(2)
    qtapp.processEvents()
    proxy_idx = window.proxy.index(0, 0)
    source_idx = window.proxy.mapToSource(proxy_idx)
    first = ledger.animals[source_idx.row()].number
    rect = window.table.visualRect(proxy_idx)
    assert rect.isValid() and not rect.isEmpty()
    QTest.mouseClick(window.table.viewport(), Qt.MouseButton.LeftButton,
                     pos=rect.center())
    qtapp.processEvents()
    assert first in window.purchased
    assert window.animal_model.data(source_idx, Qt.ItemDataRole.CheckStateRole) == Qt.CheckState.Checked
    QTest.mouseClick(window.table.viewport(), Qt.MouseButton.LeftButton,
                     pos=rect.center())
    qtapp.processEvents()
    assert first not in window.purchased
    window.dirty = False
    window.close()
