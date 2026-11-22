"""Контракт настоящего AnimalModel.setData без графического окружения.

Достаём метод прямо из main.py через AST (а не копируем алгоритм),
подставляя только Qt-константы/сигналы, и проверяем изменение набора закупных.
Интеграционный тест с настоящим QTableView лежит в test_gui_checkboxes.py.
"""
import ast
from enum import Enum
from pathlib import Path
import sys
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from checkstate import parse_check_state


class CheckState(Enum):
    Unchecked = 0
    Checked = 2


class Roles:
    EditRole = 0
    CheckStateRole = 1
    UserRole = 2


def get_actual_set_data():
    code = (Path(__file__).resolve().parents[1] / 'main.py').read_text(encoding="utf-8")
    tree = ast.parse(code)
    model = next(node for node in tree.body if isinstance(node, ast.ClassDef)
                 and node.name == 'AnimalModel')
    method = next(node for node in model.body if isinstance(node, ast.FunctionDef)
                  and node.name == 'setData')
    namespace = {
        'Qt': SimpleNamespace(ItemDataRole=Roles),
        'parse_check_state': parse_check_state,
    }
    exec(compile(ast.fix_missing_locations(ast.Module(body=[method], type_ignores=[])),
                 'main.py', 'exec'), namespace)
    return namespace['setData']


class SignalSpy:
    def __init__(self):
        self.calls = []

    def emit(self, *args):
        self.calls.append(args)


class Index:
    def __init__(self, number=0, col=0):
        self._number = number
        self._col = col

    def isValid(self):
        return True

    def row(self):
        return self._number

    def column(self):
        return self._col


class FakeModel:
    def __init__(self):
        self.animals = [SimpleNamespace(number=8, slaughter_weight=45.0)]
        self.purchased = set()
        self.dataChanged = SignalSpy()
        self.purchase_changed = SignalSpy()


def test_real_setdata_checked_unchecked_and_signals():
    setter = get_actual_set_data()
    model = FakeModel()
    idx = Index()
    assert setter(model, idx, CheckState.Checked, Roles.CheckStateRole)
    assert model.purchased == {8}
    assert len(model.dataChanged.calls) == 1
    assert len(model.purchase_changed.calls) == 1
    assert setter(model, idx, CheckState.Checked, Roles.CheckStateRole)
    assert len(model.purchase_changed.calls) == 1  # повтор ничего не меняет
    assert setter(model, idx, 0, Roles.CheckStateRole)
    assert model.purchased == set()
    assert len(model.purchase_changed.calls) == 2
    assert setter(model, idx, 2, Roles.CheckStateRole)
    assert model.purchased == {8}
    assert setter(model, idx, CheckState.Unchecked, Roles.CheckStateRole)
    assert model.purchased == set()


def test_real_setdata_rejects_invalid_states():
    setter = get_actual_set_data()
    model = FakeModel()
    assert setter(model, Index(), 1, Roles.CheckStateRole) is False
    assert setter(model, Index(), True, Roles.CheckStateRole) is False
    assert setter(model, Index(col=1), 2, Roles.CheckStateRole) is False
    assert setter(model, Index(), 2, Roles.EditRole) is False
    assert model.purchased == set()
    assert not model.dataChanged.calls
