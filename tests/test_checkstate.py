"""Pure Python: валидируем значения из разных версий PySide6 без GUI."""
from enum import Enum, IntEnum
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from checkstate import parse_check_state


class NewQtCheckState(Enum):
    Unchecked = 0
    PartiallyChecked = 1
    Checked = 2


class OldQtCheckState(IntEnum):
    Unchecked = 0
    PartiallyChecked = 1
    Checked = 2


def test_pyside6_enum_with_no_int_conversion():
    assert parse_check_state(NewQtCheckState.Checked) is True
    assert parse_check_state(NewQtCheckState.Unchecked) is False
    assert parse_check_state(NewQtCheckState.PartiallyChecked) is None


def test_oldqt_intenum_and_raw_int_values():
    assert parse_check_state(OldQtCheckState.Checked) is True
    assert parse_check_state(OldQtCheckState.Unchecked) is False
    assert parse_check_state(2) is True
    assert parse_check_state(0) is False


def test_reject_other_values():
    assert parse_check_state(1) is None
    assert parse_check_state(-1) is None
    assert parse_check_state(3) is None
    assert parse_check_state(True) is None
    assert parse_check_state(None) is None
    assert parse_check_state('2') is None
