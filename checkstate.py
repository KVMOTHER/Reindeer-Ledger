"""Преобразование значения Qt.CheckStateRole без зависимости от версии PySide6.

PySide6 6.12 возвращает enum Qt.CheckState: int(enum) у него не поддерживается.
Qt также может передавать обычные целые 0/2 через прокси/делегат.
"""
from __future__ import annotations


def parse_check_state(value: object) -> bool | None:
    """True = Checked (2); False = Unchecked (0); None = недопустимое значение."""
    raw = getattr(value, 'value', value)
    # Не принимаем True/False как числа (bool наследует int),
    # а также PartiallyChecked=1: приложение не использует трёхсостояние.
    if type(raw) is not int:
        return None
    if raw == 2:
        return True
    if raw == 0:
        return False
    return None
