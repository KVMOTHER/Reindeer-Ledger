"""Регрессии на реальном файле опциональны (он не включается в исходники)."""
import os
import sys
from pathlib import Path
import pytest
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from ledger_core import read_ledger, summarize, parse_numbers


def test_17_nov_if_fixture_provided():
    path = os.environ.get('REINDEER_TEST_XLSX')
    if not path:
        pytest.skip('Укажите REINDEER_TEST_XLSX=/путь/к/17.11.26(11 Бр).xlsx')
    ledger = read_ledger(Path(path))
    summary = summarize(ledger.animals)
    assert (summary['count'], summary['slaughter_weight'],
            summary['free_count'], summary['free_weight'],
            summary['industrial_count'], summary['industrial_weight']) == (
                102, 2542, 62, 1700, 40, 842)
    marked = parse_numbers('8,11,12,14,21,25,28,30,32,37,43,47,50,54,55,60,61,62,64,65,73,75,89,98,100,102',
                           {a.number for a in ledger.animals})
    assert len(marked) == 26
    assert summarize([a for a in ledger.animals if a.number in marked])['slaughter_weight'] == 852
