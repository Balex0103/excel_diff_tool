"""
tests/test_excel_diff.py - egyszerü unit tesztek az excel_diff logikahoz.
Futtatas: pytest tests/
"""

import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import pandas as pd
from excel_diff import compare_dataframes


def test_changed_value_detected():
    df_old = pd.DataFrame({"id": [1, 2], "osszeg": [1000, 2000]})
    df_new = pd.DataFrame({"id": [1, 2], "osszeg": [1000, 2500]})

    only_old, only_new, changed = compare_dataframes(df_old, df_new, ["id"])

    assert only_old.empty
    assert only_new.empty
    assert len(changed) == 1
    assert changed.iloc[0]["oszlop"] == "osszeg"
    assert changed.iloc[0]["regi_ertek"] == 2000
    assert changed.iloc[0]["uj_ertek"] == 2500


def test_only_in_old_and_new():
    df_old = pd.DataFrame({"id": [1, 2, 3], "nev": ["A", "B", "C"]})
    df_new = pd.DataFrame({"id": [1, 2, 4], "nev": ["A", "B", "D"]})

    only_old, only_new, changed = compare_dataframes(df_old, df_new, ["id"])

    assert len(only_old) == 1
    assert only_old.iloc[0]["id"] == 3
    assert len(only_new) == 1
    assert only_new.iloc[0]["id"] == 4
    assert changed.empty