import pytest
import pandas as pd
from pandas.testing import assert_frame_equal

from task_with_test.report import build_report

CONFIG = {
    "sheets": {"raw": "raw", "dict": "dict"},
    "columns": {
        "date": "date",
        "code": "code",
        "qty": "qty",
        "price": "price",
        "price_override": "price_override",
        "category": "category",
    },
    "output": {"sheet_name": "report", "revenue_column": "revenue"},
}


def sort_for_compare(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"])
    sort_cols = [c for c in ["category", "date"] if c in df.columns]
    return df.sort_values(sort_cols).reset_index(drop=True)


def test_build_report_basic():
    df_raw = pd.DataFrame({
        "date": ["01.01.2025", "01.01.2025", "02.01.2025"],
        "code": ["A", "B", "A"],
        "qty": [2, 1, 3],
        "price_override": [None, 12.0, None],
        "category": ["Food", "Drinks", "Food"],
        "noise": [1, 2, 3],
    })
    df_dict = pd.DataFrame({
        "code": ["A", "B"],
        "price": [10.0, 11.0],
        "category": ["Food", "Drinks"],
    })

    actual = build_report(df_raw, df_dict, CONFIG)

    expected = pd.DataFrame({
        "category": ["Drinks", "Food", "Food"],
        "date": pd.to_datetime(["01.01.2025", "01.01.2025", "02.01.2025"], dayfirst=True),
        "revenue": [12.0, 20.0, 30.0],
    })

    assert_frame_equal(sort_for_compare(actual), sort_for_compare(expected), check_dtype=False)


def test_build_report_handles_types_and_parsing():
    df_raw = pd.DataFrame({
        "date": ["31/12/2024", "01/01/2025"],
        "code": ["X", "X"],
        "qty": ["2", "3"],
        "price_override": ["", None],
        "category": ["Misc", "Misc"],
    })
    df_dict = pd.DataFrame({
        "code": ["X"],
        "price": ["5.5"],
        "category": ["Misc"],
    })

    actual = build_report(df_raw, df_dict, CONFIG)

    expected = pd.DataFrame({
        "category": ["Misc", "Misc"],
        "date": pd.to_datetime(["31/12/2024", "01/01/2025"], dayfirst=True),
        "revenue": [11.0, 16.5],
    })

    assert_frame_equal(sort_for_compare(actual), sort_for_compare(expected), check_dtype=False)


def test_dict_duplicates_raise_merge_error():
    df_raw = pd.DataFrame({
        "date": ["01.01.2025"],
        "code": ["A"],
        "qty": [1],
        "price_override": [None],
        "category": ["Food"],
    })
    df_dict = pd.DataFrame({
        "code": ["A", "A"], 
        "price": [10.0, 11.0],
        "category": ["Food", "Food"],
    })

    with pytest.raises(Exception):
        build_report(df_raw, df_dict, CONFIG)
