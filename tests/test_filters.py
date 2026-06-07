import pandas as pd
import pytest

from src.excel_utils_v2 import apply_filters


@pytest.fixture
def sample_df():
    return pd.DataFrame(
        {
            "name": ["Alice", "Bob", "Charlie", "Alice", None, "DAVID"],
            "age": [25, 30, 35, 25, 40, 28],
            "score": [85.5, 92.0, 78.0, 85.5, None, 91.0],
            "department": [
                "Sales",
                "Engineering",
                "Sales",
                "Marketing",
                "Engineering",
                "sales",
            ],
        }
    )


class TestEquals:
    def test_equals_exact_match(self, sample_df):
        filters = [{"column": "age", "operator": "equals", "value": 25}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 2
        assert all(result["age"] == 25)

    def test_equals_case_insensitive(self, sample_df):
        filters = [{"column": "name", "operator": "equals", "value": "alice"}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 2

    def test_equals_no_match(self, sample_df):
        filters = [{"column": "age", "operator": "equals", "value": 99}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 0


class TestNotEquals:
    def test_not_equals_exact_match(self, sample_df):
        filters = [{"column": "age", "operator": "not_equals", "value": 25}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 4
        assert all(result["age"] != 25)

    def test_not_equals_case_insensitive(self, sample_df):
        filters = [{"column": "name", "operator": "not_equals", "value": "alice"}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 4


class TestGreaterThan:
    def test_greater_than(self, sample_df):
        filters = [{"column": "age", "operator": "greater_than", "value": 30}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 2
        assert all(result["age"] > 30)

    def test_greater_than_no_match(self, sample_df):
        filters = [{"column": "age", "operator": "greater_than", "value": 50}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 0


class TestGreaterThanOrEqual:
    def test_greater_than_or_equal(self, sample_df):
        filters = [{"column": "age", "operator": "greater_than_or_equal", "value": 30}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 3
        assert all(result["age"] >= 30)

    def test_greater_than_or_equal_boundary(self, sample_df):
        filters = [{"column": "age", "operator": "greater_than_or_equal", "value": 25}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 6


class TestLessThan:
    def test_less_than(self, sample_df):
        filters = [{"column": "age", "operator": "less_than", "value": 30}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 3
        assert all(result["age"] < 30)

    def test_less_than_no_match(self, sample_df):
        filters = [{"column": "age", "operator": "less_than", "value": 20}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 0


class TestLessThanOrEqual:
    def test_less_than_or_equal(self, sample_df):
        filters = [{"column": "age", "operator": "less_than_or_equal", "value": 30}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 4
        assert all(result["age"] <= 30)

    def test_less_than_or_equal_boundary(self, sample_df):
        filters = [{"column": "age", "operator": "less_than_or_equal", "value": 25}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 2
        assert all(result["age"] <= 25)


class TestContains:
    def test_contains(self, sample_df):
        filters = [{"column": "name", "operator": "contains", "value": "li"}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 3

    def test_contains_case_insensitive(self, sample_df):
        filters = [{"column": "name", "operator": "contains", "value": "ALI"}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 2

    def test_contains_no_match(self, sample_df):
        filters = [{"column": "name", "operator": "contains", "value": "xyz"}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 0


class TestNotContains:
    def test_not_contains(self, sample_df):
        filters = [{"column": "name", "operator": "not_contains", "value": "li"}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 3

    def test_not_contains_case_insensitive(self, sample_df):
        filters = [{"column": "name", "operator": "not_contains", "value": "ALI"}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 4


class TestStartsWith:
    def test_starts_with(self, sample_df):
        filters = [{"column": "name", "operator": "starts_with", "value": "Ali"}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 2

    def test_starts_with_no_match(self, sample_df):
        filters = [{"column": "name", "operator": "starts_with", "value": "Z"}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 0

    def test_starts_with_case_insensitive_upper_value(self, sample_df):
        # "Sales", "Sales", and lowercase "sales" all match (case-insensitive)
        filters = [
            {"column": "department", "operator": "starts_with", "value": "Sales"}
        ]
        result = apply_filters(sample_df, filters)
        assert len(result) == 3

    def test_starts_with_case_insensitive_lower_value(self, sample_df):
        # A lowercase value still matches the capitalized "Sales" rows
        filters = [
            {"column": "department", "operator": "starts_with", "value": "sales"}
        ]
        result = apply_filters(sample_df, filters)
        assert len(result) == 3

    def test_ends_with_case_insensitive(self, sample_df):
        # Uppercase value matches the lowercase "...ring" departments
        filters = [{"column": "department", "operator": "ends_with", "value": "RING"}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 2

    def test_ends_with_no_match(self, sample_df):
        filters = [{"column": "name", "operator": "ends_with", "value": "xyz"}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 0


class TestIsNull:
    def test_is_null(self, sample_df):
        filters = [{"column": "name", "operator": "is_null"}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 1

    def test_is_null_no_nulls(self):
        df = pd.DataFrame({"a": [1, 2, 3], "b": ["x", "y", "z"]})
        filters = [{"column": "a", "operator": "is_null"}]
        result = apply_filters(df, filters)
        assert len(result) == 0


class TestIsNotNull:
    def test_is_not_null(self, sample_df):
        filters = [{"column": "name", "operator": "is_not_null"}]
        result = apply_filters(sample_df, filters)
        assert len(result) == 5

    def test_is_not_null_no_nulls(self):
        df = pd.DataFrame({"a": [1, 2, 3], "b": ["x", "y", "z"]})
        filters = [{"column": "a", "operator": "is_not_null"}]
        result = apply_filters(df, filters)
        assert len(result) == 3


class TestNaNEdgeCases:
    def test_greater_than_with_nan(self):
        df = pd.DataFrame({"values": [10, 20, None, 30]})
        filters = [{"column": "values", "operator": "greater_than", "value": 15}]
        result = apply_filters(df, filters)
        assert len(result) == 2

    def test_less_than_with_nan(self):
        df = pd.DataFrame({"values": [10, None, 5, 20]})
        filters = [{"column": "values", "operator": "less_than", "value": 10}]
        result = apply_filters(df, filters)
        assert len(result) == 1

    def test_equals_with_nan(self):
        df = pd.DataFrame({"values": ["10", None, "10", "20"]})
        filters = [{"column": "values", "operator": "equals", "value": "10"}]
        result = apply_filters(df, filters)
        assert len(result) == 2


class TestChainedFilters:
    def test_chained_filters_multiple_operators(self, sample_df):
        filters = [
            {"column": "department", "operator": "starts_with", "value": "sale"},
            {"column": "age", "operator": "greater_than", "value": 20},
        ]
        result = apply_filters(sample_df, filters)
        # All three Sales/sales rows now match (case-insensitive), all aged > 20
        assert len(result) == 3

    def test_chained_filters_three_filters(self, sample_df):
        filters = [
            {"column": "name", "operator": "contains", "value": "a"},
            {"column": "age", "operator": "greater_than_or_equal", "value": 25},
            {"column": "department", "operator": "not_equals", "value": "Marketing"},
        ]
        result = apply_filters(sample_df, filters)
        assert len(result) == 3


class TestEmptyFilters:
    def test_empty_filters_returns_original(self, sample_df):
        result = apply_filters(sample_df, [])
        assert len(result) == len(sample_df)

    def test_none_filters_returns_original(self, sample_df):
        result = apply_filters(sample_df, None)
        assert len(result) == len(sample_df)
