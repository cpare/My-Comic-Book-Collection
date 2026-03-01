"""
test_unit.py — Unit tests for core.py logic.
Run with: python3 -m pytest test_unit.py -v
"""
import pytest
import pandas as pd
from unittest.mock import MagicMock, patch, call
from core import (
    similar,
    _format_grade,
    _classify_age,
    normalise_grade,
    safe_fillna,
    _pick_best_match,
    GetEbayPrice,
    SearchComicVine,
    CV_CONFIDENCE_THRESHOLD,
    GRADE_SCALE,
)


# =============================================================================
#   similar()
# =============================================================================
class TestSimilar:
    def test_identical(self):
        assert similar("ACTION COMICS #1", "ACTION COMICS #1") == 1.0

    def test_completely_different(self):
        assert similar("BATMAN #1", "SUPERMAN #500") < 0.5

    def test_partial_match(self):
        score = similar("ACTION COMICS #684", "ACTION COMICS #685")
        assert score > 0.8

    def test_case_insensitive_via_caller(self):
        # similar() is case-sensitive; callers upper() before passing
        assert similar("action comics #1".upper(), "ACTION COMICS #1") == 1.0


# =============================================================================
#   _format_grade()
# =============================================================================
class TestFormatGrade:
    def test_whole_number(self):
        assert _format_grade(9) == "9.0"

    def test_decimal(self):
        assert _format_grade(9.8) == "9.8"

    def test_low_grade(self):
        assert _format_grade(0.5) == "0.5"

    def test_perfect(self):
        assert _format_grade(10.0) == "10.0"


# =============================================================================
#   _classify_age()
# =============================================================================
class TestClassifyAge:
    def test_platinum(self):
        assert _classify_age("1930-01-01") == "Platinum Age"

    def test_golden(self):
        assert _classify_age("1945-06-01") == "Golden Age"

    def test_silver(self):
        assert _classify_age("1963-03-01") == "Silver Age"

    def test_bronze(self):
        assert _classify_age("1975-08-01") == "Bronze Age"

    def test_copper(self):
        assert _classify_age("1988-01-01") == "Copper Age"

    def test_modern(self):
        assert _classify_age("2005-05-01") == "Modern Age"

    def test_unknown_blank(self):
        assert _classify_age("") == "Unknown"

    def test_unknown_none(self):
        assert _classify_age(None) == "Unknown"

    def test_unknown_nan_string(self):
        assert _classify_age("nan") == "Unknown"

    def test_year_only(self):
        assert _classify_age("1968") == "Silver Age"

    def test_boundary_1956(self):
        assert _classify_age("1956-01-01") == "Silver Age"

    def test_boundary_1970(self):
        assert _classify_age("1970-01-01") == "Bronze Age"


# =============================================================================
#   normalise_grade()
# =============================================================================
class TestNormaliseGrade:
    def test_whole_number_string(self):
        assert normalise_grade("9") == "9.0"

    def test_decimal_string(self):
        assert normalise_grade("9.8") == "9.8"

    def test_blank(self):
        assert normalise_grade("") is None

    def test_nan_string(self):
        assert normalise_grade("nan") is None

    def test_none(self):
        assert normalise_grade(None) is None

    def test_text_grade(self):
        # Text grades like 'NM' are not numeric — return None
        assert normalise_grade("NM") is None

    def test_low_grade(self):
        assert normalise_grade("0.5") == "0.5"

    def test_perfect(self):
        assert normalise_grade("10") == "10.0"

    def test_already_normalised(self):
        assert normalise_grade("4.0") == "4.0"


# =============================================================================
#   safe_fillna()
# =============================================================================
class TestSafeFillna:
    def test_object_columns_filled_with_empty_string(self):
        df = pd.DataFrame({'title': ['Batman', None, 'Superman']})
        result = safe_fillna(df)
        assert result['title'].iloc[1] == ''

    def test_numeric_columns_filled_with_zero(self):
        df = pd.DataFrame({'value': [1.0, None, 3.0]})
        result = safe_fillna(df)
        assert result['value'].iloc[1] == 0.0

    def test_no_mutation_on_clean_df(self):
        df = pd.DataFrame({'title': ['Batman'], 'value': [9.8]})
        result = safe_fillna(df)
        assert result['title'].iloc[0] == 'Batman'
        assert result['value'].iloc[0] == 9.8

    def test_mixed_columns(self):
        df = pd.DataFrame({'title': [None], 'value': [None]})
        df['value'] = df['value'].astype(float)
        result = safe_fillna(df)
        assert result['title'].iloc[0] == ''
        assert result['value'].iloc[0] == 0.0


# =============================================================================
#   _pick_best_match()
# =============================================================================
class TestPickBestMatch:
    def _make_result(self, vol_name, issue_num, publisher_name=None):
        r = {'volume': {'name': vol_name}, 'issue_number': str(issue_num)}
        if publisher_name:
            r['publisher'] = {'name': publisher_name}
        return r

    def test_exact_match(self):
        results = [
            self._make_result('Action Comics', '684'),
            self._make_result('Detective Comics', '684'),
        ]
        best, score = _pick_best_match(results, 'ACTION COMICS #684', issue_number=684)
        assert best['volume']['name'] == 'Action Comics'
        assert score > 0.9

    def test_no_results(self):
        best, score = _pick_best_match([], 'ACTION COMICS #684')
        assert best is None
        assert score == 0.0

    def test_picks_higher_score(self):
        results = [
            self._make_result('Batman', '1'),
            self._make_result('Action Comics', '684'),
        ]
        best, score = _pick_best_match(results, 'ACTION COMICS #684', issue_number=684)
        assert best['volume']['name'] == 'Action Comics'

    def test_numeric_title_low_confidence(self):
        # "1963 #1" should NOT strongly match "True 3-D #1"
        results = [self._make_result('True 3-D', '1')]
        best, score = _pick_best_match(results, '1963 #1', issue_number=1)
        assert score < CV_CONFIDENCE_THRESHOLD

    def test_issue_number_mismatch_penalised(self):
        # "2099 Unlimited #2" should not match "2099 Unlimited #4"
        results = [self._make_result('2099 Unlimited', '4')]
        best, score = _pick_best_match(results, '2099 UNLIMITED #2', issue_number=2)
        assert score < CV_CONFIDENCE_THRESHOLD

    def test_publisher_bonus_breaks_tie(self):
        # Use a slightly-off title so base score < 1.0, leaving room for publisher bonus
        results_no_pub   = [self._make_result('Amazing Spiderman', '1')]
        results_with_pub = [self._make_result('Amazing Spiderman', '1')]
        results_with_pub[0]['publisher'] = {'name': 'Marvel'}

        _, score_no_pub   = _pick_best_match(
            results_no_pub,   'AMAZING SPIDER-MAN #1', issue_number=1, publisher='Marvel'
        )
        _, score_with_pub = _pick_best_match(
            results_with_pub, 'AMAZING SPIDER-MAN #1', issue_number=1, publisher='Marvel'
        )
        assert score_with_pub > score_no_pub

    def test_protected_fields_not_in_return_value(self):
        # _pick_best_match returns raw CV result — caller must not write
        # Title/Issue/Publisher/Volume back to sheet. This test verifies
        # the result dict does NOT contain a 'Title' key we'd accidentally use.
        results = [self._make_result('Action Comics', '684')]
        best, _ = _pick_best_match(results, 'ACTION COMICS #684', issue_number=684)
        assert 'Title' not in best
        assert 'Issue' not in best


# =============================================================================
#   GetEbayPrice() — blank/invalid grade handling
# =============================================================================
class TestGetEbayPriceGradeHandling:
    def _mock_session(self):
        return MagicMock()

    def test_blank_grade_returns_none(self):
        result = GetEbayPrice(self._mock_session(), 'BATMAN', 1, '', 'No')
        assert result is None

    def test_nan_grade_returns_none(self):
        result = GetEbayPrice(self._mock_session(), 'BATMAN', 1, 'nan', 'No')
        assert result is None

    def test_text_grade_returns_none(self):
        result = GetEbayPrice(self._mock_session(), 'BATMAN', 1, 'NM', 'No')
        assert result is None

    def test_valid_grade_calls_ebay(self):
        session = self._mock_session()
        with patch('core._ebay_sold_prices', return_value=[25.00]) as mock_ebay:
            result = GetEbayPrice(session, 'BATMAN', 1, '9.8', 'No')
            assert result == 25.00
            mock_ebay.assert_called_once()

    def test_no_sales_exact_tries_neighbours(self):
        session = self._mock_session()
        call_count = [0]

        def fake_prices(sess, query):
            call_count[0] += 1
            # Return price on the second call (first neighbour)
            return [30.00] if call_count[0] > 1 else []

        with patch('core._ebay_sold_prices', side_effect=fake_prices):
            result = GetEbayPrice(session, 'BATMAN', 1, '9.8', 'No')
            assert result is not None
            assert call_count[0] > 1

    def test_interpolation(self):
        session = self._mock_session()
        # Simulate: 9.6 = $20, 10.0 = $40, so 9.8 should interpolate to $30
        def fake_prices(sess, query):
            if '9.6' in query: return [20.00]
            if '10.0' in query: return [40.00]
            return []

        with patch('core._ebay_sold_prices', side_effect=fake_prices):
            result = GetEbayPrice(session, 'BATMAN', 1, '9.8', 'No')
            assert result == 30.00


# =============================================================================
#   SearchComicVine() — confidence threshold enforcement
# =============================================================================
class TestSearchComicVineConfidence:
    def _mock_session(self):
        return MagicMock()

    def test_low_confidence_returns_none(self):
        """A bad match (< threshold) should return None, not pollute the sheet."""
        with patch('core._cv_get') as mock_cv:
            # Return a result that won't match well
            mock_cv.return_value = [
                {'volume': {'name': 'True 3-D', 'id': 1},
                 'issue_number': '1',
                 'description': '', 'deck': '',
                 'image': None, 'cover_date': '1953-01-01',
                 'cover_price': '0.10', 'site_detail_url': '',
                 'character_credits': [], 'name': 'True 3-D #1'}
            ]
            result = SearchComicVine(self._mock_session(), 'FAKE_KEY', '1963', 1)
            assert result is None

    def test_high_confidence_returns_data(self):
        """A good match should return metadata."""
        mock_issue = {
            'volume': {'name': 'Action Comics', 'id': 12345},
            'issue_number': '684',
            'description': 'Death of Superman tie-in',
            'deck': 'Key issue — first appearance',
            'image': {'original_url': 'https://example.com/cover.jpg'},
            'cover_date': '1993-01-01',
            'cover_price': '1.25',
            'site_detail_url': 'https://comicvine.com/action-comics-684',
            'character_credits': [{'name': 'Superman'}],
            'name': 'Action Comics #684',
            'resource_type': 'issue',
        }
        with patch('core._cv_get') as mock_cv:
            mock_cv.return_value = [mock_issue]
            result = SearchComicVine(
                self._mock_session(), 'FAKE_KEY', 'ACTION COMICS', 684
            )
            assert result is not None
            assert result['cover_image'] == 'https://example.com/cover.jpg'
            assert result['key_issue'] == 'Yes'
            assert result['comic_age'] == 'Modern Age'
            assert result['confidence'] > CV_CONFIDENCE_THRESHOLD


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
