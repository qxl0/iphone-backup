import sys
sys.path.insert(0, ".")
from iphone_downloader import format_elapsed, format_bar, format_eta


class TestFormatElapsed:
    def test_less_than_a_minute(self):
        assert format_elapsed(45) == "less than a minute"

    def test_one_minute(self):
        assert format_elapsed(60) == "1 minute"

    def test_plural_minutes(self):
        assert format_elapsed(180) == "3 minutes"

    def test_one_hour_zero_minutes(self):
        assert format_elapsed(3600) == "1 hour 0 minutes"

    def test_one_hour_one_minute(self):
        assert format_elapsed(3660) == "1 hour 1 minute"

    def test_hours_and_minutes(self):
        assert format_elapsed(6120) == "1 hour 42 minutes"

    def test_two_hours(self):
        assert format_elapsed(7200) == "2 hours 0 minutes"


class TestFormatBar:
    def test_empty(self):
        assert format_bar(0, 100, width=10) == "[░░░░░░░░░░]"

    def test_full(self):
        assert format_bar(100, 100, width=10) == "[██████████]"

    def test_half(self):
        assert format_bar(50, 100, width=10) == "[█████░░░░░]"

    def test_zero_total(self):
        assert format_bar(0, 0, width=4) == "[░░░░]"


class TestFormatEta:
    def test_no_progress_yet(self):
        result = format_eta(0, 100, elapsed=10.0)
        assert result == "estimating..."

    def test_zero_elapsed(self):
        result = format_eta(10, 100, elapsed=0.0)
        assert result == "estimating..."

    def test_less_than_one_minute(self):
        # 50 done in 10s → rate 5/s → 50 remaining → 10s left
        result = format_eta(50, 100, elapsed=10.0)
        assert result == "< 1 min left"

    def test_minutes(self):
        # 10 done in 60s → rate 1/6 per s → 90 remaining → 540s = 9 min
        result = format_eta(10, 100, elapsed=60.0)
        assert result == "~9 min left"

    def test_hours(self):
        # 10 done in 600s → rate 1/60 per s → 90 remaining → 5400s = 90 min = 1h 30m
        result = format_eta(10, 100, elapsed=600.0)
        assert result == "~1h 30m left"
