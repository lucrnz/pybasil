"""Tests for the VBScript interpreter."""

import pytest
import io
from pybasil import (
    Interpreter,
    parse,
    VBScriptError,
)


class TestDateTimeCDate:
    """Test CDate conversion function."""

    def test_cdate_string_mdy(self):
        program = parse('d = CDate("6/15/2025")')
        interp = Interpreter()
        interp.interpret(program)
        from pybasil.runtime import VBScriptDate
        d = interp._environment.get('d')
        assert isinstance(d, VBScriptDate)
        assert d.year == 2025
        assert d.month == 6
        assert d.day == 15

    def test_cdate_string_ymd(self):
        from pybasil.runtime import VBScriptDate
        program = parse('d = CDate("2025-01-20")')
        interp = Interpreter()
        interp.interpret(program)
        d = interp._environment.get('d')
        assert isinstance(d, VBScriptDate)
        assert d.year == 2025
        assert d.month == 1
        assert d.day == 20

    def test_cdate_number(self):
        program = parse('''
        d = CDate(0)
        y = Year(d)
        m = Month(d)
        dy = Day(d)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('y') == 1899
        assert interp._environment.get('m') == 12
        assert interp._environment.get('dy') == 30

    def test_cdate_passthrough(self):
        program = parse('''
        d1 = CDate("3/1/2025")
        d2 = CDate(d1)
        result = (Year(d1) = Year(d2))
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is True

    def test_cdate_invalid_string_error(self):
        program = parse('d = CDate("not a date")')
        interp = Interpreter()
        with pytest.raises(VBScriptError):
            interp.interpret(program)

class TestDateTimeIsDate:
    """Test IsDate function."""

    def test_isdate_date_object(self):
        program = parse('result = IsDate(CDate("1/1/2025"))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is True

    def test_isdate_valid_string(self):
        program = parse('result = IsDate("12/25/2024")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is True

    def test_isdate_invalid_string(self):
        program = parse('result = IsDate("hello")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is False

    def test_isdate_number(self):
        program = parse('result = IsDate(42)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is False

    def test_isdate_empty(self):
        program = parse('result = IsDate(Empty)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is False

class TestDateTimeTypeName:
    """Test TypeName and VarType for dates."""

    def test_typename_date(self):
        program = parse('result = TypeName(CDate("1/1/2025"))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'Date'

    def test_vartype_date(self):
        program = parse('result = VarType(CDate("1/1/2025"))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 7  # vbDate

class TestDateTimeExtractors:
    """Test Year, Month, Day, Hour, Minute, Second, Weekday."""

    def test_year(self):
        program = parse('result = Year(CDate("6/15/2025"))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 2025

    def test_month(self):
        program = parse('result = Month(CDate("6/15/2025"))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 6

    def test_day(self):
        program = parse('result = Day(CDate("6/15/2025"))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 15

    def test_hour(self):
        program = parse('result = Hour(TimeSerial(14, 30, 45))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 14

    def test_minute(self):
        program = parse('result = Minute(TimeSerial(14, 30, 45))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 30

    def test_second(self):
        program = parse('result = Second(TimeSerial(14, 30, 45))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 45

    def test_weekday_sunday(self):
        """June 15, 2025 is a Sunday."""
        program = parse('result = Weekday(CDate("6/15/2025"))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 1  # Sunday

    def test_weekday_monday(self):
        """June 16, 2025 is a Monday."""
        program = parse('result = Weekday(CDate("6/16/2025"))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 2  # Monday

    def test_weekday_saturday(self):
        """June 14, 2025 is a Saturday."""
        program = parse('result = Weekday(CDate("6/14/2025"))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 7  # Saturday

    def test_extractors_from_string(self):
        """Extract components from a date string directly."""
        program = parse('result = Year(CDate("12/25/2024"))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 2024

class TestDateTimeSerial:
    """Test DateSerial and TimeSerial."""

    def test_dateserial_basic(self):
        program = parse('''
        d = DateSerial(2025, 3, 15)
        result = Year(d) & "/" & Month(d) & "/" & Day(d)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == '2025/3/15'

    def test_dateserial_leap_year(self):
        program = parse('''
        d = DateSerial(2024, 2, 29)
        result = Day(d)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 29

    def test_timeserial_basic(self):
        program = parse('''
        t = TimeSerial(9, 30, 0)
        result = Hour(t) & ":" & Minute(t) & ":" & Second(t)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == '9:30:0'

    def test_timeserial_midnight(self):
        program = parse('''
        t = TimeSerial(0, 0, 0)
        result = Hour(t)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 0

    def test_timeserial_end_of_day(self):
        program = parse('''
        t = TimeSerial(23, 59, 59)
        result = Hour(t) & ":" & Minute(t) & ":" & Second(t)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == '23:59:59'

class TestDateTimeValueFunctions:
    """Test DateValue and TimeValue."""

    def test_datevalue_strips_time(self):
        program = parse('''
        d = CDate("6/15/2025")
        dv = DateValue(d)
        result = (Year(dv) = 2025) And (Month(dv) = 6) And (Day(dv) = 15)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is True

    def test_timevalue_strips_date(self):
        program = parse('''
        t = TimeSerial(14, 30, 0)
        tv = TimeValue(t)
        result = Hour(tv)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 14

class TestDateTimeNames:
    """Test WeekdayName and MonthName."""

    def test_weekdayname_sunday(self):
        program = parse('result = WeekdayName(1)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'Sunday'

    def test_weekdayname_monday(self):
        program = parse('result = WeekdayName(2)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'Monday'

    def test_weekdayname_saturday(self):
        program = parse('result = WeekdayName(7)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'Saturday'

    def test_weekdayname_abbreviated(self):
        program = parse('result = WeekdayName(1, True)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'Sun'

    def test_monthname_january(self):
        program = parse('result = MonthName(1)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'January'

    def test_monthname_december(self):
        program = parse('result = MonthName(12)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'December'

    def test_monthname_abbreviated(self):
        program = parse('result = MonthName(6, True)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'Jun'

    def test_monthname_invalid(self):
        program = parse('result = MonthName(0)')
        interp = Interpreter()
        with pytest.raises(VBScriptError):
            interp.interpret(program)

    def test_monthname_13_invalid(self):
        program = parse('result = MonthName(13)')
        interp = Interpreter()
        with pytest.raises(VBScriptError):
            interp.interpret(program)

class TestDateTimeAdd:
    """Test DateAdd function."""

    def test_dateadd_days(self):
        program = parse('''
        d = DateAdd("d", 10, CDate("6/15/2025"))
        result = Day(d) & "/" & Month(d)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == '25/6'

    def test_dateadd_months(self):
        program = parse('''
        d = DateAdd("m", 3, CDate("6/15/2025"))
        result = Month(d)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 9

    def test_dateadd_years(self):
        program = parse('''
        d = DateAdd("yyyy", 1, CDate("6/15/2025"))
        result = Year(d)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 2026

    def test_dateadd_hours(self):
        program = parse('''
        d = DateAdd("h", 3, TimeSerial(10, 0, 0))
        result = Hour(d)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 13

    def test_dateadd_minutes(self):
        program = parse('''
        d = DateAdd("n", 45, TimeSerial(10, 30, 0))
        result = Hour(d) & ":" & Minute(d)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == '11:15'

    def test_dateadd_seconds(self):
        program = parse('''
        d = DateAdd("s", 90, TimeSerial(10, 0, 0))
        result = Minute(d) & ":" & Second(d)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == '1:30'

    def test_dateadd_negative_days(self):
        program = parse('''
        d = DateAdd("d", -5, CDate("6/15/2025"))
        result = Day(d)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 10

    def test_dateadd_weeks(self):
        program = parse('''
        d = DateAdd("ww", 2, CDate("6/1/2025"))
        result = Day(d)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 15

    def test_dateadd_quarter(self):
        program = parse('''
        d = DateAdd("q", 1, CDate("1/15/2025"))
        result = Month(d)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 4

class TestDateTimeDiff:
    """Test DateDiff function."""

    def test_datediff_days(self):
        program = parse('''
        result = DateDiff("d", CDate("6/1/2025"), CDate("6/15/2025"))
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 14

    def test_datediff_months(self):
        program = parse('''
        result = DateDiff("m", CDate("1/1/2025"), CDate("6/1/2025"))
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 5

    def test_datediff_years(self):
        program = parse('''
        result = DateDiff("yyyy", CDate("1/1/2020"), CDate("1/1/2025"))
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 5

    def test_datediff_hours(self):
        program = parse('''
        result = DateDiff("h", TimeSerial(10, 0, 0), TimeSerial(14, 0, 0))
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 4

    def test_datediff_minutes(self):
        program = parse('''
        result = DateDiff("n", TimeSerial(10, 0, 0), TimeSerial(10, 45, 0))
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 45

    def test_datediff_seconds(self):
        program = parse('''
        result = DateDiff("s", TimeSerial(10, 0, 0), TimeSerial(10, 0, 30))
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 30

    def test_datediff_negative(self):
        program = parse('''
        result = DateDiff("d", CDate("6/15/2025"), CDate("6/1/2025"))
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == -14

    def test_datediff_quarters(self):
        program = parse('''
        result = DateDiff("q", CDate("1/1/2025"), CDate("10/1/2025"))
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 3

class TestDateTimePart:
    """Test DatePart function."""

    def test_datepart_year(self):
        program = parse('result = DatePart("yyyy", CDate("6/15/2025"))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 2025

    def test_datepart_quarter(self):
        program = parse('result = DatePart("q", CDate("6/15/2025"))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 2

    def test_datepart_month(self):
        program = parse('result = DatePart("m", CDate("6/15/2025"))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 6

    def test_datepart_day(self):
        program = parse('result = DatePart("d", CDate("6/15/2025"))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 15

    def test_datepart_day_of_year(self):
        program = parse('result = DatePart("y", CDate("2/1/2025"))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 32  # Jan has 31 days + 1

    def test_datepart_hour(self):
        program = parse('result = DatePart("h", TimeSerial(14, 30, 0))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 14

    def test_datepart_minute(self):
        program = parse('result = DatePart("n", TimeSerial(14, 30, 0))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 30

    def test_datepart_second(self):
        program = parse('result = DatePart("s", TimeSerial(14, 30, 45))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 45

class TestDateTimeFormat:
    """Test FormatDateTime function."""

    def test_formatdatetime_short_date(self):
        program = parse('result = FormatDateTime(CDate("6/15/2025"), 2)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == '6/15/2025'

    def test_formatdatetime_long_date(self):
        program = parse('result = FormatDateTime(CDate("6/15/2025"), 1)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'Sunday, June 15, 2025'

    def test_formatdatetime_long_time(self):
        program = parse('result = FormatDateTime(TimeSerial(14, 30, 0), 3)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == '2:30:00 PM'

    def test_formatdatetime_short_time(self):
        program = parse('result = FormatDateTime(TimeSerial(14, 30, 0), 4)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == '14:30'

    def test_formatdatetime_am(self):
        program = parse('result = FormatDateTime(TimeSerial(9, 5, 0), 3)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == '9:05:00 AM'

class TestDateTimeNowAndTimer:
    """Test Now, Date, Time, and Timer (non-deterministic, so we test types/ranges)."""

    def test_now_returns_date(self):
        program = parse('result = TypeName(Now())')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'Date'

    def test_date_returns_date(self):
        program = parse('result = TypeName(Date())')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'Date'

    def test_time_returns_date(self):
        program = parse('result = TypeName(Time())')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'Date'

    def test_now_has_valid_year(self):
        program = parse('result = Year(Now())')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') >= 2025

    def test_timer_returns_number(self):
        program = parse('result = (Timer() >= 0)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is True

    def test_timer_less_than_day(self):
        program = parse('result = (Timer() < 86400)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is True

class TestDateTimeCoercion:
    """Test date coercion to number, string, boolean."""

    def test_date_to_string(self):
        output = io.StringIO()
        program = parse('''
        d = DateSerial(2025, 1, 1)
        WScript.Echo d
        ''')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        assert output.getvalue().strip() == '1/1/2025'

    def test_date_to_number(self):
        program = parse('''
        d = DateSerial(2025, 1, 1)
        result = CDbl(d)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        # Jan 1, 2025 is 45658 days from Dec 30, 1899
        assert interp._environment.get('result') == 45658.0

    def test_date_to_boolean_nonzero(self):
        program = parse('''
        d = DateSerial(2025, 1, 1)
        result = CBool(d)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is True

    def test_date_concat(self):
        program = parse('''
        d = DateSerial(2025, 6, 15)
        result = "Date is: " & d
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'Date is: 6/15/2025'

    def test_time_to_string(self):
        output = io.StringIO()
        program = parse('''
        t = TimeSerial(14, 30, 0)
        WScript.Echo t
        ''')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        assert output.getvalue().strip() == '2:30:00 PM'

    def test_datetime_to_string(self):
        program = parse('''
        d = CDate("6/15/2025")
        result = CStr(d)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == '6/15/2025'

class TestDateTimeIntegration:
    """Integration tests combining multiple date/time functions."""

    def test_date_arithmetic_loop(self):
        output = io.StringIO()
        program = parse('''
        Dim d, i
        d = DateSerial(2025, 1, 1)
        For i = 1 To 3
            d = DateAdd("m", 1, d)
        Next
        WScript.Echo Month(d)
        ''')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        assert output.getvalue().strip() == '4'

    def test_weekday_and_name_roundtrip(self):
        program = parse('''
        d = DateSerial(2025, 6, 15)
        wd = Weekday(d)
        result = WeekdayName(wd)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'Sunday'

    def test_month_and_name_roundtrip(self):
        program = parse('''
        d = DateSerial(2025, 6, 15)
        m = Month(d)
        result = MonthName(m)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'June'

    def test_datepart_all_components(self):
        program = parse('''
        d = CDate("6/15/2025")
        result = DatePart("yyyy", d) & "-" & DatePart("m", d) & "-" & DatePart("d", d)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == '2025-6-15'

    def test_datediff_and_dateadd_inverse(self):
        """DateAdd(DateDiff(d1, d2)) should return d2."""
        program = parse('''
        d1 = DateSerial(2025, 1, 1)
        d2 = DateSerial(2025, 6, 15)
        diff = DateDiff("d", d1, d2)
        d3 = DateAdd("d", diff, d1)
        result = (Year(d3) = 2025) And (Month(d3) = 6) And (Day(d3) = 15)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is True

    def test_all_month_names(self):
        output = io.StringIO()
        program = parse('''
        Dim i, names
        names = ""
        For i = 1 To 12
            If i > 1 Then names = names & ","
            names = names & MonthName(i, True)
        Next
        WScript.Echo names
        ''')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        assert output.getvalue().strip() == 'Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec'

    def test_all_weekday_names(self):
        output = io.StringIO()
        program = parse('''
        Dim i, names
        names = ""
        For i = 1 To 7
            If i > 1 Then names = names & ","
            names = names & WeekdayName(i, True)
        Next
        WScript.Echo names
        ''')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        assert output.getvalue().strip() == 'Sun,Mon,Tue,Wed,Thu,Fri,Sat'
