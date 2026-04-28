"""VBScript built-in functions.

Each function is a plain callable that receives already-evaluated arguments
and uses the coercion helpers (_to_number, _to_string, _to_boolean) provided
by the interpreter at registration time via a thin wrapper.

The module is organised by category:
  - String functions   (Len, Left, Right, Mid, Trim, LTrim, RTrim, UCase,
                        LCase, InStr, Replace, Split, Join)
  - Conversion functions (CStr, CInt, CLng, CDbl, CBool, CDate)
  - Type-checking functions (IsNumeric, IsArray, IsDate, IsEmpty, IsNull,
                             IsObject, TypeName, VarType)
  - Numeric functions  (Abs, Sqr, Int, Fix, Round, Rnd, Randomize)
  - Object functions   (CreateObject, GetObject)
  - Array functions    (UBound, LBound, Array)
  - Dialog stubs       (MsgBox, InputBox)
"""

from __future__ import annotations
import math
from typing import Any, TYPE_CHECKING

from lark.exceptions import UnexpectedInput

from ..runtime import (
    VBScriptError,
    VBScriptObject,
    VBScriptNothing,
    VBScriptEmpty,
    VBScriptNull,
    VBScriptDate,
    VBScriptArray,
    VBScriptDictionary,
    VBScriptFileSystemObject,
    VBScriptClassInstance,
    WScriptObject,
)

if TYPE_CHECKING:
    from ..interpreter import Interpreter


# ---------------------------------------------------------------------------
#  Dialog stubs
# ---------------------------------------------------------------------------

def builtin_msgbox(interp: Interpreter, *args: Any) -> int:
    """MsgBox function (simplified)."""
    if args:
        print(interp._to_string(args[0]))
    return 1  # vbOK


def builtin_inputbox(
    interp: Interpreter, prompt: str, title: str = '', default: str = '',
) -> str:
    """InputBox function (simplified)."""
    return default


# ---------------------------------------------------------------------------
#  String functions
# ---------------------------------------------------------------------------

def builtin_len(interp: Interpreter, value: Any) -> int:
    """Len function."""
    if isinstance(value, str):
        return len(value)
    if isinstance(value, VBScriptNull):
        from ..runtime import NULL
        return NULL
    return len(interp._to_string(value))


def builtin_left(interp: Interpreter, string: str, length: int) -> str:
    """Left function."""
    n = int(length)
    if n < 0:
        raise VBScriptError('Invalid procedure call or argument')
    return interp._to_string(string)[:n]


def builtin_right(interp: Interpreter, string: str, length: int) -> str:
    """Right function."""
    s = interp._to_string(string)
    n = int(length)
    if n < 0:
        raise VBScriptError('Invalid procedure call or argument')
    return s[-n:] if n > 0 else ''


def builtin_mid(
    interp: Interpreter, string: str, start: int, length: int = None,
) -> str:
    """Mid function."""
    s = interp._to_string(string)
    start_val = int(start)
    if start_val < 1:
        raise VBScriptError('Invalid procedure call or argument')
    start_idx = start_val - 1  # VBScript is 1-indexed
    if length is None:
        return s[start_idx:]
    return s[start_idx : start_idx + int(length)]


def builtin_trim(interp: Interpreter, string: str) -> str:
    """Trim function."""
    return interp._to_string(string).strip()


def builtin_ltrim(interp: Interpreter, string: str) -> str:
    """LTrim function."""
    return interp._to_string(string).lstrip()


def builtin_rtrim(interp: Interpreter, string: str) -> str:
    """RTrim function."""
    return interp._to_string(string).rstrip()


def builtin_ucase(interp: Interpreter, string: str) -> str:
    """UCase function."""
    return interp._to_string(string).upper()


def builtin_lcase(interp: Interpreter, string: str) -> str:
    """LCase function."""
    return interp._to_string(string).lower()


def builtin_instr(interp: Interpreter, *args: Any) -> int:
    """InStr function."""
    if len(args) == 2:
        start = 1
        string1, string2 = args
        compare = 0
    elif len(args) == 3:
        start = int(args[0])
        string1, string2 = args[1], args[2]
        compare = 0
    elif len(args) >= 4:
        start = int(args[0])
        string1, string2 = args[1], args[2]
        compare = int(args[3])
    else:
        return 0

    s1 = string1 if isinstance(string1, str) else interp._to_string(string1)
    s2 = string2 if isinstance(string2, str) else interp._to_string(string2)
    start_idx = start - 1  # VBScript is 1-indexed
    if compare == 1:
        idx = s1.lower().find(s2.lower(), start_idx)
    else:
        idx = s1.find(s2, start_idx)
    return idx + 1 if idx >= 0 else 0


def builtin_replace(
    interp: Interpreter,
    string: str,
    find: str,
    replace_with: str,
    start: int = 1,
    count: int = -1,
    compare: int = 0,
) -> str:
    """Replace function."""
    s = interp._to_string(string)
    f = interp._to_string(find)
    r = interp._to_string(replace_with)
    start_val = int(start)
    count_val = int(count)
    compare_val = int(compare)

    # VBScript Replace returns the substring starting at 'start'
    s = s[start_val - 1:]

    if compare_val == 1:
        # Case-insensitive replace
        result = []
        lower_s = s.lower()
        lower_f = f.lower()
        i = 0
        replacements = 0
        while i < len(s):
            if lower_s[i:i + len(lower_f)] == lower_f and (count_val < 0 or replacements < count_val):
                result.append(r)
                i += len(f)
                replacements += 1
            else:
                result.append(s[i])
                i += 1
        return ''.join(result)
    else:
        if count_val >= 0:
            return s.replace(f, r, count_val)
        else:
            return s.replace(f, r)


def builtin_split(
    interp: Interpreter,
    string: str,
    delimiter: str = ' ',
    count: int = -1,
    compare: int = 0,
) -> VBScriptArray:
    """Split function."""
    s = interp._to_string(string)
    d = interp._to_string(delimiter)
    if count > 0:
        parts = s.split(d, count - 1)
    else:
        parts = s.split(d)
    arr = VBScriptArray([len(parts) - 1], is_dynamic=True)
    for i, part in enumerate(parts):
        arr.set_element([i], part)
    return arr


def builtin_join(interp: Interpreter, array: list, delimiter: str = ' ') -> str:
    """Join function."""
    d = interp._to_string(delimiter)
    return d.join(interp._to_string(item) for item in array)


def builtin_instrrev(interp: Interpreter, *args: Any) -> int:
    """InStrRev function - search from the right."""
    if len(args) == 2:
        string1, string2 = args
        start = -1
        compare = 0
    elif len(args) == 3:
        string1, string2 = args[0], args[1]
        start = int(args[2])
        compare = 0
    elif len(args) >= 4:
        string1, string2 = args[0], args[1]
        start = int(args[2])
        compare = int(args[3])
    else:
        return 0

    s1 = string1 if isinstance(string1, str) else interp._to_string(string1)
    s2 = string2 if isinstance(string2, str) else interp._to_string(string2)

    if start == -1:
        start = len(s1)
    if start < 1:
        raise VBScriptError('Invalid procedure call or argument')

    search_in = s1[:start]
    if compare == 1:
        idx = search_in.lower().rfind(s2.lower())
    else:
        idx = search_in.rfind(s2)
    return idx + 1 if idx >= 0 else 0


def builtin_strcomp(interp: Interpreter, string1: Any, string2: Any, compare: int = 0) -> int:
    """StrComp function - compare two strings."""
    s1 = interp._to_string(string1)
    s2 = interp._to_string(string2)
    compare_val = int(compare)
    if compare_val == 1:
        s1 = s1.lower()
        s2 = s2.lower()
    if s1 < s2:
        return -1
    elif s1 > s2:
        return 1
    return 0


def builtin_string(interp: Interpreter, number: Any, character: Any) -> str:
    """String function - repeat a character n times."""
    n = int(interp._to_number(number))
    if n < 0:
        raise VBScriptError('Invalid procedure call or argument')
    if isinstance(character, (int, float)) and not isinstance(character, bool):
        ch = chr(int(character))
    else:
        s = interp._to_string(character)
        ch = s[0] if s else ''
    return ch * n


def builtin_space(interp: Interpreter, number: Any) -> str:
    """Space function - return n spaces."""
    n = int(interp._to_number(number))
    if n < 0:
        raise VBScriptError('Invalid procedure call or argument')
    return ' ' * n


def builtin_strreverse(interp: Interpreter, string: Any) -> str:
    """StrReverse function - reverse a string."""
    return interp._to_string(string)[::-1]


def builtin_asc(interp: Interpreter, string: Any) -> int:
    """Asc function - return ASCII code of first character."""
    s = interp._to_string(string)
    if not s:
        raise VBScriptError('Invalid procedure call or argument')
    return ord(s[0])


def builtin_ascw(interp: Interpreter, string: Any) -> int:
    """AscW function - return Unicode code of first character."""
    s = interp._to_string(string)
    if not s:
        raise VBScriptError('Invalid procedure call or argument')
    return ord(s[0])


def builtin_chr(interp: Interpreter, charcode: Any) -> str:
    """Chr function - return character from ASCII code."""
    code = int(interp._to_number(charcode))
    if code < 0 or code > 255:
        raise VBScriptError('Invalid procedure call or argument')
    return chr(code)


def builtin_chrw(interp: Interpreter, charcode: Any) -> str:
    """ChrW function - return character from Unicode code."""
    code = int(interp._to_number(charcode))
    return chr(code)


def builtin_hex(interp: Interpreter, number: Any) -> str:
    """Hex function - convert number to hex string."""
    n = int(interp._to_number(number))
    if n < 0:
        if -32767 <= n <= 32767:
            n = n & 0xFFFF
        else:
            n = n & 0xFFFFFFFF
    return format(n, 'X')


def builtin_oct(interp: Interpreter, number: Any) -> str:
    """Oct function - convert number to octal string."""
    n = int(interp._to_number(number))
    if n < 0:
        n = n & 0xFFFFFFFF
    return format(n, 'o')


# ---------------------------------------------------------------------------
#  Conversion functions
# ---------------------------------------------------------------------------

def builtin_cstr(interp: Interpreter, value: Any) -> str:
    """CStr function."""
    return interp._to_string(value)


def builtin_cint(interp: Interpreter, value: Any) -> int:
    """CInt function."""
    if isinstance(value, int) and not isinstance(value, bool):
        return value
    return round(interp._to_number(value))


def builtin_clng(interp: Interpreter, value: Any) -> int:
    """CLng function."""
    if isinstance(value, int) and not isinstance(value, bool):
        return value
    return round(interp._to_number(value))


def builtin_cdbl(interp: Interpreter, value: Any) -> float:
    """CDbl function."""
    if isinstance(value, float):
        return value
    if isinstance(value, int) and not isinstance(value, bool):
        return float(value)
    return interp._to_number(value)


def builtin_cbool(interp: Interpreter, value: Any) -> bool:
    """CBool function."""
    return interp._to_boolean(value)


def builtin_cdate(interp: Interpreter, value: Any) -> VBScriptDate:
    """CDate function - convert a value to a Date."""
    if isinstance(value, VBScriptDate):
        return value
    if isinstance(value, str):
        return VBScriptDate.from_string(value)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return VBScriptDate(float(value))
    raise VBScriptError('Type mismatch')


# ---------------------------------------------------------------------------
#  Type-checking functions
# ---------------------------------------------------------------------------

def builtin_isnumeric(interp: Interpreter, value: Any) -> bool:
    """IsNumeric function."""
    if isinstance(value, (int, float)):
        return True
    if isinstance(value, str):
        try:
            float(value)
            return True
        except ValueError:
            return False
    return False


def builtin_isarray(interp: Interpreter, value: Any) -> bool:
    """IsArray function."""
    return isinstance(value, (VBScriptArray, list, tuple))


def builtin_isdate(interp: Interpreter, value: Any) -> bool:
    """IsDate function - returns True if value is or can be converted to a Date."""
    if isinstance(value, VBScriptDate):
        return True
    if isinstance(value, str):
        try:
            VBScriptDate.from_string(value)
            return True
        except VBScriptError:
            return False
    return False


def builtin_isempty(interp: Interpreter, value: Any) -> bool:
    """IsEmpty function."""
    return isinstance(value, VBScriptEmpty)


def builtin_isnull(interp: Interpreter, value: Any) -> bool:
    """IsNull function."""
    return isinstance(value, VBScriptNull)


def builtin_isobject(interp: Interpreter, value: Any) -> bool:
    """IsObject function."""
    return isinstance(value, (WScriptObject, VBScriptObject, VBScriptNothing, VBScriptDictionary))


def builtin_typename(interp: Interpreter, value: Any) -> str:
    """TypeName function."""
    if isinstance(value, VBScriptEmpty):
        return 'Empty'
    if isinstance(value, VBScriptNull):
        return 'Null'
    if isinstance(value, VBScriptNothing):
        return 'Nothing'
    if isinstance(value, VBScriptDate):
        return 'Date'
    if isinstance(value, VBScriptClassInstance):
        return value.class_name
    if isinstance(value, VBScriptArray):
        return 'Variant()'
    if isinstance(value, bool):
        return 'Boolean'
    if isinstance(value, int):
        if -32767 <= value <= 32767:
            return 'Integer'
        return 'Long'
    if isinstance(value, float):
        return 'Double'
    if isinstance(value, str):
        return 'String'
    if isinstance(value, (list, tuple)):
        return 'Variant()'
    return 'Object'


def builtin_vartype(interp: Interpreter, value: Any) -> int:
    """VarType function."""
    if isinstance(value, VBScriptEmpty):
        return 0  # vbEmpty
    if isinstance(value, VBScriptNull):
        return 1  # vbNull
    if isinstance(value, VBScriptDate):
        return 7  # vbDate
    if isinstance(value, bool):
        return 11  # vbBoolean
    if isinstance(value, int):
        return 2  # vbInteger
    if isinstance(value, float):
        return 5  # vbDouble
    if isinstance(value, str):
        return 8  # vbString
    if isinstance(value, (VBScriptArray, list, tuple)):
        return 8204  # vbArray + vbVariant
    if isinstance(value, (VBScriptNothing, VBScriptDictionary, VBScriptClassInstance, VBScriptObject, WScriptObject)):
        return 9  # vbObject
    return 12  # vbVariant


# ---------------------------------------------------------------------------
#  Numeric functions
# ---------------------------------------------------------------------------

def builtin_abs(interp: Interpreter, value: Any) -> float:
    """Abs function."""
    return abs(interp._to_number(value))


def builtin_sqr(interp: Interpreter, value: Any) -> float:
    """Sqr function."""
    return math.sqrt(interp._to_number(value))


def builtin_int(interp: Interpreter, value: Any) -> int:
    """Int function."""
    return int(math.floor(interp._to_number(value)))


def builtin_fix(interp: Interpreter, value: Any) -> int:
    """Fix function."""
    return int(interp._to_number(value))


def builtin_round(interp: Interpreter, value: Any, decimals: int = 0) -> float:
    """Round function using Decimal to match VBScript banker's rounding."""
    from decimal import Decimal, ROUND_HALF_EVEN
    n = interp._to_number(value)
    d = int(decimals)
    result = Decimal(str(n)).quantize(Decimal(10) ** -d, rounding=ROUND_HALF_EVEN)
    return int(result) if d == 0 else float(result)


def builtin_rnd(interp: Interpreter, number: float = 1) -> float:
    """Rnd function."""
    import random
    return random.random()


def builtin_randomize(interp: Interpreter, seed: Any = None) -> None:
    """Randomize statement."""
    import random
    if seed is not None:
        random.seed(int(interp._to_number(seed)))
    else:
        random.seed()


def builtin_sgn(interp: Interpreter, value: Any) -> int:
    """Sgn function - returns the sign of a number."""
    n = interp._to_number(value)
    if n > 0:
        return 1
    elif n < 0:
        return -1
    return 0


def builtin_log(interp: Interpreter, value: Any) -> float:
    """Log function - natural logarithm."""
    n = interp._to_number(value)
    if n <= 0:
        raise VBScriptError('Invalid procedure call or argument')
    return math.log(n)


def builtin_exp(interp: Interpreter, value: Any) -> float:
    """Exp function - e raised to a power."""
    return math.exp(interp._to_number(value))


def builtin_cos(interp: Interpreter, value: Any) -> float:
    """Cos function - cosine."""
    return math.cos(interp._to_number(value))


def builtin_sin(interp: Interpreter, value: Any) -> float:
    """Sin function - sine."""
    return math.sin(interp._to_number(value))


def builtin_tan(interp: Interpreter, value: Any) -> float:
    """Tan function - tangent."""
    return math.tan(interp._to_number(value))


def builtin_atn(interp: Interpreter, value: Any) -> float:
    """Atn function - arctangent."""
    return math.atan(interp._to_number(value))


# ---------------------------------------------------------------------------
#  Date/time functions
# ---------------------------------------------------------------------------

def _ensure_date(interp: Interpreter, value: Any) -> VBScriptDate:
    """Coerce a value to VBScriptDate."""
    if isinstance(value, VBScriptDate):
        return value
    if isinstance(value, str):
        return VBScriptDate.from_string(value)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return VBScriptDate(float(value))
    raise VBScriptError('Type mismatch')


def builtin_now(interp: Interpreter) -> VBScriptDate:
    """Now function - current date and time."""
    from datetime import datetime as _dt
    return VBScriptDate.from_datetime(_dt.now())


def builtin_date_func(interp: Interpreter) -> VBScriptDate:
    """Date function - current date (no time component)."""
    from datetime import datetime as _dt
    today = _dt.now().replace(hour=0, minute=0, second=0, microsecond=0)
    return VBScriptDate.from_datetime(today)


def builtin_time_func(interp: Interpreter) -> VBScriptDate:
    """Time function - current time (no date component)."""
    from datetime import datetime as _dt
    now = _dt.now()
    return VBScriptDate.from_time_parts(now.hour, now.minute, now.second)


def builtin_timer(interp: Interpreter) -> float:
    """Timer function - seconds elapsed since midnight."""
    from datetime import datetime as _dt
    now = _dt.now()
    return now.hour * 3600 + now.minute * 60 + now.second + now.microsecond / 1e6


def builtin_year(interp: Interpreter, date: Any) -> int:
    """Year function - extract year from a date."""
    return _ensure_date(interp, date).year


def builtin_month(interp: Interpreter, date: Any) -> int:
    """Month function - extract month from a date."""
    return _ensure_date(interp, date).month


def builtin_day(interp: Interpreter, date: Any) -> int:
    """Day function - extract day from a date."""
    return _ensure_date(interp, date).day


def builtin_hour(interp: Interpreter, time: Any) -> int:
    """Hour function - extract hour from a date/time."""
    return _ensure_date(interp, time).hour


def builtin_minute(interp: Interpreter, time: Any) -> int:
    """Minute function - extract minute from a date/time."""
    return _ensure_date(interp, time).minute


def builtin_second(interp: Interpreter, time: Any) -> int:
    """Second function - extract second from a date/time."""
    return _ensure_date(interp, time).second


def builtin_weekday(interp: Interpreter, date: Any, first_day: int = 1) -> int:
    """Weekday function - day of week (1=Sunday by default)."""
    d = _ensure_date(interp, date)
    # VBScript weekday: 1=Sunday..7=Saturday (when firstdayofweek=vbSunday=1)
    vbs_wd = d.weekday  # already 1=Sun..7=Sat
    if first_day == 1:
        return vbs_wd
    # Rotate: result = ((vbs_wd - first_day) % 7) + 1
    return ((vbs_wd - first_day) % 7) + 1


def builtin_weekdayname(
    interp: Interpreter, weekday: Any, abbreviate: Any = False, first_day: int = 1,
) -> str:
    """WeekdayName function - name of a weekday."""
    wd = int(interp._to_number(weekday))
    abbr = interp._to_boolean(abbreviate)
    # Map weekday number (1=Sun..7=Sat when firstdayofweek=1) to name
    # Adjust for first_day
    actual_wd = ((wd - 1 + (first_day - 1)) % 7)  # 0=Sun..6=Sat
    names = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
    abbr_names = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']
    if actual_wd < 0 or actual_wd > 6:
        raise VBScriptError('Invalid procedure call or argument')
    return abbr_names[actual_wd] if abbr else names[actual_wd]


def builtin_monthname(interp: Interpreter, month: Any, abbreviate: Any = False) -> str:
    """MonthName function - name of a month."""
    m = int(interp._to_number(month))
    abbr = interp._to_boolean(abbreviate)
    if m < 1 or m > 12:
        raise VBScriptError('Invalid procedure call or argument')
    names = [
        '', 'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December',
    ]
    abbr_names = [
        '', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
        'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
    ]
    return abbr_names[m] if abbr else names[m]


def builtin_dateserial(interp: Interpreter, year: Any, month: Any, day: Any) -> VBScriptDate:
    """DateSerial function - create a date from year, month, day."""
    from datetime import datetime as _dt
    y = int(interp._to_number(year))
    m = int(interp._to_number(month))
    d = int(interp._to_number(day))
    # VBScript DateSerial handles overflow: DateSerial(2000, 13, 1) = 1/1/2001
    # Use timedelta arithmetic from a base date
    base = _dt(y, 1, 1)
    from datetime import timedelta
    result = base + timedelta(days=d - 1)
    # Adjust month: add (m-1) months
    if m != 1:
        target_month = base.month + (m - 1)
        target_year = base.year
        while target_month > 12:
            target_month -= 12
            target_year += 1
        while target_month < 1:
            target_month += 12
            target_year -= 1
        # Set to first of target month, then add days
        result = _dt(target_year, target_month, 1) + timedelta(days=d - 1)
    return VBScriptDate.from_datetime(result)


def builtin_datevalue(interp: Interpreter, date: Any) -> VBScriptDate:
    """DateValue function - extract the date portion (strip time)."""
    d = _ensure_date(interp, date)
    return VBScriptDate(float(int(d.serial)))


def builtin_timeserial(interp: Interpreter, hour: Any, minute: Any, second: Any) -> VBScriptDate:
    """TimeSerial function - create a time from hour, minute, second."""
    h = int(interp._to_number(hour))
    m = int(interp._to_number(minute))
    s = int(interp._to_number(second))
    total_seconds = h * 3600 + m * 60 + s
    from ..runtime import _SECONDS_PER_DAY
    return VBScriptDate(total_seconds / _SECONDS_PER_DAY)


def builtin_timevalue(interp: Interpreter, time: Any) -> VBScriptDate:
    """TimeValue function - extract the time portion (strip date)."""
    d = _ensure_date(interp, time)
    frac = d.serial - int(d.serial)
    return VBScriptDate(frac)


def builtin_dateadd(interp: Interpreter, interval: Any, number: Any, date: Any) -> VBScriptDate:
    """DateAdd function - add an interval to a date."""
    from datetime import timedelta
    intv = interp._to_string(interval).lower()
    n = int(interp._to_number(number))
    d = _ensure_date(interp, date)
    dt = d.to_datetime()

    if intv == 'yyyy':
        dt = dt.replace(year=dt.year + n)
    elif intv == 'q':
        # Quarter: add n*3 months
        m = dt.month + n * 3
        y = dt.year
        while m > 12:
            m -= 12
            y += 1
        while m < 1:
            m += 12
            y -= 1
        dt = dt.replace(year=y, month=m)
    elif intv == 'm':
        m = dt.month + n
        y = dt.year
        while m > 12:
            m -= 12
            y += 1
        while m < 1:
            m += 12
            y -= 1
        # Handle day overflow (e.g. Jan 31 + 1 month)
        import calendar
        max_day = calendar.monthrange(y, m)[1]
        day = min(dt.day, max_day)
        dt = dt.replace(year=y, month=m, day=day)
    elif intv == 'y' or intv == 'd':
        dt = dt + timedelta(days=n)
    elif intv == 'w':
        dt = dt + timedelta(weeks=0, days=n)
    elif intv == 'ww':
        dt = dt + timedelta(weeks=n)
    elif intv == 'h':
        dt = dt + timedelta(hours=n)
    elif intv == 'n':
        dt = dt + timedelta(minutes=n)
    elif intv == 's':
        dt = dt + timedelta(seconds=n)
    else:
        raise VBScriptError('Invalid procedure call or argument')

    return VBScriptDate.from_datetime(dt)


def builtin_datediff(
    interp: Interpreter, interval: Any, date1: Any, date2: Any,
    first_day: int = 1, first_week: int = 1,
) -> int:
    """DateDiff function - difference between two dates."""
    intv = interp._to_string(interval).lower()
    d1 = _ensure_date(interp, date1).to_datetime()
    d2 = _ensure_date(interp, date2).to_datetime()

    if intv == 'yyyy':
        return d2.year - d1.year
    elif intv == 'q':
        return (d2.year * 4 + (d2.month - 1) // 3) - (d1.year * 4 + (d1.month - 1) // 3)
    elif intv == 'm':
        return (d2.year * 12 + d2.month) - (d1.year * 12 + d1.month)
    elif intv in ('y', 'd'):
        delta = d2 - d1
        return delta.days
    elif intv == 'w':
        delta = d2 - d1
        return delta.days // 7
    elif intv == 'ww':
        delta = d2 - d1
        return delta.days // 7
    elif intv == 'h':
        delta = d2 - d1
        return int(delta.total_seconds() // 3600)
    elif intv == 'n':
        delta = d2 - d1
        return int(delta.total_seconds() // 60)
    elif intv == 's':
        delta = d2 - d1
        return int(delta.total_seconds())
    else:
        raise VBScriptError('Invalid procedure call or argument')


def builtin_datepart(
    interp: Interpreter, interval: Any, date: Any,
    first_day: int = 1, first_week: int = 1,
) -> int:
    """DatePart function - extract a part of a date."""
    intv = interp._to_string(interval).lower()
    d = _ensure_date(interp, date).to_datetime()

    if intv == 'yyyy':
        return d.year
    elif intv == 'q':
        return (d.month - 1) // 3 + 1
    elif intv == 'm':
        return d.month
    elif intv == 'y':
        return d.timetuple().tm_yday
    elif intv == 'd':
        return d.day
    elif intv == 'w':
        # Weekday, respecting first_day
        vbs_wd = _ensure_date(interp, date).weekday
        if first_day == 1:
            return vbs_wd
        return ((vbs_wd - first_day) % 7) + 1
    elif intv == 'ww':
        # Week of year
        return d.isocalendar()[1]
    elif intv == 'h':
        return d.hour
    elif intv == 'n':
        return d.minute
    elif intv == 's':
        return d.second
    else:
        raise VBScriptError('Invalid procedure call or argument')


def builtin_formatdatetime(interp: Interpreter, date: Any, format_type: int = 0) -> str:
    """FormatDateTime function - format a date/time value."""
    d = _ensure_date(interp, date)
    dt = d.to_datetime()
    fmt = int(format_type)

    if fmt == 0:  # vbGeneralDate
        return str(d)
    elif fmt == 1:  # vbLongDate
        return dt.strftime('%A, %B %d, %Y')
    elif fmt == 2:  # vbShortDate
        return d._format_date(dt)
    elif fmt == 3:  # vbLongTime
        return d._format_time(dt)
    elif fmt == 4:  # vbShortTime
        return f'{dt.hour:02d}:{dt.minute:02d}'
    else:
        raise VBScriptError('Invalid procedure call or argument')


# ---------------------------------------------------------------------------
#  Object functions
# ---------------------------------------------------------------------------

def builtin_createobject(
    interp: Interpreter, class_name: str, server_name: str = None,
) -> Any:
    """CreateObject function - creates COM objects (simplified)."""
    class_lower = class_name.lower()
    if class_lower == 'scripting.dictionary':
        return VBScriptDictionary()
    if class_lower == 'scripting.filesystemobject':
        return VBScriptFileSystemObject()
    return {'_class': class_name}


def builtin_getobject(
    interp: Interpreter, path_name: str = None, class_name: str = None,
) -> Any:
    """GetObject function (simplified)."""
    return {'_path': path_name, '_class': class_name}


# ---------------------------------------------------------------------------
#  Array functions
# ---------------------------------------------------------------------------

def builtin_ubound(interp: Interpreter, array: Any, dimension: int = 1) -> int:
    """UBound function - returns the upper bound of an array dimension."""
    if isinstance(array, VBScriptArray):
        return array.ubound(dimension)
    elif isinstance(array, list):
        if dimension != 1:
            raise VBScriptError('Subscript out of range')
        return len(array) - 1
    else:
        raise VBScriptError('Type mismatch: UBound requires an array')


def builtin_lbound(interp: Interpreter, array: Any, dimension: int = 1) -> int:
    """LBound function - returns the lower bound of an array dimension."""
    if isinstance(array, VBScriptArray):
        return array.lbound(dimension)
    elif isinstance(array, list):
        if dimension != 1:
            raise VBScriptError('Subscript out of range')
        return 0
    else:
        raise VBScriptError('Type mismatch: LBound requires an array')


def builtin_array(interp: Interpreter, *args: Any) -> VBScriptArray:
    """Array function - creates a variant array from the given values."""
    if len(args) == 0:
        return VBScriptArray([-1], is_dynamic=True)
    arr = VBScriptArray([len(args) - 1], is_dynamic=False)
    for i, val in enumerate(args):
        arr.set_element([i], val)
    return arr


# ---------------------------------------------------------------------------
#  Dynamic code execution (Execute, ExecuteGlobal, Eval)
# ---------------------------------------------------------------------------

def _parse_dynamic_program(source: str):
    from ..parser import parse as vbs_parse

    try:
        return vbs_parse(source)
    except UnexpectedInput as exc:
        raise VBScriptError('Syntax error') from exc


def builtin_eval(interp: 'Interpreter', expr_string: Any) -> Any:
    """Eval function - evaluate a VBScript expression string and return its value."""
    from ..ast_nodes import AssignmentStatement

    code_str = interp._to_string(expr_string)
    wrapper = f"__pybasil_eval__ = {code_str}"
    program = _parse_dynamic_program(wrapper)
    if len(program.statements) != 1:
        raise VBScriptError('Syntax error')

    statement = program.statements[0]
    if (
        not isinstance(statement, AssignmentStatement)
        or statement.variable.lower() != '__pybasil_eval__'
    ):
        raise VBScriptError('Syntax error')

    return interp._evaluate(statement.expression)


def builtin_execute(interp: 'Interpreter', code_string: Any) -> None:
    """Execute statement - parse and execute VBScript code in the current scope."""
    code_str = interp._to_string(code_string)
    if not code_str.strip():
        return None

    program = _parse_dynamic_program(code_str)
    for stmt in program.statements:
        interp._execute_with_error_handling(stmt)
    return None


def builtin_executeglobal(interp: 'Interpreter', code_string: Any) -> None:
    """ExecuteGlobal statement - parse and execute VBScript code in the global scope."""
    code_str = interp._to_string(code_string)
    if not code_str.strip():
        return None

    program = _parse_dynamic_program(code_str)
    old_env = interp._environment
    old_definition_scope_is_global = interp._definition_scope_is_global
    old_instance = interp._current_instance
    interp._environment = interp._global_environment
    interp._definition_scope_is_global = True
    interp._current_instance = None
    try:
        for stmt in program.statements:
            interp._execute_with_error_handling(stmt)
    finally:
        interp._current_instance = old_instance
        interp._definition_scope_is_global = old_definition_scope_is_global
        interp._environment = old_env
    return None


# ---------------------------------------------------------------------------
#  Registration helper
# ---------------------------------------------------------------------------

def get_builtin_table(interp: Interpreter) -> dict:
    """Return {lowercase_name: callable} mapping for all built-in functions.

    Each callable is a closure that pre-binds *interp* as the first argument
    so the dispatch site can simply call ``builtin(*args)``.
    """
    def _bind(fn):
        return lambda *args: fn(interp, *args)

    return {
        'msgbox': _bind(builtin_msgbox),
        'inputbox': _bind(builtin_inputbox),
        'len': _bind(builtin_len),
        'left': _bind(builtin_left),
        'right': _bind(builtin_right),
        'mid': _bind(builtin_mid),
        'trim': _bind(builtin_trim),
        'ltrim': _bind(builtin_ltrim),
        'rtrim': _bind(builtin_rtrim),
        'ucase': _bind(builtin_ucase),
        'lcase': _bind(builtin_lcase),
        'instr': _bind(builtin_instr),
        'replace': _bind(builtin_replace),
        'split': _bind(builtin_split),
        'join': _bind(builtin_join),
        'instrrev': _bind(builtin_instrrev),
        'strcomp': _bind(builtin_strcomp),
        'string': _bind(builtin_string),
        'space': _bind(builtin_space),
        'strreverse': _bind(builtin_strreverse),
        'asc': _bind(builtin_asc),
        'ascw': _bind(builtin_ascw),
        'chr': _bind(builtin_chr),
        'chrw': _bind(builtin_chrw),
        'hex': _bind(builtin_hex),
        'oct': _bind(builtin_oct),
        'cstr': _bind(builtin_cstr),
        'cint': _bind(builtin_cint),
        'clng': _bind(builtin_clng),
        'cdbl': _bind(builtin_cdbl),
        'cbool': _bind(builtin_cbool),
        'cdate': _bind(builtin_cdate),
        'isnumeric': _bind(builtin_isnumeric),
        'isarray': _bind(builtin_isarray),
        'isdate': _bind(builtin_isdate),
        'isempty': _bind(builtin_isempty),
        'isnull': _bind(builtin_isnull),
        'isobject': _bind(builtin_isobject),
        'typename': _bind(builtin_typename),
        'vartype': _bind(builtin_vartype),
        'abs': _bind(builtin_abs),
        'sqr': _bind(builtin_sqr),
        'int': _bind(builtin_int),
        'fix': _bind(builtin_fix),
        'round': _bind(builtin_round),
        'rnd': _bind(builtin_rnd),
        'randomize': _bind(builtin_randomize),
        'sgn': _bind(builtin_sgn),
        'log': _bind(builtin_log),
        'exp': _bind(builtin_exp),
        'cos': _bind(builtin_cos),
        'sin': _bind(builtin_sin),
        'tan': _bind(builtin_tan),
        'atn': _bind(builtin_atn),
        'now': _bind(builtin_now),
        'date': _bind(builtin_date_func),
        'time': _bind(builtin_time_func),
        'timer': _bind(builtin_timer),
        'year': _bind(builtin_year),
        'month': _bind(builtin_month),
        'day': _bind(builtin_day),
        'hour': _bind(builtin_hour),
        'minute': _bind(builtin_minute),
        'second': _bind(builtin_second),
        'weekday': _bind(builtin_weekday),
        'weekdayname': _bind(builtin_weekdayname),
        'monthname': _bind(builtin_monthname),
        'dateserial': _bind(builtin_dateserial),
        'datevalue': _bind(builtin_datevalue),
        'timeserial': _bind(builtin_timeserial),
        'timevalue': _bind(builtin_timevalue),
        'dateadd': _bind(builtin_dateadd),
        'datediff': _bind(builtin_datediff),
        'datepart': _bind(builtin_datepart),
        'formatdatetime': _bind(builtin_formatdatetime),
        'createobject': _bind(builtin_createobject),
        'getobject': _bind(builtin_getobject),
        'ubound': _bind(builtin_ubound),
        'lbound': _bind(builtin_lbound),
        'array': _bind(builtin_array),
        'eval': _bind(builtin_eval),
        'execute': _bind(builtin_execute),
        'executeglobal': _bind(builtin_executeglobal),
    }
