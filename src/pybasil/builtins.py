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

from .runtime import (
    VBScriptError,
    VBScriptObject,
    VBScriptNothing,
    VBScriptEmpty,
    VBScriptNull,
    VBScriptArray,
    VBScriptDictionary,
    VBScriptClassInstance,
    WScriptObject,
)

if TYPE_CHECKING:
    from .interpreter import Interpreter


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
    raise VBScriptError('Type mismatch: Len requires a string')


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


def builtin_cdate(interp: Interpreter, value: Any) -> Any:
    """CDate function (simplified)."""
    return interp._to_string(value)


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
    """IsDate function (simplified)."""
    return False


def builtin_isempty(interp: Interpreter, value: Any) -> bool:
    """IsEmpty function."""
    return isinstance(value, VBScriptEmpty)


def builtin_isnull(interp: Interpreter, value: Any) -> bool:
    """IsNull function."""
    return isinstance(value, VBScriptNull)


def builtin_isobject(interp: Interpreter, value: Any) -> bool:
    """IsObject function."""
    return (
        isinstance(value, (WScriptObject, VBScriptObject))
        or value is not None
        and not isinstance(
            value,
            (str, int, float, bool, VBScriptEmpty, VBScriptNull, VBScriptNothing),
        )
    )


def builtin_typename(interp: Interpreter, value: Any) -> str:
    """TypeName function."""
    if isinstance(value, VBScriptEmpty):
        return 'Empty'
    if isinstance(value, VBScriptNull):
        return 'Null'
    if isinstance(value, VBScriptNothing):
        return 'Nothing'
    if isinstance(value, VBScriptClassInstance):
        return value.class_name
    if isinstance(value, VBScriptArray):
        return 'Variant()'
    if isinstance(value, bool):
        return 'Boolean'
    if isinstance(value, int):
        return 'Integer'
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
    """Round function."""
    return round(interp._to_number(value), int(decimals))


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
    from .parser import parse as vbs_parse

    try:
        return vbs_parse(source)
    except UnexpectedInput as exc:
        raise VBScriptError('Syntax error') from exc


def builtin_eval(interp: 'Interpreter', expr_string: Any) -> Any:
    """Eval function - evaluate a VBScript expression string and return its value."""
    from .ast_nodes import AssignmentStatement

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
    interp._environment = interp._global_environment
    interp._definition_scope_is_global = True
    try:
        for stmt in program.statements:
            interp._execute_with_error_handling(stmt)
    finally:
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
        'createobject': _bind(builtin_createobject),
        'getobject': _bind(builtin_getobject),
        'ubound': _bind(builtin_ubound),
        'lbound': _bind(builtin_lbound),
        'array': _bind(builtin_array),
        'eval': _bind(builtin_eval),
        'execute': _bind(builtin_execute),
        'executeglobal': _bind(builtin_executeglobal),
    }
