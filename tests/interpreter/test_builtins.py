"""Tests for the VBScript interpreter."""

import pytest
import io
from pybasil import (
    Interpreter,
    parse,
    run,
    VBScriptArray,
    VBScriptError,
)


class TestInterpreterBuiltins:
    """Test built-in functions."""

    def test_len(self):
        program = parse('x = Len("Hello")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 5

    def test_left(self):
        program = parse('x = Left("Hello", 3)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'Hel'

    def test_right(self):
        program = parse('x = Right("Hello", 3)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'llo'

    def test_mid(self):
        program = parse('x = Mid("Hello", 2, 3)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'ell'

    def test_trim(self):
        program = parse('x = Trim("  Hello  ")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'Hello'

    def test_ucase(self):
        program = parse('x = UCase("hello")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'HELLO'

    def test_lcase(self):
        program = parse('x = LCase("HELLO")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'hello'

    def test_cstr(self):
        program = parse('x = CStr(42)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == '42'

    def test_cint(self):
        program = parse('x = CInt(3.7)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 4

    def test_cdbl(self):
        program = parse('x = CDbl("3.14")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 3.14

    def test_cbool(self):
        program = parse('x = CBool(1)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_abs(self):
        program = parse('x = Abs(-5)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 5

    def test_sqr(self):
        program = parse('x = Sqr(16)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 4.0

    def test_int(self):
        program = parse('x = Int(3.7)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 3

    def test_fix(self):
        program = parse('x = Fix(3.7)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 3

    def test_round(self):
        program = parse('x = Round(3.14159, 2)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 3.14

    def test_isnumeric_true(self):
        program = parse('x = IsNumeric("123")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_isnumeric_false(self):
        program = parse('x = IsNumeric("abc")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is False

    def test_isempty(self):
        program = parse('x = IsEmpty(Empty)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_isnull(self):
        program = parse('x = IsNull(Null)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_typename_string(self):
        program = parse('x = TypeName("Hello")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'String'

    def test_typename_integer(self):
        program = parse('x = TypeName(42)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'Integer'

    def test_typename_integer_addition(self):
        program = parse('x = TypeName(5 + 3)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'Integer'

    def test_integer_arithmetic_preserves_type(self):
        program = parse('x = 5 + 3')
        interpreter = Interpreter()
        interpreter.interpret(program)
        result = interpreter._environment.get('x')
        assert result == 8
        assert isinstance(result, int)

    def test_replace_default_count(self):
        program = parse('x = Replace("aaa", "a", "b")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'bbb'

    def test_replace_with_count(self):
        program = parse('x = Replace("aaa", "a", "b", 1, 2)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'bba'

    def test_replace_no_match(self):
        program = parse('x = Replace("hello", "x", "y")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'hello'

    def test_replace_all_occurrences(self):
        program = parse('x = Replace("abcabc", "a", "x")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'xbcxbc'

    def test_instr_with_start_parameter(self):
        program = parse('x = InStr(4, "Hello", "l")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 4

    def test_instr_case_insensitive_compare(self):
        program = parse('x = InStr(1, "Hello", "h", 1)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 1

    def test_instr_two_args(self):
        program = parse('x = InStr("Hello", "l")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 3

    def test_left_negative_length_raises_error(self):
        program = parse('x = Left("Hello", -1)')
        interpreter = Interpreter()
        with pytest.raises(VBScriptError, match='Invalid procedure call or argument'):
            interpreter.interpret(program)

    def test_right_negative_length_raises_error(self):
        program = parse('x = Right("Hello", -1)')
        interpreter = Interpreter()
        with pytest.raises(VBScriptError, match='Invalid procedure call or argument'):
            interpreter.interpret(program)

    def test_mid_start_zero_raises_error(self):
        program = parse('x = Mid("Hello", 0, 3)')
        interpreter = Interpreter()
        with pytest.raises(VBScriptError, match='Invalid procedure call or argument'):
            interpreter.interpret(program)

    def test_mid_start_negative_raises_error(self):
        program = parse('x = Mid("Hello", -1, 3)')
        interpreter = Interpreter()
        with pytest.raises(VBScriptError, match='Invalid procedure call or argument'):
            interpreter.interpret(program)

    def test_split_array_access(self):
        code = '''Dim arr
arr = Split("a,b,c", ",")
x = arr(1)
'''
        program = parse(code)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'b'

class TestReplaceParams:
    """Tests for Replace builtin with start, count, and compare params."""

    def test_replace_with_count(self):
        program = parse('x = Replace("aaa", "a", "b", 1, 2)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'bba'

    def test_replace_with_start(self):
        program = parse('x = Replace("hello world", "o", "0", 5)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == '0 w0rld'

    def test_replace_case_insensitive(self):
        program = parse('x = Replace("Hello HELLO", "hello", "bye", 1, -1, 1)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'bye bye'

    def test_replace_case_insensitive_with_count(self):
        program = parse('x = Replace("aAbBaA", "a", "x", 1, 2, 1)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'xxbBaA'

class TestSplitCount:
    """Tests for Split with count parameter."""

    def test_split_with_count(self):
        code = '''Dim arr
arr = Split("a-b-c-d", "-", 2)
x = UBound(arr)
y = arr(0)
z = arr(1)
'''
        program = parse(code)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 1
        assert interpreter._environment.get('y') == 'a'
        assert interpreter._environment.get('z') == 'b-c-d'

    def test_split_without_count(self):
        code = '''Dim arr
arr = Split("a-b-c-d", "-")
x = UBound(arr)
'''
        program = parse(code)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 3

    def test_split_with_count_minus_one(self):
        code = '''Dim arr
arr = Split("a-b-c", "-", -1)
x = UBound(arr)
'''
        program = parse(code)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 2

class TestArrayBuiltin:
    """Tests for the Array() builtin function."""

    def test_array_no_args_returns_empty_array(self):
        from pybasil.builtins import builtin_array
        interpreter = Interpreter()
        result = builtin_array(interpreter)
        assert isinstance(result, VBScriptArray)
        assert result.ubound() == -1

    def test_array_with_args(self):
        code = '''Dim a
a = Array(10, 20, 30)
x = a(0)
y = a(2)
u = UBound(a)
'''
        program = parse(code)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 10
        assert interpreter._environment.get('y') == 30
        assert interpreter._environment.get('u') == 2

class TestHexOctalLiterals:
    """Tests for &H hex and &O octal literal support."""

    def test_hex_literal(self):
        program = parse('x = &HFF')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 255

    def test_hex_literal_lowercase(self):
        program = parse('x = &hff')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 255

    def test_hex_literal_with_trailing_ampersand(self):
        program = parse('x = &HFF&')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 255

    def test_octal_literal(self):
        program = parse('x = &O77')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 63

    def test_hex_in_expression(self):
        program = parse('x = &H10 + 1')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 17

    def test_hex_zero(self):
        program = parse('x = &H0')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 0

class TestNumericFastPaths:
    """Test fast-paths in type coercion and builtin functions for numeric types."""

    def test_to_number_int_fast_path(self):
        output = io.StringIO()
        program = parse('x = 42\ny = x + 1\nWScript.Echo y')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        assert output.getvalue().strip() == '43'

    def test_to_number_float_fast_path(self):
        program = parse('x = 3.5\ny = x + 1.5')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('y') == 5.0

    def test_to_number_bool_still_works(self):
        output = io.StringIO()
        program = parse('x = True\ny = x + 1\nWScript.Echo y')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        assert output.getvalue().strip() == '0'

    def test_to_string_str_fast_path(self):
        output = io.StringIO()
        program = parse('x = "hello" & " world"\nWScript.Echo x')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        assert output.getvalue().strip() == 'hello world'

    def test_to_string_int_fast_path(self):
        program = parse('x = CStr(42)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == '42'

    def test_to_boolean_bool_fast_path(self):
        output = io.StringIO()
        program = parse('If True Then\nWScript.Echo "yes"\nEnd If')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        assert output.getvalue().strip() == 'yes'

    def test_to_boolean_int_fast_path(self):
        output = io.StringIO()
        program = parse('x = 1\nIf x Then\nWScript.Echo "yes"\nEnd If')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        assert output.getvalue().strip() == 'yes'

    def test_cint_int_fast_path(self):
        program = parse('x = CInt(42)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 42

    def test_cint_string_conversion(self):
        program = parse('x = CInt("123")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 123

    def test_cdbl_float_fast_path(self):
        program = parse('x = CDbl(3.14)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 3.14

    def test_cdbl_int_to_float(self):
        program = parse('x = CDbl(42)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 42.0
        assert isinstance(interp._environment.get('x'), float)

    def test_cdbl_string_conversion(self):
        program = parse('x = CDbl("3.14")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 3.14

    def test_instr_string_fast_path(self):
        program = parse('x = InStr(1, "Hello World", "World")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 7

class TestCscriptCompatibility:
    """Tests for cscript compatibility fixes."""

    def test_len_number(self):
        program = parse('x = Len(12345)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 5

    def test_typename_long(self):
        program = parse('x = TypeName(100000)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'Long'

    def test_typename_integer(self):
        program = parse('x = TypeName(42)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'Integer'

    def test_vartype_nothing(self):
        program = parse('x = VarType(Nothing)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 9

    def test_vartype_dictionary(self):
        program = parse('''
        Set d = CreateObject("Scripting.Dictionary")
        x = VarType(d)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 9

    def test_isobject_nothing(self):
        program = parse('x = IsObject(Nothing)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') is True

    def test_hex_negative_one_integer_range(self):
        program = parse('x = Hex(-1)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'FFFF'

    def test_hex_negative_long_range(self):
        program = parse('x = Hex(-40000)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'FFFF63C0'

    def test_round_banker(self):
        program = parse('x = Round(2.55, 1)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 2.6

    def test_cstr_boolean(self):
        program = parse('x = CStr(True)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'True'

    def test_dict_removeall(self):
        output = io.StringIO()
        run('''
        Set d = CreateObject("Scripting.Dictionary")
        d.Add "a", 1
        d.Add "b", 2
        d.RemoveAll
        WScript.Echo d.Count
        ''', output_stream=output)
        assert output.getvalue().strip() == '0'

    def test_for_no_step_no_countdown(self):
        output = io.StringIO()
        run('''
        Dim s
        s = ""
        For i = 5 To 1
            s = s & i
        Next
        WScript.Echo s
        ''', output_stream=output)
        assert output.getvalue().strip() == ''

    def test_error_number_const_reassign(self):
        output = io.StringIO()
        run('''
        On Error Resume Next
        Const X = 42
        X = 99
        WScript.Echo Err.Number
        ''', output_stream=output)
        assert output.getvalue().strip() == '501'

    def test_error_number_undefined_var(self):
        output = io.StringIO()
        run('''
        Option Explicit
        On Error Resume Next
        y = 42
        WScript.Echo Err.Number
        ''', output_stream=output)
        assert output.getvalue().strip() == '500'


class TestEmptyParensCalls:
    """Tests for zero-argument calls with empty parentheses."""

    def test_array_empty_parens(self):
        code = 'x = Array()\n'
        program = parse(code)
        interpreter = Interpreter()
        interpreter.interpret(program)
        result = interpreter._environment.get('x')
        assert isinstance(result, VBScriptArray)

    def test_dictionary_keys_empty_parens(self):
        output = io.StringIO()
        run(
            '''Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
d.Add "b", 2
Dim k
Set k = d.Keys()
WScript.Echo k(0)
WScript.Echo k(1)
''',
            output_stream=output,
        )
        assert output.getvalue().strip() == 'a\nb'

    def test_dictionary_items_empty_parens(self):
        output = io.StringIO()
        run(
            '''Set d = CreateObject("Scripting.Dictionary")
d.Add "x", 42
Dim it
Set it = d.Items()
WScript.Echo it(0)
''',
            output_stream=output,
        )
        assert output.getvalue().strip() == '42'
