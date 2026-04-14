"""Tests for the VBScript interpreter."""

import pytest
import io
from pybasil import (
    Interpreter,
    parse,
    run,
    VBScriptArray,
    VBScriptError,
    EMPTY,
    NULL,
    NOTHING,
)


class TestInterpreterLiterals:
    """Test evaluation of literal values."""

    def test_integer_literal(self):
        program = parse('x = 42')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 42

    def test_float_literal(self):
        program = parse('x = 3.14')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 3.14

    def test_string_literal(self):
        program = parse('x = "Hello"')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'Hello'

    def test_boolean_true(self):
        program = parse('x = True')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_boolean_false(self):
        program = parse('x = False')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is False

    def test_nothing_literal(self):
        program = parse('x = Nothing')
        interpreter = Interpreter()
        interpreter.interpret(program)

        assert interpreter._environment.get('x') == NOTHING

    def test_empty_literal(self):
        program = parse('x = Empty')
        interpreter = Interpreter()
        interpreter.interpret(program)
        from pybasil import EMPTY

        assert interpreter._environment.get('x') == EMPTY

    def test_null_literal(self):
        program = parse('x = Null')
        interpreter = Interpreter()
        interpreter.interpret(program)
        from pybasil import NULL

        assert interpreter._environment.get('x') == NULL


class TestInterpreterVariables:
    """Test variable handling."""

    def test_variable_assignment(self):
        program = parse('x = 42')
        interpreter = Interpreter()
        interpreter.interpret(program)
        # Check that x is accessible
        assert interpreter._environment.get('x') == 42

    def test_variable_lookup(self):
        program = parse("""
            x = 10
            y = x
        """)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('y') == 10

    def test_implicit_variable_creation(self):
        program = parse('y = x')
        interpreter = Interpreter()
        interpreter.interpret(program)
        # x should be Empty (implicit creation)
        assert interpreter._environment.get('x') == EMPTY

    def test_dim_statement(self):
        program = parse('Dim x, y, z')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.exists('x')
        assert interpreter._environment.exists('y')
        assert interpreter._environment.exists('z')

    def test_set_statement(self):
        program = parse('Set obj = CreateObject("Test.Object")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.exists('obj')

    def test_case_insensitive_variables(self):
        program = parse("""
            x = 42
            Y = X
        """)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('y') == 42


class TestInterpreterArithmetic:
    """Test arithmetic operations."""

    def test_addition(self):
        program = parse('x = 5 + 3')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 8

    def test_subtraction(self):
        program = parse('x = 10 - 4')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 6

    def test_multiplication(self):
        program = parse('x = 6 * 7')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 42

    def test_division(self):
        program = parse('x = 15 / 3')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 5.0

    def test_integer_division(self):
        program = parse('x = 17 \\ 5')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 3

    def test_modulo(self):
        program = parse('x = 17 Mod 5')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 2

    def test_integer_division_negative(self):
        program = parse('x = -5 \\ 2')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == -2

    def test_modulo_negative(self):
        program = parse('x = -5 Mod 3')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == -2

    def test_exponentiation(self):
        program = parse('x = 2 ^ 10')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 1024

    def test_exponentiation_right_associative(self):
        program = parse('x = 2 ^ 3 ^ 2')
        interpreter = Interpreter()
        interpreter.interpret(program)
        # Right-associative: 2 ^ (3 ^ 2) = 2 ^ 9 = 512
        assert interpreter._environment.get('x') == 512

    def test_negation(self):
        program = parse('x = -5')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == -5

    def test_unary_plus(self):
        program = parse('x = +5')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 5

    def test_complex_expression(self):
        program = parse('x = 2 + 3 * 4 - 1')
        interpreter = Interpreter()
        interpreter.interpret(program)
        # 2 + 12 - 1 = 13
        assert interpreter._environment.get('x') == 13

    def test_parentheses(self):
        program = parse('x = (2 + 3) * 4')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 20

    def test_division_by_zero(self):
        program = parse('x = 10 / 0')
        interpreter = Interpreter()
        with pytest.raises(VBScriptError):
            interpreter.interpret(program)


class TestInterpreterStringOperations:
    """Test string operations."""

    def test_concatenation(self):
        program = parse('x = "Hello" & " " & "World"')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'Hello World'

    def test_string_number_concatenation(self):
        program = parse('x = "Value: " & 42')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'Value: 42'

    def test_number_string_addition(self):
        # In VBScript, + with a numeric string and number does arithmetic addition
        # (the string is converted to a number)
        program = parse('x = "5" + 3')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 8  # Numeric addition

    def test_escaped_double_quotes(self):
        program = parse('x = "He said ""hello"""')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'He said "hello"'

    def test_empty_escaped_quotes(self):
        program = parse('x = """"""')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == '""'


class TestInterpreterComparison:
    """Test comparison operations."""

    def test_equals_true(self):
        program = parse('x = (5 = 5)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_equals_false(self):
        program = parse('x = (5 = 6)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is False

    def test_not_equals(self):
        program = parse('x = (5 <> 6)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_less_than(self):
        program = parse('x = (3 < 5)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_greater_than(self):
        program = parse('x = (7 > 5)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_less_equal(self):
        program = parse('x = (5 <= 5)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_greater_equal(self):
        program = parse('x = (5 >= 5)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_string_comparison(self):
        program = parse('x = ("abc" < "def")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_is_operator_nothing(self):
        program = parse('x = (Nothing Is Nothing)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True


class TestInterpreterLogical:
    """Test logical operations."""

    def test_and_true(self):
        program = parse('x = True And True')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_and_false(self):
        program = parse('x = True And False')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is False

    def test_or_true(self):
        program = parse('x = False Or True')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_or_false(self):
        program = parse('x = False Or False')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is False

    def test_not_true(self):
        program = parse('x = Not True')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is False

    def test_not_false(self):
        program = parse('x = Not False')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_xor_true_true(self):
        program = parse('x = True Xor True')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 0

    def test_xor_true_false(self):
        program = parse('x = True Xor False')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == -1

    def test_xor_numeric_operands(self):
        output = io.StringIO()
        run('WScript.Echo 5 Xor 3', output_stream=output)
        assert output.getvalue().strip() == '6'

    def test_and_numeric_operands(self):
        output = io.StringIO()
        run('WScript.Echo 3 And 1', output_stream=output)
        assert output.getvalue().strip() == '1'

    def test_and_numeric_bitwise(self):
        output = io.StringIO()
        run('WScript.Echo 6 And 3', output_stream=output)
        assert output.getvalue().strip() == '2'

    def test_or_numeric_operands(self):
        output = io.StringIO()
        run('WScript.Echo 3 Or 1', output_stream=output)
        assert output.getvalue().strip() == '3'

    def test_eqv_true_true(self):
        program = parse('x = True Eqv True')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_eqv_true_false(self):
        program = parse('x = True Eqv False')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is False

    def test_imp_true_true(self):
        program = parse('x = True Imp True')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_imp_true_false(self):
        program = parse('x = True Imp False')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is False

    def test_imp_false_true(self):
        program = parse('x = False Imp True')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_imp_false_false(self):
        program = parse('x = False Imp False')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') is True

    def test_eqv_numeric_operands(self):
        output = io.StringIO()
        run('WScript.Echo 5 Eqv 3', output_stream=output)
        assert output.getvalue().strip() == '-7'

    def test_imp_numeric_operands(self):
        output = io.StringIO()
        run('WScript.Echo 5 Imp 3', output_stream=output)
        assert output.getvalue().strip() == '-5'


class TestInterpreterWScriptEcho:
    """Test WScript.Echo functionality."""

    def test_echo_string(self):
        output = io.StringIO()
        run('WScript.Echo "Hello, World!"', output_stream=output)
        assert output.getvalue().strip() == 'Hello, World!'

    def test_echo_number(self):
        output = io.StringIO()
        run('WScript.Echo 42', output_stream=output)
        assert output.getvalue().strip() == '42'

    def test_echo_multiple_args(self):
        output = io.StringIO()
        run('WScript.Echo "Hello", "World", 42', output_stream=output)
        assert output.getvalue().strip() == 'Hello World 42'

    def test_echo_boolean(self):
        output = io.StringIO()
        run('WScript.Echo True', output_stream=output)
        assert output.getvalue().strip() == 'True'

    def test_echo_variable(self):
        output = io.StringIO()
        run(
            """
            x = "Test Value"
            WScript.Echo x
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'Test Value'

    def test_echo_expression(self):
        output = io.StringIO()
        run('WScript.Echo 2 + 2', output_stream=output)
        assert output.getvalue().strip() == '4'

    def test_echo_unary_minus(self):
        output = io.StringIO()
        run('WScript.Echo -1', output_stream=output)
        assert output.getvalue().strip() == '-1'

    def test_echo_unary_plus(self):
        output = io.StringIO()
        run('WScript.Echo +5', output_stream=output)
        assert output.getvalue().strip() == '5'


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


class TestInterpreterEdgeCases:
    """Test edge cases and special behaviors."""

    def test_empty_plus_empty(self):
        program = parse("""
            Dim x, y
            z = x + y
        """)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('z') == 0

    def test_empty_in_arithmetic(self):
        program = parse('x = Empty + 5')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 5

    def test_empty_in_string_concat(self):
        program = parse('x = Empty & "text"')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'text'

    def test_null_propagation(self):
        program = parse('x = Null + 5')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == NULL

    def test_null_comparison(self):
        program = parse('x = (Null = Null)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == NULL

    def test_multiple_statements(self):
        program = parse("""
            Dim a, b, c
            a = 1
            b = 2
            c = a + b
        """)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('c') == 3

    def test_comments_ignored(self):
        output = io.StringIO()
        run(
            """
            ' This is a comment
            x = 5 ' Another comment
            WScript.Echo x
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '5'


class TestInterpreterTypeCoercion:
    """Test VBScript type coercion behavior."""

    def test_string_to_number_addition(self):
        program = parse('x = "10" + 5')
        interpreter = Interpreter()
        interpreter.interpret(program)
        # String + number = numeric addition (string is converted to number)
        assert interpreter._environment.get('x') == 15

    def test_non_numeric_string_addition_raises_error(self):
        # Non-numeric string + number should raise type mismatch
        program = parse('x = "abc" + 5')
        interpreter = Interpreter()
        with pytest.raises(VBScriptError, match='Type mismatch'):
            interpreter.interpret(program)

    def test_number_to_string_comparison(self):
        program = parse('x = (5 = "5")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        # String comparison
        assert interpreter._environment.get('x') is True

    def test_boolean_in_arithmetic(self):
        program = parse('x = True + 1')
        interpreter = Interpreter()
        interpreter.interpret(program)
        # True is -1 in VBScript
        assert interpreter._environment.get('x') == 0


class TestInterpreterIfStatement:
    """Test If/Then/Else/ElseIf statements."""

    def test_if_then_true(self):
        output = io.StringIO()
        run(
            """
            If True Then
                WScript.Echo "yes"
            End If
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'yes'

    def test_if_then_false(self):
        output = io.StringIO()
        run(
            """
            If False Then
                WScript.Echo "yes"
            End If
            WScript.Echo "done"
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'done'

    def test_if_then_else_true(self):
        output = io.StringIO()
        run(
            """
            If True Then
                WScript.Echo "then"
            Else
                WScript.Echo "else"
            End If
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'then'

    def test_if_then_else_false(self):
        output = io.StringIO()
        run(
            """
            If False Then
                WScript.Echo "then"
            Else
                WScript.Echo "else"
            End If
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'else'

    def test_if_elseif_else_first(self):
        output = io.StringIO()
        run(
            """
            x = 1
            If x = 1 Then
                WScript.Echo "one"
            ElseIf x = 2 Then
                WScript.Echo "two"
            Else
                WScript.Echo "other"
            End If
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'one'

    def test_if_elseif_else_second(self):
        output = io.StringIO()
        run(
            """
            x = 2
            If x = 1 Then
                WScript.Echo "one"
            ElseIf x = 2 Then
                WScript.Echo "two"
            Else
                WScript.Echo "other"
            End If
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'two'

    def test_if_elseif_else_fallback(self):
        output = io.StringIO()
        run(
            """
            x = 3
            If x = 1 Then
                WScript.Echo "one"
            ElseIf x = 2 Then
                WScript.Echo "two"
            Else
                WScript.Echo "other"
            End If
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'other'

    def test_if_multiple_elseif(self):
        output = io.StringIO()
        run(
            """
            x = 3
            If x = 1 Then
                WScript.Echo "one"
            ElseIf x = 2 Then
                WScript.Echo "two"
            ElseIf x = 3 Then
                WScript.Echo "three"
            ElseIf x = 4 Then
                WScript.Echo "four"
            Else
                WScript.Echo "other"
            End If
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'three'

    def test_if_nested(self):
        output = io.StringIO()
        run(
            """
            x = 5
            If x > 0 Then
                If x > 10 Then
                    WScript.Echo "big"
                Else
                    WScript.Echo "small"
                End If
            Else
                WScript.Echo "negative"
            End If
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'small'

    def test_if_case_insensitive(self):
        output = io.StringIO()
        run(
            """
            IF TRUE THEN
                wscript.echo "yes"
            END IF
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'yes'


class TestInterpreterForStatement:
    """Test For...Next statements."""

    def test_for_basic(self):
        output = io.StringIO()
        run(
            """
            For i = 1 To 3
                WScript.Echo i
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['1', '2', '3']

    def test_for_with_step(self):
        output = io.StringIO()
        run(
            """
            For i = 0 To 10 Step 2
                WScript.Echo i
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0', '2', '4', '6', '8', '10']

    def test_for_negative_step(self):
        output = io.StringIO()
        run(
            """
            For i = 5 To 1 Step -1
                WScript.Echo i
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['5', '4', '3', '2', '1']

    def test_for_exit_for(self):
        output = io.StringIO()
        run(
            """
            For i = 1 To 10
                If i = 3 Then
                    Exit For
                End If
                WScript.Echo i
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['1', '2']

    def test_for_nested(self):
        output = io.StringIO()
        run(
            """
            For i = 1 To 2
                For j = 1 To 2
                    WScript.Echo i & "," & j
                Next
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['1,1', '1,2', '2,1', '2,2']

    def test_for_variable_after_loop(self):
        program = parse("""
            For i = 1 To 5
            Next
        """)
        interpreter = Interpreter()
        interpreter.interpret(program)
        # After loop, i should be 6 (last value + step)
        assert interpreter._environment.get('i') == 6

    def test_for_empty_body(self):
        program = parse("""
            For i = 1 To 5
            Next
        """)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.exists('i')

    def test_for_default_step_countdown(self):
        output = io.StringIO()
        run(
            """
            For i = 5 To 1
                WScript.Echo i
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['5', '4', '3', '2', '1']

    def test_for_default_step_countdown_single(self):
        output = io.StringIO()
        run(
            """
            For i = 3 To 3
                WScript.Echo i
            Next
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '3'


class TestInterpreterForEachDictionary:
    """Test For Each over Scripting.Dictionary."""

    def test_for_each_dictionary_yields_keys(self):
        output = io.StringIO()
        run(
            """
            Set d = CreateObject("Scripting.Dictionary")
            d.Add "a", 100
            d.Add "b", 200
            For Each k In d
                WScript.Echo k
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['a', 'b']

    def test_for_each_dictionary_access_values_via_item(self):
        output = io.StringIO()
        run(
            """
            Set d = CreateObject("Scripting.Dictionary")
            d.Add "x", 42
            d.Add "y", 99
            For Each k In d
                WScript.Echo d.Item(k)
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['42', '99']

    def test_empty_dictionary_items_returns_empty_array(self):
        from pybasil.interpreter import VBScriptDictionary
        d = VBScriptDictionary()
        arr = d.Items()
        assert isinstance(arr, VBScriptArray)
        assert arr.ubound() == -1

    def test_empty_dictionary_keys_returns_empty_array(self):
        from pybasil.interpreter import VBScriptDictionary
        d = VBScriptDictionary()
        arr = d.Keys()
        assert isinstance(arr, VBScriptArray)
        assert arr.ubound() == -1


class TestInterpreterWhileStatement:
    """Test While...Wend statements."""

    def test_while_basic(self):
        output = io.StringIO()
        run(
            """
            x = 0
            While x < 3
                WScript.Echo x
                x = x + 1
            Wend
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0', '1', '2']

    def test_while_false_initially(self):
        output = io.StringIO()
        run(
            """
            While False
                WScript.Echo "never"
            Wend
            WScript.Echo "done"
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'done'

    def test_while_nested(self):
        output = io.StringIO()
        run(
            """
            i = 0
            While i < 2
                j = 0
                While j < 2
                    WScript.Echo i & "," & j
                    j = j + 1
                Wend
                i = i + 1
            Wend
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0,0', '0,1', '1,0', '1,1']

    def test_exit_for_in_while_raises_error(self):
        program = parse("""
            Dim count
            count = 0
            While count < 5
                count = count + 1
                Exit For
            Wend
        """)
        interpreter = Interpreter()
        with pytest.raises(VBScriptError, match='Exit For not valid in While loop'):
            interpreter.interpret(program)


class TestInterpreterDoLoop:
    """Test Do...Loop statements."""

    def test_do_while_pre_test(self):
        output = io.StringIO()
        run(
            """
            x = 0
            Do While x < 3
                WScript.Echo x
                x = x + 1
            Loop
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0', '1', '2']

    def test_do_until_pre_test(self):
        output = io.StringIO()
        run(
            """
            x = 0
            Do Until x >= 3
                WScript.Echo x
                x = x + 1
            Loop
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0', '1', '2']

    def test_do_loop_while_post_test(self):
        output = io.StringIO()
        run(
            """
            x = 0
            Do
                WScript.Echo x
                x = x + 1
            Loop While x < 3
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0', '1', '2']

    def test_do_loop_until_post_test(self):
        output = io.StringIO()
        run(
            """
            x = 0
            Do
                WScript.Echo x
                x = x + 1
            Loop Until x >= 3
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0', '1', '2']

    def test_do_while_post_executes_once(self):
        output = io.StringIO()
        run(
            """
            x = 10
            Do
                WScript.Echo "once"
            Loop While x < 5
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'once'

    def test_do_until_post_executes_once(self):
        output = io.StringIO()
        run(
            """
            x = 10
            Do
                WScript.Echo "once"
            Loop Until x > 5
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'once'

    def test_do_exit_do(self):
        output = io.StringIO()
        run(
            """
            x = 0
            Do While True
                WScript.Echo x
                x = x + 1
                If x >= 3 Then
                    Exit Do
                End If
            Loop
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0', '1', '2']

    def test_do_nested_exit(self):
        output = io.StringIO()
        run(
            """
            x = 0
            Do
                y = 0
                Do
                    WScript.Echo x & "," & y
                    y = y + 1
                    If y >= 2 Then
                        Exit Do
                    End If
                Loop
                x = x + 1
                If x >= 2 Then
                    Exit Do
                End If
            Loop
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0,0', '0,1', '1,0', '1,1']


class TestInterpreterExitStatement:
    """Test Exit statements."""

    def test_exit_for_basic(self):
        output = io.StringIO()
        run(
            """
            For i = 1 To 10
                WScript.Echo i
                Exit For
            Next
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '1'

    def test_exit_do_basic(self):
        output = io.StringIO()
        run(
            """
            Do While True
                WScript.Echo "once"
                Exit Do
            Loop
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'once'

    def test_exit_for_nested(self):
        output = io.StringIO()
        run(
            """
            For i = 1 To 3
                For j = 1 To 3
                    If j = 2 Then
                        Exit For
                    End If
                    WScript.Echo i & "," & j
                Next
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['1,1', '2,1', '3,1']


class TestInterpreterSub:
    """Test Sub procedures."""

    def test_sub_no_params(self):
        output = io.StringIO()
        run(
            """
            Sub SayHello
                WScript.Echo "Hello"
            End Sub
            
            SayHello
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'Hello'

    def test_sub_with_params(self):
        output = io.StringIO()
        run(
            """
            Sub Greet(name)
                WScript.Echo "Hello, " & name
            End Sub
            
            Greet "World"
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'Hello, World'

    def test_sub_call_with_call_keyword(self):
        output = io.StringIO()
        run(
            """
            Sub SayHello
                WScript.Echo "Hello"
            End Sub
            
            Call SayHello
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'Hello'

    def test_sub_exit_sub(self):
        output = io.StringIO()
        run(
            """
            Sub TestExit
                WScript.Echo "Before"
                Exit Sub
                WScript.Echo "After"
            End Sub
            
            TestExit
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['Before']

    def test_sub_local_scope(self):
        output = io.StringIO()
        # Note: Using Call keyword for procedure call without arguments
        # to avoid ambiguity with member access chains
        run(
            """
            x = 10
            
            Sub TestScope
                Dim x
                x = 20
                WScript.Echo "Inside: " & x
            End Sub
            
            Call TestScope
            WScript.Echo "Outside: " & x
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['Inside: 20', 'Outside: 10']

    def test_sub_called_in_expression_returns_empty(self):
        output = io.StringIO()
        run(
            """
            Sub MySub(x)
                WScript.Echo x
            End Sub

            result = MySub(42)
            WScript.Echo TypeName(result)
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['42', 'Empty']


class TestInterpreterFunction:
    """Test Function procedures."""

    def test_function_no_params(self):
        output = io.StringIO()
        run(
            """
            Function GetAnswer
                GetAnswer = 42
            End Function
            
            result = GetAnswer
            WScript.Echo result
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '42'

    def test_function_with_params(self):
        output = io.StringIO()
        run(
            """
            Function Add(a, b)
                Add = a + b
            End Function
            
            result = Add(3, 4)
            WScript.Echo result
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '7'

    def test_function_in_expression(self):
        output = io.StringIO()
        run(
            """
            Function DoubleVal(x)
                DoubleVal = x * 2
            End Function
            
            result = DoubleVal(5) + 1
            WScript.Echo result
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '11'

    def test_function_exit_function(self):
        output = io.StringIO()
        run(
            """
            Function EarlyReturn
                EarlyReturn = 1
                Exit Function
                EarlyReturn = 2
            End Function
            
            WScript.Echo EarlyReturn
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '1'

    def test_function_nested_call(self):
        output = io.StringIO()
        run(
            """
            Function Square(x)
                Square = x * x
            End Function
            
            Function SumOfSquares(a, b)
                SumOfSquares = Square(a) + Square(b)
            End Function
            
            WScript.Echo SumOfSquares(3, 4)
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '25'


class TestInterpreterByRefByVal:
    """Test ByRef and ByVal parameter passing."""

    def test_byref_modifies_original(self):
        output = io.StringIO()
        run(
            """
            Sub Increment(ByRef x)
                x = x + 1
            End Sub
            
            value = 5
            Increment value
            WScript.Echo value
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '6'

    def test_byval_does_not_modify_original(self):
        output = io.StringIO()
        run(
            """
            Sub TryToModify(ByVal x)
                x = x + 1
            End Sub
            
            value = 5
            TryToModify value
            WScript.Echo value
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '5'

    def test_default_is_byref(self):
        output = io.StringIO()
        # Note: Using Call keyword to avoid ambiguity with member access chains
        run(
            """
            Sub DoubleVal(x)
                x = x * 2
            End Sub
            
            value = 10
            Call DoubleVal(value)
            WScript.Echo value
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '20'

    def test_byref_with_expression(self):
        output = io.StringIO()
        run(
            """
            Sub Increment(ByRef x)
                x = x + 1
            End Sub
            
            value = 5
            Increment value + 0
            WScript.Echo value
        """,
            output_stream=output,
        )
        # When passing an expression to ByRef, it should not modify the original
        assert output.getvalue().strip() == '5'

    def test_mixed_byref_byval(self):
        output = io.StringIO()
        run(
            """
            Sub Process(ByRef refVar, ByVal valVar)
                refVar = refVar + 1
                valVar = valVar + 1
                WScript.Echo "Inside: " & refVar & ", " & valVar
            End Sub
            
            a = 10
            b = 20
            Process a, b
            WScript.Echo "Outside: " & a & ", " & b
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['Inside: 11, 21', 'Outside: 11, 20']


class TestInterpreterProcedureScoping:
    """Test procedure scoping rules."""

    def test_access_outer_variable(self):
        output = io.StringIO()
        run(
            """
            x = 10
            
            Sub ShowX
                WScript.Echo x
            End Sub
            
            ShowX
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '10'

    def test_local_shadows_outer(self):
        output = io.StringIO()
        # Note: Using Call keyword to avoid ambiguity
        run(
            """
            x = 10
            
            Sub TestShadow
                Dim x
                x = 20
                WScript.Echo x
            End Sub
            
            Call TestShadow
            WScript.Echo x
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['20', '10']

    def test_modify_outer_without_dim(self):
        output = io.StringIO()
        run(
            """
            x = 10
            
            Sub ModifyX
                x = 20
            End Sub
            
            Call ModifyX
            WScript.Echo x
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '20'

    def test_procedure_recursion(self):
        output = io.StringIO()
        run(
            """
            Function Factorial(n)
                If n <= 1 Then
                    Factorial = 1
                Else
                    Factorial = n * Factorial(n - 1)
                End If
            End Function
            
            WScript.Echo Factorial(5)
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '120'


class TestErrorHandling:
    """Test error handling with On Error statements."""

    def test_on_error_resume_next_continues_after_error(self):
        """On Error Resume Next should continue execution after error."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            x = 1 / 0
            WScript.Echo "After error"
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'After error'

    def test_on_error_goto_0_resets_error_handling(self):
        """On Error GoTo 0 should reset error handling to default."""
        output = io.StringIO()
        interpreter = Interpreter(output_stream=output)
        program = parse(
            """
            On Error Resume Next
            x = 1 / 0
            On Error GoTo 0
            y = 1 / 0
        """
        )
        with pytest.raises(VBScriptError):
            interpreter.interpret(program)

    def test_err_number_after_error(self):
        """Err.Number should be set after an error."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            x = 1 / 0
            WScript.Echo Err.Number
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '11'  # Division by zero error number

    def test_err_description_after_error(self):
        """Err.Description should contain error message."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            x = 1 / 0
            WScript.Echo Err.Description
        """,
            output_stream=output,
        )
        assert 'Division by zero' in output.getvalue()

    def test_err_clear(self):
        """Err.Clear should reset error information."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            x = 1 / 0
            errNum = Err.Number
            WScript.Echo errNum
            Err.Clear
        """,
            output_stream=output,
        )
        # Just check that we got the error number before clear
        assert output.getvalue().strip() == '11'

    def test_err_number_zero_initially(self):
        """Err.Number should be 0 initially."""
        output = io.StringIO()
        run(
            """
            WScript.Echo Err.Number
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '0'

    def test_error_propagates_without_resume_next(self):
        """Errors should propagate without On Error Resume Next."""
        program = parse('x = 1 / 0')
        interpreter = Interpreter()
        with pytest.raises(VBScriptError):
            interpreter.interpret(program)

    def test_on_error_resume_next_in_procedure(self):
        """On Error Resume Next in procedure should be scoped."""
        output = io.StringIO()
        run(
            """
            Sub TestSub
                On Error Resume Next
                x = 1 / 0
                WScript.Echo "In sub after error"
            End Sub
            
            Call TestSub
            WScript.Echo "After sub"
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert 'In sub after error' in lines
        assert 'After sub' in lines

    def test_error_mode_resets_on_procedure_exit(self):
        """Error mode should reset when exiting procedure."""
        output = io.StringIO()
        interpreter = Interpreter(output_stream=output)
        program = parse(
            """
            Sub TestSub
                On Error Resume Next
            End Sub
            
            Call TestSub
            x = 1 / 0
        """
        )
        # Error should propagate because error mode resets after procedure
        with pytest.raises(VBScriptError):
            interpreter.interpret(program)

    def test_err_raise(self):
        """Err.Raise should raise an error."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            Err.Raise 100, "TestSource", "Test Description"
            n = Err.Number
            s = Err.Source
            d = Err.Description
            WScript.Echo n
            WScript.Echo s
            WScript.Echo d
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines[0] == '100'
        assert lines[1] == 'TestSource'
        assert lines[2] == 'Test Description'

    def test_multiple_errors_resume_next(self):
        """Multiple errors should each set Err object."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            x = 1 / 0
            WScript.Echo Err.Number
            y = CInt("not a number")
            WScript.Echo Err.Number
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines[0] == '11'  # Division by zero
        # Second error number may vary

    def test_type_mismatch_error_number(self):
        """Type mismatch should have error number 13."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            x = CInt("abc")
            WScript.Echo Err.Number
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '13'

    def test_err_source_after_error(self):
        """Err.Source should be set after an error."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            x = 1 / 0
            WScript.Echo Err.Source
        """,
            output_stream=output,
        )
        assert 'VBScript' in output.getvalue()

    def test_case_insensitive_on_error(self):
        """On Error statements should be case insensitive."""
        output = io.StringIO()
        run(
            """
            ON ERROR RESUME NEXT
            x = 1 / 0
            wscript.echo "Works"
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'Works'

    def test_case_insensitive_err_object(self):
        """Err object access should be case insensitive."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            x = 1 / 0
            WScript.Echo err.number
            WScript.Echo ERR.NUMBER
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines[0] == '11'
        assert lines[1] == '11'


class TestSelectCase:
    """Test Select Case statement."""

    def test_single_value_match(self):
        output = io.StringIO()
        run(
            """
            Dim x
            x = 2
            Select Case x
                Case 1
                    WScript.Echo "one"
                Case 2
                    WScript.Echo "two"
                Case 3
                    WScript.Echo "three"
            End Select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'two'

    def test_no_match_no_else(self):
        output = io.StringIO()
        run(
            """
            Dim x
            x = 99
            Select Case x
                Case 1
                    WScript.Echo "one"
                Case 2
                    WScript.Echo "two"
            End Select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == ''

    def test_case_else(self):
        output = io.StringIO()
        run(
            """
            Dim x
            x = 42
            Select Case x
                Case 1
                    WScript.Echo "one"
                Case Else
                    WScript.Echo "other"
            End Select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'other'

    def test_comma_separated_values(self):
        output = io.StringIO()
        run(
            """
            Dim x
            x = 5
            Select Case x
                Case 1, 2, 3
                    WScript.Echo "small"
                Case 4, 5, 6
                    WScript.Echo "medium"
                Case Else
                    WScript.Echo "large"
            End Select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'medium'

    def test_relational_checks_with_true(self):
        output = io.StringIO()
        run(
            """
            Dim score
            score = 85
            Select Case True
                Case score >= 90
                    WScript.Echo "A"
                Case score >= 80
                    WScript.Echo "B"
                Case score >= 70
                    WScript.Echo "C"
                Case Else
                    WScript.Echo "F"
            End Select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'B'

    def test_string_matching(self):
        output = io.StringIO()
        run(
            """
            Dim fruit
            fruit = "Banana"
            Select Case fruit
                Case "Apple", "Pear"
                    WScript.Echo "pome"
                Case "Banana", "Mango"
                    WScript.Echo "tropical"
                Case Else
                    WScript.Echo "unknown"
            End Select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'tropical'

    def test_select_case_in_sub(self):
        output = io.StringIO()
        run(
            """
            Sub Classify(n)
                Select Case True
                    Case n < 0
                        WScript.Echo "negative"
                    Case n = 0
                        WScript.Echo "zero"
                    Case n > 0
                        WScript.Echo "positive"
                End Select
            End Sub

            Call Classify(-5)
            Call Classify(0)
            Call Classify(42)
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['negative', 'zero', 'positive']

    def test_select_case_first_match_wins(self):
        output = io.StringIO()
        run(
            """
            Dim x
            x = 5
            Select Case True
                Case x > 0
                    WScript.Echo "positive"
                Case x > 3
                    WScript.Echo "greater than 3"
                Case x > 1
                    WScript.Echo "greater than 1"
            End Select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'positive'

    def test_select_case_with_assignment_in_body(self):
        output = io.StringIO()
        run(
            """
            Dim x, result
            x = 2
            Select Case x
                Case 1
                    result = "a"
                Case 2
                    result = "b"
                Case 3
                    result = "c"
            End Select
            WScript.Echo result
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'b'

    def test_select_case_multiple_statements_in_body(self):
        output = io.StringIO()
        run(
            """
            Dim x
            x = 1
            Select Case x
                Case 1
                    WScript.Echo "line1"
                    WScript.Echo "line2"
                Case 2
                    WScript.Echo "other"
            End Select
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['line1', 'line2']

    def test_select_case_string_comma_list_with_grade(self):
        output = io.StringIO()
        run(
            """
            Dim score, grade
            score = 85
            Select Case score
                Case 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100
                    grade = "A"
                Case 80, 81, 82, 83, 84, 85, 86, 87, 88, 89
                    grade = "B"
                Case 70, 71, 72, 73, 74, 75, 76, 77, 78, 79
                    grade = "C"
                Case Else
                    grade = "F"
            End Select
            WScript.Echo "Score " & score & " = Grade " & grade
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'Score 85 = Grade B'

    def test_select_case_relational_with_else(self):
        output = io.StringIO()
        run(
            """
            Dim score, grade
            score = 95
            Select Case True
                Case score = 100
                    grade = "Perfect"
                Case score >= 90
                    grade = "Excellent"
                Case score >= 80
                    grade = "Good"
                Case Else
                    grade = "Keep trying"
            End Select
            WScript.Echo "Score " & score & " = " & grade
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'Score 95 = Excellent'

    def test_case_insensitive_select_case(self):
        output = io.StringIO()
        run(
            """
            Dim x
            x = 1
            select case x
                case 1
                    wscript.echo "matched"
            end select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'matched'

    def test_nested_select_case(self):
        output = io.StringIO()
        run(
            """
            Dim x, y
            x = 1
            y = 2
            Select Case x
                Case 1
                    Select Case y
                        Case 1
                            WScript.Echo "1-1"
                        Case 2
                            WScript.Echo "1-2"
                    End Select
                Case 2
                    WScript.Echo "2-x"
            End Select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '1-2'


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


class TestSelectCaseRangeAndIs:
    """Tests for Case x To y and Case Is <op> expr in Select Case."""

    def test_case_range_match(self):
        output = io.StringIO()
        run(
            '''Dim x
x = 5
Select Case x
    Case 1 To 3
        WScript.Echo "low"
    Case 4 To 7
        WScript.Echo "mid"
    Case Else
        WScript.Echo "other"
End Select
''',
            output_stream=output,
        )
        assert output.getvalue().strip() == 'mid'

    def test_case_range_no_match(self):
        output = io.StringIO()
        run(
            '''Dim x
x = 10
Select Case x
    Case 1 To 3
        WScript.Echo "low"
    Case 4 To 7
        WScript.Echo "mid"
    Case Else
        WScript.Echo "other"
End Select
''',
            output_stream=output,
        )
        assert output.getvalue().strip() == 'other'

    def test_case_is_greater_than(self):
        output = io.StringIO()
        run(
            '''Dim x
x = 10
Select Case x
    Case Is > 5
        WScript.Echo "big"
    Case Else
        WScript.Echo "small"
End Select
''',
            output_stream=output,
        )
        assert output.getvalue().strip() == 'big'

    def test_case_is_less_than(self):
        output = io.StringIO()
        run(
            '''Dim x
x = 2
Select Case x
    Case Is < 5
        WScript.Echo "small"
    Case Else
        WScript.Echo "big"
End Select
''',
            output_stream=output,
        )
        assert output.getvalue().strip() == 'small'

    def test_case_mixed_range_is_value(self):
        output = io.StringIO()
        run(
            '''Dim x
x = 15
Select Case x
    Case 1 To 5
        WScript.Echo "1-5"
    Case 6, 7, 8
        WScript.Echo "6-8"
    Case Is >= 10
        WScript.Echo ">=10"
    Case Else
        WScript.Echo "other"
End Select
''',
            output_stream=output,
        )
        assert output.getvalue().strip() == '>=10'


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
        interpreter = Interpreter()
        result = interpreter._builtin_array()
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


class TestBuiltinsDictIntegrity:
    """Ensure builtins dictionary has no issues."""

    def test_isnumeric_registered_once(self):
        interpreter = Interpreter()
        assert 'isnumeric' in interpreter._builtins
        assert interpreter._builtins['isnumeric'] is not None
