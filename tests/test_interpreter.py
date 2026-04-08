"""Tests for the VBScript interpreter."""

import pytest
import io
from pybasil import (
    Interpreter,
    parse,
    run,
    VBScriptError,
    EMPTY,
    NULL,
    NOTHING,
)


class TestInterpreterLiterals:
    """Test evaluation of literal values."""

    def test_integer_literal(self):
        program = parse("x = 42")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 42

    def test_float_literal(self):
        program = parse("x = 3.14")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 3.14

    def test_string_literal(self):
        program = parse('x = "Hello"')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == "Hello"

    def test_boolean_true(self):
        program = parse("x = True")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_boolean_false(self):
        program = parse("x = False")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is False

    def test_nothing_literal(self):
        program = parse("x = Nothing")
        interpreter = Interpreter()
        interpreter.interpret(program)
        from pybasil import NOTHING

        assert interpreter._environment.get("x") == NOTHING

    def test_empty_literal(self):
        program = parse("x = Empty")
        interpreter = Interpreter()
        interpreter.interpret(program)
        from pybasil import EMPTY

        assert interpreter._environment.get("x") == EMPTY

    def test_null_literal(self):
        program = parse("x = Null")
        interpreter = Interpreter()
        interpreter.interpret(program)
        from pybasil import NULL

        assert interpreter._environment.get("x") == NULL


class TestInterpreterVariables:
    """Test variable handling."""

    def test_variable_assignment(self):
        program = parse("x = 42")
        interpreter = Interpreter()
        interpreter.interpret(program)
        # Check that x is accessible
        assert interpreter._environment.get("x") == 42

    def test_variable_lookup(self):
        program = parse("""
            x = 10
            y = x
        """)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("y") == 10

    def test_implicit_variable_creation(self):
        program = parse("y = x")
        interpreter = Interpreter()
        interpreter.interpret(program)
        # x should be Empty (implicit creation)
        assert interpreter._environment.get("x") == EMPTY

    def test_dim_statement(self):
        program = parse("Dim x, y, z")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.exists("x")
        assert interpreter._environment.exists("y")
        assert interpreter._environment.exists("z")

    def test_set_statement(self):
        program = parse('Set obj = CreateObject("Test.Object")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.exists("obj")

    def test_case_insensitive_variables(self):
        program = parse("""
            x = 42
            Y = X
        """)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("y") == 42


class TestInterpreterArithmetic:
    """Test arithmetic operations."""

    def test_addition(self):
        program = parse("x = 5 + 3")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 8

    def test_subtraction(self):
        program = parse("x = 10 - 4")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 6

    def test_multiplication(self):
        program = parse("x = 6 * 7")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 42

    def test_division(self):
        program = parse("x = 15 / 3")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 5.0

    def test_integer_division(self):
        program = parse("x = 17 \\ 5")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 3

    def test_modulo(self):
        program = parse("x = 17 Mod 5")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 2

    def test_exponentiation(self):
        program = parse("x = 2 ^ 10")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 1024

    def test_negation(self):
        program = parse("x = -5")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == -5

    def test_unary_plus(self):
        program = parse("x = +5")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 5

    def test_complex_expression(self):
        program = parse("x = 2 + 3 * 4 - 1")
        interpreter = Interpreter()
        interpreter.interpret(program)
        # 2 + 12 - 1 = 13
        assert interpreter._environment.get("x") == 13

    def test_parentheses(self):
        program = parse("x = (2 + 3) * 4")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 20

    def test_division_by_zero(self):
        program = parse("x = 10 / 0")
        interpreter = Interpreter()
        with pytest.raises(VBScriptError):
            interpreter.interpret(program)


class TestInterpreterStringOperations:
    """Test string operations."""

    def test_concatenation(self):
        program = parse('x = "Hello" & " " & "World"')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == "Hello World"

    def test_string_number_concatenation(self):
        program = parse('x = "Value: " & 42')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == "Value: 42"

    def test_number_string_addition(self):
        # In VBScript, + with strings can do concatenation or addition
        program = parse('x = "5" + 3')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == "53"  # String concatenation


class TestInterpreterComparison:
    """Test comparison operations."""

    def test_equals_true(self):
        program = parse("x = (5 = 5)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_equals_false(self):
        program = parse("x = (5 = 6)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is False

    def test_not_equals(self):
        program = parse("x = (5 <> 6)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_less_than(self):
        program = parse("x = (3 < 5)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_greater_than(self):
        program = parse("x = (7 > 5)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_less_equal(self):
        program = parse("x = (5 <= 5)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_greater_equal(self):
        program = parse("x = (5 >= 5)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_string_comparison(self):
        program = parse('x = ("abc" < "def")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_is_operator_nothing(self):
        program = parse("x = (Nothing Is Nothing)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True


class TestInterpreterLogical:
    """Test logical operations."""

    def test_and_true(self):
        program = parse("x = True And True")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_and_false(self):
        program = parse("x = True And False")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is False

    def test_or_true(self):
        program = parse("x = False Or True")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_or_false(self):
        program = parse("x = False Or False")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is False

    def test_not_true(self):
        program = parse("x = Not True")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is False

    def test_not_false(self):
        program = parse("x = Not False")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_xor_true_true(self):
        program = parse("x = True Xor True")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is False

    def test_xor_true_false(self):
        program = parse("x = True Xor False")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_eqv_true_true(self):
        program = parse("x = True Eqv True")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_eqv_true_false(self):
        program = parse("x = True Eqv False")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is False

    def test_imp_true_true(self):
        program = parse("x = True Imp True")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_imp_true_false(self):
        program = parse("x = True Imp False")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is False

    def test_imp_false_true(self):
        program = parse("x = False Imp True")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_imp_false_false(self):
        program = parse("x = False Imp False")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True


class TestInterpreterWScriptEcho:
    """Test WScript.Echo functionality."""

    def test_echo_string(self):
        output = io.StringIO()
        run('WScript.Echo "Hello, World!"', output_stream=output)
        assert output.getvalue().strip() == "Hello, World!"

    def test_echo_number(self):
        output = io.StringIO()
        run("WScript.Echo 42", output_stream=output)
        assert output.getvalue().strip() == "42"

    def test_echo_multiple_args(self):
        output = io.StringIO()
        run('WScript.Echo "Hello", "World", 42', output_stream=output)
        assert output.getvalue().strip() == "Hello World 42"

    def test_echo_boolean(self):
        output = io.StringIO()
        run("WScript.Echo True", output_stream=output)
        assert output.getvalue().strip() == "True"

    def test_echo_variable(self):
        output = io.StringIO()
        run(
            """
            x = "Test Value"
            WScript.Echo x
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == "Test Value"

    def test_echo_expression(self):
        output = io.StringIO()
        run("WScript.Echo 2 + 2", output_stream=output)
        assert output.getvalue().strip() == "4"


class TestInterpreterBuiltins:
    """Test built-in functions."""

    def test_len(self):
        program = parse('x = Len("Hello")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 5

    def test_left(self):
        program = parse('x = Left("Hello", 3)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == "Hel"

    def test_right(self):
        program = parse('x = Right("Hello", 3)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == "llo"

    def test_mid(self):
        program = parse('x = Mid("Hello", 2, 3)')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == "ell"

    def test_trim(self):
        program = parse('x = Trim("  Hello  ")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == "Hello"

    def test_ucase(self):
        program = parse('x = UCase("hello")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == "HELLO"

    def test_lcase(self):
        program = parse('x = LCase("HELLO")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == "hello"

    def test_cstr(self):
        program = parse("x = CStr(42)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == "42"

    def test_cint(self):
        program = parse("x = CInt(3.7)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 3

    def test_cdbl(self):
        program = parse('x = CDbl("3.14")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 3.14

    def test_cbool(self):
        program = parse("x = CBool(1)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_abs(self):
        program = parse("x = Abs(-5)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 5

    def test_sqr(self):
        program = parse("x = Sqr(16)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 4.0

    def test_int(self):
        program = parse("x = Int(3.7)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 3

    def test_fix(self):
        program = parse("x = Fix(3.7)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 3

    def test_round(self):
        program = parse("x = Round(3.14159, 2)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 3.14

    def test_isnumeric_true(self):
        program = parse('x = IsNumeric("123")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_isnumeric_false(self):
        program = parse('x = IsNumeric("abc")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is False

    def test_isempty(self):
        program = parse("x = IsEmpty(Empty)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_isnull(self):
        program = parse("x = IsNull(Null)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") is True

    def test_typename_string(self):
        program = parse('x = TypeName("Hello")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == "String"

    def test_typename_integer(self):
        program = parse("x = TypeName(42)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == "Integer"


class TestInterpreterEdgeCases:
    """Test edge cases and special behaviors."""

    def test_empty_in_arithmetic(self):
        program = parse("x = Empty + 5")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == 5

    def test_empty_in_string_concat(self):
        program = parse('x = Empty & "text"')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == "text"

    def test_null_propagation(self):
        program = parse("x = Null + 5")
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("x") == NULL

    def test_null_comparison(self):
        program = parse("x = (Null = Null)")
        interpreter = Interpreter()
        interpreter.interpret(program)
        # Null comparisons return False in VBScript
        assert interpreter._environment.get("x") is False

    def test_multiple_statements(self):
        program = parse("""
            Dim a, b, c
            a = 1
            b = 2
            c = a + b
        """)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get("c") == 3

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
        assert output.getvalue().strip() == "5"


class TestInterpreterTypeCoercion:
    """Test VBScript type coercion behavior."""

    def test_string_to_number_addition(self):
        program = parse('x = "10" + 5')
        interpreter = Interpreter()
        interpreter.interpret(program)
        # String + number = string concatenation
        assert interpreter._environment.get("x") == "105"

    def test_number_to_string_comparison(self):
        program = parse('x = (5 = "5")')
        interpreter = Interpreter()
        interpreter.interpret(program)
        # String comparison
        assert interpreter._environment.get("x") is True

    def test_boolean_in_arithmetic(self):
        program = parse("x = True + 1")
        interpreter = Interpreter()
        interpreter.interpret(program)
        # True is -1 in VBScript
        assert interpreter._environment.get("x") == 0


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
        assert output.getvalue().strip() == "yes"

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
        assert output.getvalue().strip() == "done"

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
        assert output.getvalue().strip() == "then"

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
        assert output.getvalue().strip() == "else"

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
        assert output.getvalue().strip() == "one"

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
        assert output.getvalue().strip() == "two"

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
        assert output.getvalue().strip() == "other"

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
        assert output.getvalue().strip() == "three"

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
        assert output.getvalue().strip() == "small"

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
        assert output.getvalue().strip() == "yes"


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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["1", "2", "3"]

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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["0", "2", "4", "6", "8", "10"]

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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["5", "4", "3", "2", "1"]

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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["1", "2"]

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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["1,1", "1,2", "2,1", "2,2"]

    def test_for_variable_after_loop(self):
        program = parse("""
            For i = 1 To 5
            Next
        """)
        interpreter = Interpreter()
        interpreter.interpret(program)
        # After loop, i should be 6 (last value + step)
        assert interpreter._environment.get("i") == 6

    def test_for_empty_body(self):
        program = parse("""
            For i = 1 To 5
            Next
        """)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.exists("i")


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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["0", "1", "2"]

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
        assert output.getvalue().strip() == "done"

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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["0,0", "0,1", "1,0", "1,1"]


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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["0", "1", "2"]

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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["0", "1", "2"]

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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["0", "1", "2"]

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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["0", "1", "2"]

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
        assert output.getvalue().strip() == "once"

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
        assert output.getvalue().strip() == "once"

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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["0", "1", "2"]

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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["0,0", "0,1", "1,0", "1,1"]


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
        assert output.getvalue().strip() == "1"

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
        assert output.getvalue().strip() == "once"

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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["1,1", "2,1", "3,1"]


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
        assert output.getvalue().strip() == "Hello"

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
        assert output.getvalue().strip() == "Hello, World"

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
        assert output.getvalue().strip() == "Hello"

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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["Before"]

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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["Inside: 20", "Outside: 10"]


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
        assert output.getvalue().strip() == "42"

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
        assert output.getvalue().strip() == "7"

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
        assert output.getvalue().strip() == "11"

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
        assert output.getvalue().strip() == "1"

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
        assert output.getvalue().strip() == "25"


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
        assert output.getvalue().strip() == "6"

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
        assert output.getvalue().strip() == "5"

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
        assert output.getvalue().strip() == "20"

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
        assert output.getvalue().strip() == "5"

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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["Inside: 11, 21", "Outside: 11, 20"]


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
        assert output.getvalue().strip() == "10"

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
        lines = output.getvalue().strip().split("\n")
        assert lines == ["20", "10"]

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
        assert output.getvalue().strip() == "20"

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
        assert output.getvalue().strip() == "120"
