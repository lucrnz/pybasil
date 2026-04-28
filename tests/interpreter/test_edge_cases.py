"""Tests for the VBScript interpreter."""

import pytest
import io
from pybasil import (
    Interpreter,
    parse,
    run,
    VBScriptError,
    NULL,
)


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
