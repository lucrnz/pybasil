"""Tests for the VBScript interpreter."""

from pybasil import (
    Interpreter,
    parse,
)


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
