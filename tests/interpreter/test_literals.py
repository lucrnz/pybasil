"""Tests for the VBScript interpreter."""

from pybasil import (
    Interpreter,
    parse,
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

        assert interpreter._environment.get('x') == EMPTY

    def test_null_literal(self):
        program = parse('x = Null')
        interpreter = Interpreter()
        interpreter.interpret(program)

        assert interpreter._environment.get('x') == NULL
