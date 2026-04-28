"""Tests for the VBScript interpreter."""

import pytest
from pybasil import (
    Interpreter,
    parse,
    VBScriptError,
)


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
