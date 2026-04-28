"""Tests for the VBScript interpreter."""

from pybasil import (
    Interpreter,
    parse,
    EMPTY,
)


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
