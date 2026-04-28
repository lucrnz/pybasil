"""Tests for the VBScript interpreter."""

import io
from pybasil import (
    Interpreter,
    parse,
    run,
)


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
