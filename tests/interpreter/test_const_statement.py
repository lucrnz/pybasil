"""Tests for the VBScript interpreter."""

import pytest
import io
from pybasil import (
    Interpreter,
    parse,
    VBScriptError,
)


class TestConstStatement:
    """Test Const declarations."""

    def test_const_integer(self):
        program = parse('Const MAX_SIZE = 100')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('MAX_SIZE') == 100

    def test_const_string(self):
        program = parse('Const GREETING = "Hello"')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('GREETING') == 'Hello'

    def test_const_float(self):
        program = parse('Const PI = 3.14159')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('PI') == 3.14159

    def test_const_boolean(self):
        program = parse('Const FLAG = True')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('FLAG') is True

    def test_const_expression(self):
        program = parse('Const TOTAL = 10 + 20')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('TOTAL') == 30

    def test_const_multiple(self):
        program = parse('Const A = 1, B = 2, C = 3')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('A') == 1
        assert interp._environment.get('B') == 2
        assert interp._environment.get('C') == 3

    def test_const_used_in_expression(self):
        program = parse('''
        Const X = 10
        result = X * 2
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 20

    def test_const_reassignment_raises_error(self):
        program = parse('''
        Const X = 10
        X = 20
        ''')
        interp = Interpreter()
        with pytest.raises(VBScriptError, match='constant'):
            interp.interpret(program)

    def test_const_case_insensitive(self):
        program = parse('''
        Const myConst = 42
        result = MYCONST
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 42

    def test_const_in_loop(self):
        output = io.StringIO()
        program = parse('''
        Const LIMIT = 5
        Dim total
        total = 0
        Dim i
        For i = 1 To LIMIT
            total = total + i
        Next
        WScript.Echo total
        ''')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        assert output.getvalue().strip() == '15'

    def test_const_negative_value(self):
        program = parse('Const NEG = -42')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('NEG') == -42

    def test_const_hex_value(self):
        program = parse('Const HEX_VAL = &HFF')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('HEX_VAL') == 255
