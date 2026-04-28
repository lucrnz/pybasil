"""Tests for the VBScript interpreter."""

import pytest
from pybasil import (
    Interpreter,
    parse,
    VBScriptError,
)


class TestOptionExplicit:
    """Test Option Explicit."""

    def test_option_explicit_allows_dim(self):
        program = parse('''
        Option Explicit
        Dim x
        x = 42
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 42

    def test_option_explicit_blocks_undeclared(self):
        program = parse('''
        Option Explicit
        x = 42
        ''')
        interp = Interpreter()
        with pytest.raises(VBScriptError, match='undefined'):
            interp.interpret(program)

    def test_option_explicit_blocks_set_undeclared(self):
        program = parse('''
        Option Explicit
        Set obj = CreateObject("Scripting.Dictionary")
        ''')
        interp = Interpreter()
        with pytest.raises(VBScriptError, match='undefined'):
            interp.interpret(program)

    def test_option_explicit_allows_const(self):
        program = parse('''
        Option Explicit
        Const PI = 3.14
        result = PI * 2
        ''')
        interp = Interpreter()
        # result is not declared but PI is; result assignment should fail
        with pytest.raises(VBScriptError, match='undefined'):
            interp.interpret(program)

    def test_option_explicit_const_declared(self):
        program = parse('''
        Option Explicit
        Const PI = 3.14
        Dim result
        result = PI * 2
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 3.14 * 2

    def test_without_option_explicit_allows_undeclared(self):
        program = parse('''
        x = 42
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 42

    def test_option_explicit_allows_for_loop_var(self):
        program = parse('''
        Option Explicit
        Dim i, total
        total = 0
        For i = 1 To 5
            total = total + i
        Next
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('total') == 15

    def test_option_explicit_allows_procedure_params(self):
        program = parse('''
        Option Explicit
        Dim result
        Function Add(a, b)
            Add = a + b
        End Function
        result = Add(3, 4)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 7

    def test_option_explicit_multiple_dim(self):
        program = parse('''
        Option Explicit
        Dim a, b, c
        a = 1
        b = 2
        c = a + b
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('c') == 3

    def test_option_explicit_set_with_dim(self):
        program = parse('''
        Option Explicit
        Dim d
        Set d = CreateObject("Scripting.Dictionary")
        d.Add "key", 42
        ''')
        interp = Interpreter()
        interp.interpret(program)

    def test_option_explicit_case_insensitive(self):
        program = parse('''
        Option Explicit
        Dim myVar
        MYVAR = 10
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('myVar') == 10
