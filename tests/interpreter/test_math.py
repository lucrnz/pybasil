"""Tests for the VBScript interpreter."""

import pytest
from pybasil import (
    Interpreter,
    parse,
    VBScriptError,
)


class TestMathFunctionsSgn:
    """Test Sgn function."""

    def test_sgn_positive(self):
        program = parse('x = Sgn(42)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 1

    def test_sgn_negative(self):
        program = parse('x = Sgn(-7)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == -1

    def test_sgn_zero(self):
        program = parse('x = Sgn(0)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 0

    def test_sgn_float_positive(self):
        program = parse('x = Sgn(0.001)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 1

    def test_sgn_float_negative(self):
        program = parse('x = Sgn(-3.14)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == -1

class TestMathFunctionsLogExp:
    """Test Log and Exp functions."""

    def test_log_one(self):
        program = parse('x = Log(1)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 0.0

    def test_log_e(self):
        program = parse('x = Log(Exp(1))')
        interp = Interpreter()
        interp.interpret(program)
        assert abs(interp._environment.get('x') - 1.0) < 1e-10

    def test_log_positive(self):
        import math
        program = parse('x = Log(10)')
        interp = Interpreter()
        interp.interpret(program)
        assert abs(interp._environment.get('x') - math.log(10)) < 1e-10

    def test_log_zero_error(self):
        program = parse('x = Log(0)')
        interp = Interpreter()
        with pytest.raises(VBScriptError):
            interp.interpret(program)

    def test_log_negative_error(self):
        program = parse('x = Log(-1)')
        interp = Interpreter()
        with pytest.raises(VBScriptError):
            interp.interpret(program)

    def test_exp_zero(self):
        program = parse('x = Exp(0)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 1.0

    def test_exp_one(self):
        import math
        program = parse('x = Exp(1)')
        interp = Interpreter()
        interp.interpret(program)
        assert abs(interp._environment.get('x') - math.e) < 1e-10

    def test_exp_negative(self):
        import math
        program = parse('x = Exp(-1)')
        interp = Interpreter()
        interp.interpret(program)
        assert abs(interp._environment.get('x') - math.exp(-1)) < 1e-10

class TestMathFunctionsTrig:
    """Test Sin, Cos, Tan, and Atn functions."""

    def test_sin_zero(self):
        program = parse('x = Sin(0)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 0.0

    def test_cos_zero(self):
        program = parse('x = Cos(0)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 1.0

    def test_tan_zero(self):
        program = parse('x = Tan(0)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 0.0

    def test_atn_zero(self):
        program = parse('x = Atn(0)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 0.0

    def test_sin_pi_half(self):
        program = parse('x = Sin(Atn(1) * 2)')
        interp = Interpreter()
        interp.interpret(program)
        assert abs(interp._environment.get('x') - 1.0) < 1e-10

    def test_cos_pi(self):
        program = parse('x = Cos(Atn(1) * 4)')
        interp = Interpreter()
        interp.interpret(program)
        assert abs(interp._environment.get('x') - (-1.0)) < 1e-10

    def test_atn_one_is_pi_over_4(self):
        import math
        program = parse('x = Atn(1)')
        interp = Interpreter()
        interp.interpret(program)
        assert abs(interp._environment.get('x') - math.pi / 4) < 1e-10

    def test_tan_pi_over_4(self):
        program = parse('x = Tan(Atn(1))')
        interp = Interpreter()
        interp.interpret(program)
        assert abs(interp._environment.get('x') - 1.0) < 1e-10

    def test_sin_cos_identity(self):
        """Sin^2 + Cos^2 = 1"""
        program = parse('''
        Dim angle
        angle = 1.23
        x = Sin(angle) * Sin(angle) + Cos(angle) * Cos(angle)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert abs(interp._environment.get('x') - 1.0) < 1e-10

    def test_pi_via_atn(self):
        """Classic VBScript idiom: pi = 4 * Atn(1)"""
        import math
        program = parse('pi = 4 * Atn(1)')
        interp = Interpreter()
        interp.interpret(program)
        assert abs(interp._environment.get('pi') - math.pi) < 1e-10
