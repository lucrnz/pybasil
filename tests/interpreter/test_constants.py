"""Tests for the VBScript interpreter."""

import io
from pybasil import (
    Interpreter,
    parse,
    run,
)


class TestInterpreterConstants:
    """Test VBScript built-in constants."""

    # -- String constants --

    def test_vbcrlf(self):
        program = parse('x = vbCrLf')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == '\r\n'

    def test_vbcr(self):
        program = parse('x = vbCr')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == '\r'

    def test_vblf(self):
        program = parse('x = vbLf')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == '\n'

    def test_vbtab(self):
        program = parse('x = vbTab')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == '\t'

    def test_vbnewline(self):
        program = parse('x = vbNewLine')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == '\r\n'

    def test_vbnullchar(self):
        program = parse('x = vbNullChar')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == '\x00'

    def test_vbnullstring(self):
        program = parse('x = vbNullString')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == ''

    def test_vbback(self):
        program = parse('x = vbBack')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == '\x08'

    def test_vbformfeed(self):
        program = parse('x = vbFormFeed')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == '\x0c'

    def test_vbverticaltab(self):
        program = parse('x = vbVerticalTab')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == '\x0b'

    # -- String constants used in concatenation --

    def test_vbcrlf_concatenation(self):
        output = io.StringIO()
        run('WScript.Echo "line1" & vbCrLf & "line2"', output_stream=output)
        assert output.getvalue() == 'line1\r\nline2\n'

    def test_vbtab_concatenation(self):
        output = io.StringIO()
        run('WScript.Echo "col1" & vbTab & "col2"', output_stream=output)
        assert output.getvalue() == 'col1\tcol2\n'

    # -- VarType constants --

    def test_vbempty_constant(self):
        program = parse('x = vbEmpty')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 0

    def test_vbnull_constant(self):
        program = parse('x = vbNull')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 1

    def test_vbinteger_constant(self):
        program = parse('x = vbInteger')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 2

    def test_vblong_constant(self):
        program = parse('x = vbLong')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 3

    def test_vbdouble_constant(self):
        program = parse('x = vbDouble')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 5

    def test_vbstring_constant(self):
        program = parse('x = vbString')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 8

    def test_vbboolean_constant(self):
        program = parse('x = vbBoolean')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 11

    def test_vbarray_constant(self):
        program = parse('x = vbArray')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 8192

    # -- MsgBox constants --

    def test_vbok_constant(self):
        program = parse('x = vbOK')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 1

    def test_vbyesno_constant(self):
        program = parse('x = vbYesNo')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 4

    def test_vbyes_constant(self):
        program = parse('x = vbYes')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 6

    def test_vbno_constant(self):
        program = parse('x = vbNo')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 7

    def test_vbcritical_constant(self):
        program = parse('x = vbCritical')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 16

    def test_vbinformation_constant(self):
        program = parse('x = vbInformation')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 64

    # -- Comparison constants --

    def test_vbbinarycompare_constant(self):
        program = parse('x = vbBinaryCompare')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 0

    def test_vbtextcompare_constant(self):
        program = parse('x = vbTextCompare')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 1

    # -- Tristate constants --

    def test_vbtrue_constant(self):
        program = parse('x = vbTrue')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == -1

    def test_vbfalse_constant(self):
        program = parse('x = vbFalse')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 0

    # -- Color constants --

    def test_vbred_constant(self):
        program = parse('x = vbRed')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 0x0000FF

    def test_vbblue_constant(self):
        program = parse('x = vbBlue')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 0xFF0000

    # -- Case insensitivity --

    def test_constants_are_case_insensitive(self):
        program = parse('x = VBCRLF')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == '\r\n'

    # -- Constants used in expressions --

    def test_vartype_constant_in_comparison(self):
        output = io.StringIO()
        run('''
            If VarType(42) = vbInteger Then
                WScript.Echo "integer"
            End If
        ''', output_stream=output)
        assert output.getvalue().strip() == 'integer'

    def test_msgbox_constants_as_bitmask(self):
        program = parse('x = vbYesNo + vbQuestion + vbDefaultButton2')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 4 + 32 + 256

    def test_vbobjecterror_constant(self):
        program = parse('x = vbObjectError')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == -2147221504
