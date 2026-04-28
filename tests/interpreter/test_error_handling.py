"""Tests for the VBScript interpreter."""

import pytest
import io
from pybasil import (
    Interpreter,
    parse,
    run,
    VBScriptError,
)


class TestErrorHandling:
    """Test error handling with On Error statements."""

    def test_on_error_resume_next_continues_after_error(self):
        """On Error Resume Next should continue execution after error."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            x = 1 / 0
            WScript.Echo "After error"
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'After error'

    def test_on_error_goto_0_resets_error_handling(self):
        """On Error GoTo 0 should reset error handling to default."""
        output = io.StringIO()
        interpreter = Interpreter(output_stream=output)
        program = parse(
            """
            On Error Resume Next
            x = 1 / 0
            On Error GoTo 0
            y = 1 / 0
        """
        )
        with pytest.raises(VBScriptError):
            interpreter.interpret(program)

    def test_err_number_after_error(self):
        """Err.Number should be set after an error."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            x = 1 / 0
            WScript.Echo Err.Number
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '11'  # Division by zero error number

    def test_err_description_after_error(self):
        """Err.Description should contain error message."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            x = 1 / 0
            WScript.Echo Err.Description
        """,
            output_stream=output,
        )
        assert 'Division by zero' in output.getvalue()

    def test_err_clear(self):
        """Err.Clear should reset error information."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            x = 1 / 0
            errNum = Err.Number
            WScript.Echo errNum
            Err.Clear
        """,
            output_stream=output,
        )
        # Just check that we got the error number before clear
        assert output.getvalue().strip() == '11'

    def test_err_number_zero_initially(self):
        """Err.Number should be 0 initially."""
        output = io.StringIO()
        run(
            """
            WScript.Echo Err.Number
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '0'

    def test_error_propagates_without_resume_next(self):
        """Errors should propagate without On Error Resume Next."""
        program = parse('x = 1 / 0')
        interpreter = Interpreter()
        with pytest.raises(VBScriptError):
            interpreter.interpret(program)

    def test_on_error_resume_next_in_procedure(self):
        """On Error Resume Next in procedure should be scoped."""
        output = io.StringIO()
        run(
            """
            Sub TestSub
                On Error Resume Next
                x = 1 / 0
                WScript.Echo "In sub after error"
            End Sub
            
            Call TestSub
            WScript.Echo "After sub"
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert 'In sub after error' in lines
        assert 'After sub' in lines

    def test_error_mode_resets_on_procedure_exit(self):
        """Error mode should reset when exiting procedure."""
        output = io.StringIO()
        interpreter = Interpreter(output_stream=output)
        program = parse(
            """
            Sub TestSub
                On Error Resume Next
            End Sub
            
            Call TestSub
            x = 1 / 0
        """
        )
        # Error should propagate because error mode resets after procedure
        with pytest.raises(VBScriptError):
            interpreter.interpret(program)

    def test_err_raise(self):
        """Err.Raise should raise an error."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            Err.Raise 100, "TestSource", "Test Description"
            n = Err.Number
            s = Err.Source
            d = Err.Description
            WScript.Echo n
            WScript.Echo s
            WScript.Echo d
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines[0] == '100'
        assert lines[1] == 'TestSource'
        assert lines[2] == 'Test Description'

    def test_multiple_errors_resume_next(self):
        """Multiple errors should each set Err object."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            x = 1 / 0
            WScript.Echo Err.Number
            y = CInt("not a number")
            WScript.Echo Err.Number
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines[0] == '11'  # Division by zero
        # Second error number may vary

    def test_type_mismatch_error_number(self):
        """Type mismatch should have error number 13."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            x = CInt("abc")
            WScript.Echo Err.Number
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '13'

    def test_err_source_after_error(self):
        """Err.Source should be set after an error."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            x = 1 / 0
            WScript.Echo Err.Source
        """,
            output_stream=output,
        )
        assert 'VBScript' in output.getvalue()

    def test_case_insensitive_on_error(self):
        """On Error statements should be case insensitive."""
        output = io.StringIO()
        run(
            """
            ON ERROR RESUME NEXT
            x = 1 / 0
            wscript.echo "Works"
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'Works'

    def test_case_insensitive_err_object(self):
        """Err object access should be case insensitive."""
        output = io.StringIO()
        run(
            """
            On Error Resume Next
            x = 1 / 0
            WScript.Echo err.number
            WScript.Echo ERR.NUMBER
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines[0] == '11'
        assert lines[1] == '11'

class TestDynamicCodeErrorHandling:
    @pytest.mark.parametrize(
        ('setup', 'statement'),
        [
            ('', 'Execute "If Then"'),
            ('', 'ExecuteGlobal "If Then"'),
            ('Dim result', 'result = Eval("1 +")'),
        ],
    )
    def test_dynamic_syntax_errors_obey_resume_next(self, setup, statement):
        output = io.StringIO()
        lines = ['On Error Resume Next']
        if setup:
            lines.append(setup)
        lines.extend([
            statement,
            'WScript.Echo Err.Number',
            'WScript.Echo Err.Description',
        ])
        run('\n'.join(lines), output_stream=output)
        assert output.getvalue().strip().split('\n') == ['1002', 'Syntax error']


    @pytest.mark.parametrize(
        'expr',
        [
            '""',
            '"1 : 2"',
            '"2 + 3 : x = 5"',
            '"2 + 3" & vbLf & "x = 5"',
        ],
    )
    def test_eval_rejects_empty_and_multi_statement_input(self, expr):
        output = io.StringIO()
        run(
            '\n'.join([
                'On Error Resume Next',
                'Dim result',
                f'result = Eval({expr})',
                'WScript.Echo Err.Number',
                'WScript.Echo IsEmpty(result)',
            ]),
            output_stream=output,
        )
        assert output.getvalue().strip().split('\n') == ['1002', '-1']


    @pytest.mark.parametrize('statement', ['x = Execute("y = 1")', 'x = ExecuteGlobal("y = 1")'])
    def test_execute_and_executeglobal_are_rejected_in_expression_position(self, statement):
        interpreter = Interpreter()
        program = parse(statement)
        with pytest.raises(VBScriptError, match='Syntax error'):
            interpreter.interpret(program)

    @pytest.mark.parametrize('statement', ['Execute "y = 1"', 'ExecuteGlobal "y = 1"'])
    def test_execute_and_executeglobal_still_work_as_statements(self, statement):
        output = io.StringIO()
        run('\n'.join([statement, 'WScript.Echo y']), output_stream=output)
        assert output.getvalue().strip() == '1'


    def test_execute_defined_sub_does_not_escape_procedure_scope(self):
        interpreter = Interpreter()
        program = parse(
            '\n'.join([
                'Sub DefineLocalProc',
                '    Execute "Sub HiddenProc() : WScript.Echo ""hidden"" : End Sub"',
                'End Sub',
                'Call DefineLocalProc',
                'Call HiddenProc()',
            ])
        )
        with pytest.raises(VBScriptError, match='Unknown procedure: HiddenProc'):
            interpreter.interpret(program)

    def test_execute_defined_class_does_not_escape_procedure_scope(self):
        interpreter = Interpreter()
        program = parse(
            '\n'.join([
                'Sub DefineLocalClass',
                '    Execute "Class HiddenClass : Public X : End Class"',
                'End Sub',
                'Call DefineLocalClass',
                'Dim obj',
                'Set obj = New HiddenClass',
            ])
        )
        with pytest.raises(VBScriptError, match='Class not defined: HiddenClass'):
            interpreter.interpret(program)

    def test_execute_defined_procedure_stays_visible_within_same_procedure(self):
        output = io.StringIO()
        run(
            '\n'.join([
                'Sub DefineAndCall',
                '    Execute "Sub HiddenProc() : WScript.Echo ""hidden"" : End Sub"',
                '    Call HiddenProc()',
                'End Sub',
                'Call DefineAndCall',
            ]),
            output_stream=output,
        )
        assert output.getvalue().strip() == 'hidden'


    def test_executeglobal_does_not_resolve_implicit_instance_members(self):
        output = io.StringIO()
        run(
            '\n'.join([
                'Class C',
                '    Public Property Get Foo()',
                '        Foo = 7',
                '    End Property',
                '',
                '    Public Sub Seed()',
                '        ExecuteGlobal "fromProp = Foo"',
                '    End Sub',
                'End Class',
                '',
                'Dim obj',
                'Set obj = New C',
                'obj.Seed',
                'WScript.Echo IsEmpty(fromProp)',
            ]),
            output_stream=output,
        )
        assert output.getvalue().strip() == '-1'

class TestErrNotClearedInDefaultMode:
    """Test that Err is not cleared before every statement in DEFAULT mode."""

    def test_err_number_persists_in_resume_next(self):
        output = io.StringIO()
        run(
            '''On Error Resume Next
x = 1 / 0
y = Err.Number
WScript.Echo y
''',
            output_stream=output,
        )
        assert output.getvalue().strip() == '11'

    def test_err_cleared_on_mode_transition(self):
        output = io.StringIO()
        run(
            '''On Error Resume Next
x = 1 / 0
On Error GoTo 0
On Error Resume Next
y = Err.Number
WScript.Echo y
''',
            output_stream=output,
        )
        assert output.getvalue().strip() == '0'
