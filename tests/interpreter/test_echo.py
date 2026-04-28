"""Tests for the VBScript interpreter."""

import io
from pybasil import (
    run,
)


class TestInterpreterWScriptEcho:
    """Test WScript.Echo functionality."""

    def test_echo_string(self):
        output = io.StringIO()
        run('WScript.Echo "Hello, World!"', output_stream=output)
        assert output.getvalue().strip() == 'Hello, World!'

    def test_echo_number(self):
        output = io.StringIO()
        run('WScript.Echo 42', output_stream=output)
        assert output.getvalue().strip() == '42'

    def test_echo_multiple_args(self):
        output = io.StringIO()
        run('WScript.Echo "Hello", "World", 42', output_stream=output)
        assert output.getvalue().strip() == 'Hello World 42'

    def test_echo_boolean(self):
        output = io.StringIO()
        run('WScript.Echo True', output_stream=output)
        assert output.getvalue().strip() == '-1'

    def test_echo_variable(self):
        output = io.StringIO()
        run(
            """
            x = "Test Value"
            WScript.Echo x
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'Test Value'

    def test_echo_expression(self):
        output = io.StringIO()
        run('WScript.Echo 2 + 2', output_stream=output)
        assert output.getvalue().strip() == '4'

    def test_echo_unary_minus(self):
        output = io.StringIO()
        run('WScript.Echo -1', output_stream=output)
        assert output.getvalue().strip() == '-1'

    def test_echo_unary_plus(self):
        output = io.StringIO()
        run('WScript.Echo +5', output_stream=output)
        assert output.getvalue().strip() == '5'

    def test_echo_boolean_false(self):
        output = io.StringIO()
        run('WScript.Echo False', output_stream=output)
        assert output.getvalue().strip() == '0'

    def test_echo_comparison_result(self):
        output = io.StringIO()
        run('WScript.Echo (5 = 5)', output_stream=output)
        assert output.getvalue().strip() == '-1'

    def test_echo_cbool_display(self):
        output = io.StringIO()
        run('WScript.Echo CBool(1)', output_stream=output)
        assert output.getvalue().strip() == '-1'

    def test_echo_isnull_display(self):
        output = io.StringIO()
        run('WScript.Echo IsNull(Null)', output_stream=output)
        assert output.getvalue().strip() == '-1'

    def test_echo_minus_five_plus_three(self):
        output = io.StringIO()
        run('WScript.Echo -5+3', output_stream=output)
        assert output.getvalue().strip() == '-2'

    def test_echo_paren_neg_pow(self):
        output = io.StringIO()
        run('WScript.Echo (-2) ^ 3', output_stream=output)
        assert output.getvalue().strip() == '-8'

    def test_echo_float_precision(self):
        output = io.StringIO()
        run('WScript.Echo 10 / 3', output_stream=output)
        assert output.getvalue().strip() == '3.33333333333333'

    def test_echo_float_precision_sum(self):
        output = io.StringIO()
        run('WScript.Echo 0.1 + 0.2', output_stream=output)
        assert output.getvalue().strip() == '0.3'

    def test_echo_default_member(self):
        output = io.StringIO()
        run('''
            Class MyObj
                Public Default Property Get Value()
                    Value = "hello"
                End Property
            End Class
            Dim o
            Set o = New MyObj
            WScript.Echo o
        ''', output_stream=output)
        assert output.getvalue().strip() == 'hello'

    def test_echo_null_concat(self):
        output = io.StringIO()
        run('WScript.Echo Null & "hello"', output_stream=output)
        assert output.getvalue().strip() == 'hello'
