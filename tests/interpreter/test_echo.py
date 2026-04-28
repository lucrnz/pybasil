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
        assert output.getvalue().strip() == 'True'

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
