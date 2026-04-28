"""Regression tests for the interpreter."""

import io
from pathlib import Path

from pybasil import Interpreter, parse


class RegressionTestsUtils:
    @staticmethod
    def load_script(name: str) -> str:
        if '.vbs' not in name:
            raise ValueError('Invalid filename, must be a VBScript')
        script_path = Path(__file__).parent / 'regression_scripts' / name
        return script_path.read_text()


class TestRegressionScripts:
    def test_regression_script_alpha_passes(self):
        script = RegressionTestsUtils.load_script('alpha.vbs')
        program = parse(script)
        interpreter = Interpreter()
        interpreter.interpret(program)
        # Check expected values
        assert interpreter._environment.get('totalCount') == 151
        assert interpreter._environment.get('passCount') == 151
        assert interpreter._environment.get('failCount') == 0

    def test_regression_script_beta_greedy_arglist_bugfix(self):
        """Implicit calls without parens must not consume the next line as args."""
        script = RegressionTestsUtils.load_script('beta-greedy-arglist-bugfix.vbs')
        program = parse(script)
        output = io.StringIO()
        interpreter = Interpreter(output_stream=output)
        interpreter.interpret(program)
        lines = output.getvalue().splitlines()
        assert lines == [
            'test1: 0',
            'test2: 42',
            '  got: hello',
            'test3: after PrintVal',
            '  done',
            'test4: both ran',
            'test5: i=1',
            'test5: i=2',
        ]

    def test_regression_script_gamma_execute_eval(self):
        """Execute, ExecuteGlobal, and Eval dynamic code execution."""
        script = RegressionTestsUtils.load_script('gamma-execute-eval.vbs')
        program = parse(script)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('totalCount') == 86
        assert interpreter._environment.get('passCount') == 86
        assert interpreter._environment.get('failCount') == 0

    def test_regression_script_delta_filesystemobject(self):
        """Scripting.FileSystemObject file system operations."""
        script = RegressionTestsUtils.load_script('delta-filesystemobject.vbs')
        program = parse(script)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('totalCount') == 64
        assert interpreter._environment.get('passCount') == 64
        assert interpreter._environment.get('failCount') == 0

    def test_regression_script_epsilon_cscript_compat(self):
        """cscript.exe compatibility: boolean display, bitwise Not/And/Or,
        Null concatenation, date literals, float precision, Hex, TypeName,
        VarType, IsObject, Len, Round, error numbers, and more.
        All expected values verified against cscript.exe //nologo on Windows 11."""
        script = RegressionTestsUtils.load_script('epsilon-cscript-compat.vbs')
        program = parse(script)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('totalCount') == 136
        assert interpreter._environment.get('passCount') == 136
        assert interpreter._environment.get('failCount') == 0
