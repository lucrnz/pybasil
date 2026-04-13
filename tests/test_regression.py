"""Regression tests for the interpreter."""

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
