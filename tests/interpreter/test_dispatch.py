"""Tests for the VBScript interpreter."""

import pytest
import io
from pybasil import (
    Interpreter,
    parse,
)


class TestBinopDispatch:
    """Test binary operator dispatch dict and numeric fast-paths."""

    def test_int_add_fast_path(self):
        program = parse('x = 3 + 4\nWScript.Echo x')
        output = io.StringIO()
        Interpreter(output_stream=output).interpret(program)
        assert output.getvalue().strip() == '7'

    def test_int_sub_fast_path(self):
        program = parse('x = 10 - 3\nWScript.Echo x')
        output = io.StringIO()
        Interpreter(output_stream=output).interpret(program)
        assert output.getvalue().strip() == '7'

    def test_int_mul_fast_path(self):
        program = parse('x = 6 * 7\nWScript.Echo x')
        output = io.StringIO()
        Interpreter(output_stream=output).interpret(program)
        assert output.getvalue().strip() == '42'

    def test_int_div_fast_path(self):
        program = parse('x = 10 / 2\nWScript.Echo x')
        output = io.StringIO()
        Interpreter(output_stream=output).interpret(program)
        assert output.getvalue().strip() == '5'

    def test_int_intdiv_fast_path(self):
        program = parse('x = 7 \\ 2\nWScript.Echo x')
        output = io.StringIO()
        Interpreter(output_stream=output).interpret(program)
        assert output.getvalue().strip() == '3'

    def test_int_mod_fast_path(self):
        program = parse('x = 10 Mod 3\nWScript.Echo x')
        output = io.StringIO()
        Interpreter(output_stream=output).interpret(program)
        assert output.getvalue().strip() == '1'

    def test_int_pow_fast_path(self):
        program = parse('x = 2 ^ 10\nWScript.Echo x')
        output = io.StringIO()
        Interpreter(output_stream=output).interpret(program)
        assert output.getvalue().strip() == '1024'

    def test_int_div_by_zero(self):
        program = parse('x = 10 / 0')
        with pytest.raises(Exception, match='Division by zero'):
            Interpreter().interpret(program)

    def test_int_intdiv_by_zero(self):
        program = parse('x = 10 \\ 0')
        with pytest.raises(Exception, match='Division by zero'):
            Interpreter().interpret(program)

    def test_int_mod_by_zero(self):
        program = parse('x = 10 Mod 0')
        with pytest.raises(Exception, match='Division by zero'):
            Interpreter().interpret(program)

    def test_concat_dispatch(self):
        program = parse('x = "hello" & " world"\nWScript.Echo x')
        output = io.StringIO()
        Interpreter(output_stream=output).interpret(program)
        assert output.getvalue().strip() == 'hello world'

    def test_empty_operand_falls_through(self):
        """Empty values should coerce and use the dispatch path."""
        program = parse('Dim x\ny = x + 1\nWScript.Echo y')
        output = io.StringIO()
        Interpreter(output_stream=output).interpret(program)
        assert output.getvalue().strip() == '1'

    def test_binop_dispatch_all_ops_covered(self):
        from pybasil.ast_nodes import BinaryOp
        interp = Interpreter()
        for op in BinaryOp:
            assert op in interp._binop_dispatch, f'{op} missing from _binop_dispatch'

class TestNodeDispatchTables:
    """Test that dispatch tables cover all AST node types."""

    def test_execute_dispatch_covers_statements(self):
        interp = Interpreter()
        from pybasil.ast_nodes import (
            DimStatement, AssignmentStatement, ExpressionStatement,
            IfStatement, ForStatement, WhileStatement, SubStatement,
            FunctionStatement, ExitStatement,
        )
        for cls in (DimStatement, AssignmentStatement, ExpressionStatement,
                    IfStatement, ForStatement, WhileStatement, SubStatement,
                    FunctionStatement, ExitStatement):
            assert cls in interp._execute_dispatch, f'{cls.__name__} missing from execute dispatch'

    def test_evaluate_dispatch_covers_expressions(self):
        interp = Interpreter()
        from pybasil.ast_nodes import (
            NumberLiteral, StringLiteral, BooleanLiteral, Identifier,
            BinaryExpression, UnaryExpression, ComparisonExpression,
            MemberAccess, FunctionCall, MethodCall, ArrayAccess,
        )
        for cls in (NumberLiteral, StringLiteral, BooleanLiteral, Identifier,
                    BinaryExpression, UnaryExpression, ComparisonExpression,
                    MemberAccess, FunctionCall, MethodCall, ArrayAccess):
            assert cls in interp._evaluate_dispatch, f'{cls.__name__} missing from evaluate dispatch'

    def test_dispatch_matches_getattr_approach(self):
        """Ensure the dict dispatch finds the same handler as the old getattr approach."""
        interp = Interpreter()
        from pybasil.ast_nodes import NumberLiteral, DimStatement
        assert interp._evaluate_dispatch[NumberLiteral] == interp._evaluate_NumberLiteral
        assert interp._execute_dispatch[DimStatement] == interp._execute_DimStatement

class TestBuiltinsDictIntegrity:
    """Ensure builtins dictionary has no issues."""

    def test_isnumeric_registered_once(self):
        interpreter = Interpreter()
        assert 'isnumeric' in interpreter._builtins
        assert interpreter._builtins['isnumeric'] is not None

class TestCachedLowercaseIdentifiers:
    """Test pre-lowered Identifier._lower used by the interpreter."""

    def test_identifier_lower_computed(self):
        from pybasil.ast_nodes import Identifier
        ident = Identifier(name='MyVar')
        assert ident._lower == 'myvar'

    def test_identifier_lower_already_lowercase(self):
        from pybasil.ast_nodes import Identifier
        ident = Identifier(name='x')
        assert ident._lower == 'x'

    def test_identifier_lookup_uses_cached_lower(self):
        """Variable set with mixed case should be found via cached lower."""
        program = parse('MyVar = 42\nWScript.Echo MyVar')
        output = io.StringIO()
        Interpreter(output_stream=output).interpret(program)
        assert output.getvalue().strip() == '42'

    def test_identifier_case_insensitive_via_cache(self):
        program = parse('myvar = 10\nWScript.Echo MYVAR')
        output = io.StringIO()
        Interpreter(output_stream=output).interpret(program)
        assert output.getvalue().strip() == '10'

    def test_identifier_in_nested_scope(self):
        program = parse('x = 5\nFunction GetX()\nGetX = x\nEnd Function\nWScript.Echo GetX()')
        output = io.StringIO()
        Interpreter(output_stream=output).interpret(program)
        assert output.getvalue().strip() == '5'
