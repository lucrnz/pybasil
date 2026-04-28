"""Tests for the VBScript parser."""

from pybasil import (
    parse,
    SubStatement,
    FunctionStatement,
    Parameter,
)


class TestParserSubStatement:
    """Test parsing of Sub statements."""

    def test_parse_sub_no_params(self):
        program = parse("""
            Sub SayHello
                WScript.Echo "Hello"
            End Sub
        """)
        assert len(program.statements) == 1
        stmt = program.statements[0]
        assert isinstance(stmt, SubStatement)
        assert stmt.name == 'SayHello'
        assert len(stmt.parameters) == 0
        assert len(stmt.body) == 1

    def test_parse_sub_with_params(self):
        program = parse("""
            Sub Greet(name)
                WScript.Echo "Hello, " & name
            End Sub
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, SubStatement)
        assert stmt.name == 'Greet'
        assert len(stmt.parameters) == 1
        assert isinstance(stmt.parameters[0], Parameter)
        assert stmt.parameters[0].name == 'name'

    def test_parse_sub_with_byref_param(self):
        program = parse("""
            Sub Increment(ByRef x)
                x = x + 1
            End Sub
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, SubStatement)
        assert len(stmt.parameters) == 1
        assert stmt.parameters[0].is_byref is True

    def test_parse_sub_with_byval_param(self):
        program = parse("""
            Sub Process(ByVal value)
                WScript.Echo value
            End Sub
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, SubStatement)
        assert len(stmt.parameters) == 1
        assert stmt.parameters[0].is_byref is False  # ByVal means not ByRef

    def test_parse_sub_multiple_params(self):
        program = parse("""
            Sub AddValues(a, b, c)
                WScript.Echo a + b + c
            End Sub
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, SubStatement)
        assert len(stmt.parameters) == 3
        # Default is ByRef
        assert all(p.is_byref for p in stmt.parameters)

class TestParserFunctionStatement:
    """Test parsing of Function statements."""

    def test_parse_function_no_params(self):
        program = parse("""
            Function GetAnswer
                GetAnswer = 42
            End Function
        """)
        assert len(program.statements) == 1
        stmt = program.statements[0]
        assert isinstance(stmt, FunctionStatement)
        assert stmt.name == 'GetAnswer'
        assert len(stmt.parameters) == 0

    def test_parse_function_with_params(self):
        program = parse("""
            Function Add(a, b)
                Add = a + b
            End Function
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, FunctionStatement)
        assert stmt.name == 'Add'
        assert len(stmt.parameters) == 2

    def test_parse_function_with_mixed_params(self):
        program = parse("""
            Function Process(ByVal x, ByRef y, z)
                Process = x + y + z
            End Function
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, FunctionStatement)
        assert len(stmt.parameters) == 3
        assert stmt.parameters[0].is_byref is False  # ByVal
        assert stmt.parameters[1].is_byref is True  # ByRef
        assert stmt.parameters[2].is_byref is True  # Default (ByRef)
