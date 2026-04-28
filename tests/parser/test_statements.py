"""Tests for the VBScript parser."""

from pybasil import (
    parse,
    DimStatement,
    AssignmentStatement,
    CallStatement,
    ExpressionStatement,
)


class TestParserMultipleStatements:
    """Test parsing multiple statements."""

    def test_multiple_statements(self):
        program = parse("""
            Dim x
            x = 5
            y = x + 1
        """)
        assert len(program.statements) == 3
        assert isinstance(program.statements[0], DimStatement)
        assert isinstance(program.statements[1], AssignmentStatement)
        assert isinstance(program.statements[2], AssignmentStatement)

class TestParserCallStatement:
    """Test parsing of Call statements."""

    def test_parse_call_with_parens(self):
        program = parse('Call MySub("Hello")')
        stmt = program.statements[0]
        assert isinstance(stmt, CallStatement)
        assert stmt.name == 'MySub'
        assert len(stmt.arguments) == 1

    def test_parse_call_no_args(self):
        program = parse('Call MySub()')
        stmt = program.statements[0]
        assert isinstance(stmt, CallStatement)
        assert stmt.name == 'MySub'
        assert len(stmt.arguments) == 0

    def test_parse_call_multiple_args(self):
        program = parse('Call MySub("Hello", "World", 42)')
        stmt = program.statements[0]
        assert isinstance(stmt, CallStatement)
        assert stmt.name == 'MySub'
        assert len(stmt.arguments) == 3

    def test_parse_implicit_call(self):
        # Implicit call (without Call keyword) is parsed as expression statement
        program = parse('MySub "Hello"')
        stmt = program.statements[0]
        assert isinstance(stmt, ExpressionStatement)
