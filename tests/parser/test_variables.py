"""Tests for the VBScript parser."""

from pybasil import (
    parse,
    DimStatement,
    AssignmentStatement,
    SetStatement,
    Identifier,
)


class TestParserVariables:
    """Test parsing of variable declarations and lookups."""

    def test_parse_dim_single(self):
        program = parse('Dim x')
        assert len(program.statements) == 1
        stmt = program.statements[0]
        assert isinstance(stmt, DimStatement)
        assert len(stmt.variables) == 1
        assert stmt.variables[0].name == 'x'
        assert stmt.variables[0].dimensions is None

    def test_parse_dim_multiple(self):
        program = parse('Dim x, y, z')
        stmt = program.statements[0]
        assert isinstance(stmt, DimStatement)
        assert len(stmt.variables) == 3
        assert stmt.variables[0].name == 'x'
        assert stmt.variables[1].name == 'y'
        assert stmt.variables[2].name == 'z'

    def test_parse_dim_case_insensitive(self):
        program = parse('DIM x, Y, Z')
        stmt = program.statements[0]
        assert isinstance(stmt, DimStatement)
        assert len(stmt.variables) == 3
        assert stmt.variables[0].name == 'x'
        assert stmt.variables[1].name == 'Y'
        assert stmt.variables[2].name == 'Z'

    def test_parse_assignment(self):
        program = parse('x = 42')
        stmt = program.statements[0]
        assert isinstance(stmt, AssignmentStatement)
        assert stmt.variable == 'x'

    def test_parse_assignment_with_let(self):
        program = parse('Let x = 42')
        stmt = program.statements[0]
        assert isinstance(stmt, AssignmentStatement)
        assert stmt.variable == 'x'

    def test_parse_set_statement(self):
        program = parse('Set obj = CreateObject("Scripting.FileSystemObject")')
        stmt = program.statements[0]
        assert isinstance(stmt, SetStatement)
        assert stmt.variable == 'obj'

    def test_parse_variable_lookup(self):
        program = parse('x = y')
        stmt = program.statements[0]
        assert isinstance(stmt.expression, Identifier)
        assert stmt.expression.name == 'y'
