"""Tests for the VBScript parser."""

from pybasil import (
    parse,
    AssignmentStatement,
    NumberLiteral,
    StringLiteral,
    BooleanLiteral,
    NothingLiteral,
    EmptyLiteral,
    NullLiteral,
)


class TestParserLiterals:
    """Test parsing of literal values."""

    def test_parse_integer(self):
        program = parse('x = 42')
        assert len(program.statements) == 1
        stmt = program.statements[0]
        assert isinstance(stmt, AssignmentStatement)
        assert stmt.variable == 'x'
        assert isinstance(stmt.expression, NumberLiteral)
        assert stmt.expression.value == 42

    def test_parse_float(self):
        program = parse('x = 3.14')
        stmt = program.statements[0]
        assert isinstance(stmt.expression, NumberLiteral)
        assert stmt.expression.value == 3.14

    def test_parse_scientific_notation(self):
        program = parse('x = 1.5e10')
        stmt = program.statements[0]
        assert isinstance(stmt.expression, NumberLiteral)
        assert stmt.expression.value == 1.5e10

    def test_parse_string(self):
        program = parse('x = "Hello, World!"')
        stmt = program.statements[0]
        assert isinstance(stmt.expression, StringLiteral)
        assert stmt.expression.value == 'Hello, World!'

    def test_parse_empty_string(self):
        program = parse('x = ""')
        stmt = program.statements[0]
        assert isinstance(stmt.expression, StringLiteral)
        assert stmt.expression.value == ''

    def test_parse_true(self):
        program = parse('x = True')
        stmt = program.statements[0]
        assert isinstance(stmt.expression, BooleanLiteral)
        assert stmt.expression.value is True

    def test_parse_false(self):
        program = parse('x = False')
        stmt = program.statements[0]
        assert isinstance(stmt.expression, BooleanLiteral)
        assert stmt.expression.value is False

    def test_parse_nothing(self):
        program = parse('x = Nothing')
        stmt = program.statements[0]
        assert isinstance(stmt.expression, NothingLiteral)

    def test_parse_empty(self):
        program = parse('x = Empty')
        stmt = program.statements[0]
        assert isinstance(stmt.expression, EmptyLiteral)

    def test_parse_null(self):
        program = parse('x = Null')
        stmt = program.statements[0]
        assert isinstance(stmt.expression, NullLiteral)

    def test_case_insensitive_true(self):
        program = parse('x = true')
        stmt = program.statements[0]
        assert isinstance(stmt.expression, BooleanLiteral)
        assert stmt.expression.value is True

    def test_case_insensitive_false(self):
        program = parse('x = FALSE')
        stmt = program.statements[0]
        assert isinstance(stmt.expression, BooleanLiteral)
        assert stmt.expression.value is False
