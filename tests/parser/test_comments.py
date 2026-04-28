"""Tests for the VBScript parser."""

from pybasil import (
    parse,
    AssignmentStatement,
    StringLiteral,
)


class TestParserComments:
    """Test parsing of comments."""

    def test_single_quote_comment(self):
        program = parse("x = 5 ' This is a comment")
        assert len(program.statements) == 1
        stmt = program.statements[0]
        assert isinstance(stmt, AssignmentStatement)

    def test_rem_comment(self):
        program = parse('x = 5 Rem This is a comment')
        assert len(program.statements) == 1

    def test_rem_inside_string_literal_not_replaced(self):
        program = parse('x = "Rem this is not a comment"')
        assert len(program.statements) == 1
        stmt = program.statements[0]
        assert isinstance(stmt, AssignmentStatement)
        assert isinstance(stmt.expression, StringLiteral)
        assert stmt.expression.value == "Rem this is not a comment"

    def test_comment_only_line(self):
        program = parse("' This is a full line comment")
        assert len(program.statements) == 0
