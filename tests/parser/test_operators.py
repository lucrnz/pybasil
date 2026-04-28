"""Tests for the VBScript parser."""

from pybasil import (
    parse,
    BinaryExpression,
    UnaryExpression,
    ComparisonExpression,
    BinaryOp,
    UnaryOp,
    ComparisonOp,
)


class TestParserOperators:
    """Test parsing of operators."""

    def test_parse_addition(self):
        program = parse('x = 1 + 2')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.ADD

    def test_parse_subtraction(self):
        program = parse('x = 5 - 3')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.SUB

    def test_parse_multiplication(self):
        program = parse('x = 4 * 2')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.MUL

    def test_parse_division(self):
        program = parse('x = 10 / 2')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.DIV

    def test_parse_integer_division(self):
        program = parse('x = 10 \\ 3')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.INTDIV

    def test_parse_modulo(self):
        program = parse('x = 10 Mod 3')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.MOD

    def test_parse_exponentiation(self):
        program = parse('x = 2 ^ 3')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.POW

    def test_parse_concatenation(self):
        program = parse('x = "Hello" & " World"')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.CONCAT

    def test_parse_negation(self):
        program = parse('x = -5')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, UnaryExpression)
        assert expr.operator == UnaryOp.NEG

    def test_parse_not(self):
        program = parse('x = Not True')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, UnaryExpression)
        assert expr.operator == UnaryOp.NOT

    def test_parse_and(self):
        program = parse('x = True And False')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.AND

    def test_parse_or(self):
        program = parse('x = True Or False')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.OR

    def test_parse_xor(self):
        program = parse('x = True Xor False')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.XOR

    def test_parse_eqv(self):
        program = parse('x = True Eqv False')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.EQV

    def test_parse_imp(self):
        program = parse('x = True Imp False')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.IMP

class TestParserComparisons:
    """Test parsing of comparison operators."""

    def test_parse_equals(self):
        program = parse('x = (a = b)')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, ComparisonExpression)
        assert expr.operator == ComparisonOp.EQ

    def test_parse_not_equals(self):
        program = parse('x = (a <> b)')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, ComparisonExpression)
        assert expr.operator == ComparisonOp.NE

    def test_parse_less_than(self):
        program = parse('x = (a < b)')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, ComparisonExpression)
        assert expr.operator == ComparisonOp.LT

    def test_parse_greater_than(self):
        program = parse('x = (a > b)')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, ComparisonExpression)
        assert expr.operator == ComparisonOp.GT

    def test_parse_less_equal(self):
        program = parse('x = (a <= b)')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, ComparisonExpression)
        assert expr.operator == ComparisonOp.LE

    def test_parse_greater_equal(self):
        program = parse('x = (a >= b)')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, ComparisonExpression)
        assert expr.operator == ComparisonOp.GE

    def test_parse_is(self):
        program = parse('x = (a Is b)')
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, ComparisonExpression)
        assert expr.operator == ComparisonOp.IS

class TestParserPrecedence:
    """Test operator precedence."""

    def test_multiplication_before_addition(self):
        program = parse('x = 1 + 2 * 3')
        stmt = program.statements[0]
        expr = stmt.expression
        # Should be: 1 + (2 * 3)
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.ADD
        assert isinstance(expr.right, BinaryExpression)
        assert expr.right.operator == BinaryOp.MUL

    def test_exponentiation_right_associative(self):
        program = parse('x = 2 ^ 3 ^ 2')
        stmt = program.statements[0]
        expr = stmt.expression
        # Should be: 2 ^ (3 ^ 2) = 2 ^ 9 = 512
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.POW

    def test_parentheses_override_precedence(self):
        program = parse('x = (1 + 2) * 3')
        stmt = program.statements[0]
        expr = stmt.expression
        # Should be: (1 + 2) * 3
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.MUL
        assert isinstance(expr.left, BinaryExpression)
        assert expr.left.operator == BinaryOp.ADD
