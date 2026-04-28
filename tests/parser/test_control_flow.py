"""Tests for the VBScript parser."""

from pybasil import (
    parse,
    NumberLiteral,
    BooleanLiteral,
    UnaryExpression,
    ComparisonExpression,
    UnaryOp,
    IfStatement,
    ElseIfClause,
    ElseClause,
    ForStatement,
    WhileStatement,
    DoLoopStatement,
    ExitStatement,
    ExitType,
    LoopConditionType,
)


class TestParserIfStatement:
    """Test parsing of If statements."""

    def test_parse_if_then(self):
        program = parse("""
            If True Then
                x = 1
            End If
        """)
        assert len(program.statements) == 1
        stmt = program.statements[0]
        assert isinstance(stmt, IfStatement)
        assert isinstance(stmt.condition, BooleanLiteral)
        assert stmt.condition.value is True
        assert len(stmt.then_body) == 1

    def test_parse_if_then_else(self):
        program = parse("""
            If True Then
                x = 1
            Else
                x = 2
            End If
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, IfStatement)
        assert stmt.else_clause is not None
        assert isinstance(stmt.else_clause, ElseClause)
        assert len(stmt.else_clause.body) == 1

    def test_parse_if_elseif_else(self):
        program = parse("""
            If x = 1 Then
                y = 1
            ElseIf x = 2 Then
                y = 2
            Else
                y = 3
            End If
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, IfStatement)
        assert len(stmt.elseif_clauses) == 1
        assert isinstance(stmt.elseif_clauses[0], ElseIfClause)
        assert stmt.else_clause is not None

    def test_parse_if_multiple_elseif(self):
        program = parse("""
            If x = 1 Then
                y = 1
            ElseIf x = 2 Then
                y = 2
            ElseIf x = 3 Then
                y = 3
            End If
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, IfStatement)
        assert len(stmt.elseif_clauses) == 2

    def test_parse_if_nested(self):
        program = parse("""
            If True Then
                If False Then
                    x = 1
                End If
            End If
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, IfStatement)
        assert isinstance(stmt.then_body[0], IfStatement)

class TestParserForStatement:
    """Test parsing of For statements."""

    def test_parse_for_basic(self):
        program = parse("""
            For i = 1 To 10
                x = i
            Next
        """)
        assert len(program.statements) == 1
        stmt = program.statements[0]
        assert isinstance(stmt, ForStatement)
        assert stmt.variable == 'i'
        assert isinstance(stmt.start, NumberLiteral)
        assert stmt.start.value == 1
        assert isinstance(stmt.end, NumberLiteral)
        assert stmt.end.value == 10
        assert stmt.step is None
        assert len(stmt.body) == 1

    def test_parse_for_with_step(self):
        program = parse("""
            For i = 0 To 10 Step 2
                x = i
            Next
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, ForStatement)
        assert stmt.step is not None
        assert isinstance(stmt.step, NumberLiteral)
        assert stmt.step.value == 2

    def test_parse_for_negative_step(self):
        program = parse("""
            For i = 10 To 1 Step -1
                x = i
            Next
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, ForStatement)
        assert stmt.step is not None
        # Negative step is parsed as a unary expression
        assert isinstance(stmt.step, UnaryExpression)
        assert stmt.step.operator == UnaryOp.NEG

    def test_parse_for_nested(self):
        program = parse("""
            For i = 1 To 5
                For j = 1 To 5
                    x = i + j
                Next
            Next
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, ForStatement)
        assert isinstance(stmt.body[0], ForStatement)

class TestParserWhileStatement:
    """Test parsing of While statements."""

    def test_parse_while_basic(self):
        program = parse("""
            While x < 10
                x = x + 1
            Wend
        """)
        assert len(program.statements) == 1
        stmt = program.statements[0]
        assert isinstance(stmt, WhileStatement)
        assert isinstance(stmt.condition, ComparisonExpression)
        assert len(stmt.body) == 1

    def test_parse_while_nested(self):
        program = parse("""
            While x < 10
                While y < 10
                    y = y + 1
                Wend
                x = x + 1
            Wend
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, WhileStatement)
        assert isinstance(stmt.body[0], WhileStatement)

class TestParserDoLoop:
    """Test parsing of Do Loop statements."""

    def test_parse_do_while_pre_test(self):
        program = parse("""
            Do While x < 10
                x = x + 1
            Loop
        """)
        assert len(program.statements) == 1
        stmt = program.statements[0]
        assert isinstance(stmt, DoLoopStatement)
        assert stmt.pre_condition is not None
        assert stmt.pre_condition.condition_type == LoopConditionType.WHILE
        assert stmt.post_condition is None

    def test_parse_do_until_pre_test(self):
        program = parse("""
            Do Until x >= 10
                x = x + 1
            Loop
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, DoLoopStatement)
        assert stmt.pre_condition is not None
        assert stmt.pre_condition.condition_type == LoopConditionType.UNTIL

    def test_parse_do_loop_while_post_test(self):
        program = parse("""
            Do
                x = x + 1
            Loop While x < 10
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, DoLoopStatement)
        assert stmt.pre_condition is None
        assert stmt.post_condition is not None
        assert stmt.post_condition.condition_type == LoopConditionType.WHILE

    def test_parse_do_loop_until_post_test(self):
        program = parse("""
            Do
                x = x + 1
            Loop Until x >= 10
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, DoLoopStatement)
        assert stmt.pre_condition is None
        assert stmt.post_condition is not None
        assert stmt.post_condition.condition_type == LoopConditionType.UNTIL

    def test_parse_do_infinite(self):
        program = parse("""
            Do
                x = x + 1
            Loop
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, DoLoopStatement)
        assert stmt.pre_condition is None
        assert stmt.post_condition is None

class TestParserExitStatement:
    """Test parsing of Exit statements."""

    def test_parse_exit_for(self):
        program = parse('Exit For')
        assert len(program.statements) == 1
        stmt = program.statements[0]
        assert isinstance(stmt, ExitStatement)
        assert stmt.exit_type == ExitType.FOR

    def test_parse_exit_do(self):
        program = parse('Exit Do')
        assert len(program.statements) == 1
        stmt = program.statements[0]
        assert isinstance(stmt, ExitStatement)
        assert stmt.exit_type == ExitType.DO

    def test_parse_exit_for_in_loop(self):
        program = parse("""
            For i = 1 To 10
                Exit For
            Next
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, ForStatement)
        assert isinstance(stmt.body[0], ExitStatement)
        assert stmt.body[0].exit_type == ExitType.FOR

    def test_parse_exit_do_in_loop(self):
        program = parse("""
            Do While True
                Exit Do
            Loop
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, DoLoopStatement)
        assert isinstance(stmt.body[0], ExitStatement)
        assert stmt.body[0].exit_type == ExitType.DO
