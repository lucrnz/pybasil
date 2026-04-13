"""Tests for the VBScript parser."""

from pybasil import (
    parse,
    DimStatement,
    AssignmentStatement,
    SetStatement,
    CallStatement,
    ExpressionStatement,
    NumberLiteral,
    StringLiteral,
    BooleanLiteral,
    NothingLiteral,
    EmptyLiteral,
    NullLiteral,
    Identifier,
    BinaryExpression,
    UnaryExpression,
    ComparisonExpression,
    BinaryOp,
    UnaryOp,
    ComparisonOp,
    IfStatement,
    ElseIfClause,
    ElseClause,
    ForStatement,
    WhileStatement,
    DoLoopStatement,
    ExitStatement,
    SubStatement,
    FunctionStatement,
    Parameter,
    ExitType,
    LoopConditionType,
)


class TestParserLiterals:
    """Test parsing of literal values."""

    def test_parse_integer(self):
        program = parse("x = 42")
        assert len(program.statements) == 1
        stmt = program.statements[0]
        assert isinstance(stmt, AssignmentStatement)
        assert stmt.variable == "x"
        assert isinstance(stmt.expression, NumberLiteral)
        assert stmt.expression.value == 42

    def test_parse_float(self):
        program = parse("x = 3.14")
        stmt = program.statements[0]
        assert isinstance(stmt.expression, NumberLiteral)
        assert stmt.expression.value == 3.14

    def test_parse_scientific_notation(self):
        program = parse("x = 1.5e10")
        stmt = program.statements[0]
        assert isinstance(stmt.expression, NumberLiteral)
        assert stmt.expression.value == 1.5e10

    def test_parse_string(self):
        program = parse('x = "Hello, World!"')
        stmt = program.statements[0]
        assert isinstance(stmt.expression, StringLiteral)
        assert stmt.expression.value == "Hello, World!"

    def test_parse_empty_string(self):
        program = parse('x = ""')
        stmt = program.statements[0]
        assert isinstance(stmt.expression, StringLiteral)
        assert stmt.expression.value == ""

    def test_parse_true(self):
        program = parse("x = True")
        stmt = program.statements[0]
        assert isinstance(stmt.expression, BooleanLiteral)
        assert stmt.expression.value is True

    def test_parse_false(self):
        program = parse("x = False")
        stmt = program.statements[0]
        assert isinstance(stmt.expression, BooleanLiteral)
        assert stmt.expression.value is False

    def test_parse_nothing(self):
        program = parse("x = Nothing")
        stmt = program.statements[0]
        assert isinstance(stmt.expression, NothingLiteral)

    def test_parse_empty(self):
        program = parse("x = Empty")
        stmt = program.statements[0]
        assert isinstance(stmt.expression, EmptyLiteral)

    def test_parse_null(self):
        program = parse("x = Null")
        stmt = program.statements[0]
        assert isinstance(stmt.expression, NullLiteral)

    def test_case_insensitive_true(self):
        program = parse("x = true")
        stmt = program.statements[0]
        assert isinstance(stmt.expression, BooleanLiteral)
        assert stmt.expression.value is True

    def test_case_insensitive_false(self):
        program = parse("x = FALSE")
        stmt = program.statements[0]
        assert isinstance(stmt.expression, BooleanLiteral)
        assert stmt.expression.value is False


class TestParserVariables:
    """Test parsing of variable declarations and lookups."""

    def test_parse_dim_single(self):
        program = parse("Dim x")
        assert len(program.statements) == 1
        stmt = program.statements[0]
        assert isinstance(stmt, DimStatement)
        assert len(stmt.variables) == 1
        assert stmt.variables[0].name == "x"
        assert stmt.variables[0].dimensions is None

    def test_parse_dim_multiple(self):
        program = parse("Dim x, y, z")
        stmt = program.statements[0]
        assert isinstance(stmt, DimStatement)
        assert len(stmt.variables) == 3
        assert stmt.variables[0].name == "x"
        assert stmt.variables[1].name == "y"
        assert stmt.variables[2].name == "z"

    def test_parse_dim_case_insensitive(self):
        program = parse("DIM x, Y, Z")
        stmt = program.statements[0]
        assert isinstance(stmt, DimStatement)
        assert len(stmt.variables) == 3
        assert stmt.variables[0].name == "x"
        assert stmt.variables[1].name == "Y"
        assert stmt.variables[2].name == "Z"

    def test_parse_assignment(self):
        program = parse("x = 42")
        stmt = program.statements[0]
        assert isinstance(stmt, AssignmentStatement)
        assert stmt.variable == "x"

    def test_parse_assignment_with_let(self):
        program = parse("Let x = 42")
        stmt = program.statements[0]
        assert isinstance(stmt, AssignmentStatement)
        assert stmt.variable == "x"

    def test_parse_set_statement(self):
        program = parse('Set obj = CreateObject("Scripting.FileSystemObject")')
        stmt = program.statements[0]
        assert isinstance(stmt, SetStatement)
        assert stmt.variable == "obj"

    def test_parse_variable_lookup(self):
        program = parse("x = y")
        stmt = program.statements[0]
        assert isinstance(stmt.expression, Identifier)
        assert stmt.expression.name == "y"


class TestParserOperators:
    """Test parsing of operators."""

    def test_parse_addition(self):
        program = parse("x = 1 + 2")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.ADD

    def test_parse_subtraction(self):
        program = parse("x = 5 - 3")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.SUB

    def test_parse_multiplication(self):
        program = parse("x = 4 * 2")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.MUL

    def test_parse_division(self):
        program = parse("x = 10 / 2")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.DIV

    def test_parse_integer_division(self):
        program = parse("x = 10 \\ 3")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.INTDIV

    def test_parse_modulo(self):
        program = parse("x = 10 Mod 3")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.MOD

    def test_parse_exponentiation(self):
        program = parse("x = 2 ^ 3")
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
        program = parse("x = -5")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, UnaryExpression)
        assert expr.operator == UnaryOp.NEG

    def test_parse_not(self):
        program = parse("x = Not True")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, UnaryExpression)
        assert expr.operator == UnaryOp.NOT

    def test_parse_and(self):
        program = parse("x = True And False")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.AND

    def test_parse_or(self):
        program = parse("x = True Or False")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.OR

    def test_parse_xor(self):
        program = parse("x = True Xor False")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.XOR

    def test_parse_eqv(self):
        program = parse("x = True Eqv False")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.EQV

    def test_parse_imp(self):
        program = parse("x = True Imp False")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.IMP


class TestParserComparisons:
    """Test parsing of comparison operators."""

    def test_parse_equals(self):
        program = parse("x = (a = b)")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, ComparisonExpression)
        assert expr.operator == ComparisonOp.EQ

    def test_parse_not_equals(self):
        program = parse("x = (a <> b)")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, ComparisonExpression)
        assert expr.operator == ComparisonOp.NE

    def test_parse_less_than(self):
        program = parse("x = (a < b)")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, ComparisonExpression)
        assert expr.operator == ComparisonOp.LT

    def test_parse_greater_than(self):
        program = parse("x = (a > b)")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, ComparisonExpression)
        assert expr.operator == ComparisonOp.GT

    def test_parse_less_equal(self):
        program = parse("x = (a <= b)")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, ComparisonExpression)
        assert expr.operator == ComparisonOp.LE

    def test_parse_greater_equal(self):
        program = parse("x = (a >= b)")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, ComparisonExpression)
        assert expr.operator == ComparisonOp.GE

    def test_parse_is(self):
        program = parse("x = (a Is b)")
        stmt = program.statements[0]
        expr = stmt.expression
        assert isinstance(expr, ComparisonExpression)
        assert expr.operator == ComparisonOp.IS


class TestParserPrecedence:
    """Test operator precedence."""

    def test_multiplication_before_addition(self):
        program = parse("x = 1 + 2 * 3")
        stmt = program.statements[0]
        expr = stmt.expression
        # Should be: 1 + (2 * 3)
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.ADD
        assert isinstance(expr.right, BinaryExpression)
        assert expr.right.operator == BinaryOp.MUL

    def test_exponentiation_right_associative(self):
        program = parse("x = 2 ^ 3 ^ 2")
        stmt = program.statements[0]
        expr = stmt.expression
        # Should be: 2 ^ (3 ^ 2) = 2 ^ 9 = 512
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.POW

    def test_parentheses_override_precedence(self):
        program = parse("x = (1 + 2) * 3")
        stmt = program.statements[0]
        expr = stmt.expression
        # Should be: (1 + 2) * 3
        assert isinstance(expr, BinaryExpression)
        assert expr.operator == BinaryOp.MUL
        assert isinstance(expr.left, BinaryExpression)
        assert expr.left.operator == BinaryOp.ADD


class TestParserComments:
    """Test parsing of comments."""

    def test_single_quote_comment(self):
        program = parse("x = 5 ' This is a comment")
        assert len(program.statements) == 1
        stmt = program.statements[0]
        assert isinstance(stmt, AssignmentStatement)

    def test_rem_comment(self):
        program = parse("x = 5 Rem This is a comment")
        assert len(program.statements) == 1

    def test_comment_only_line(self):
        program = parse("' This is a full line comment")
        assert len(program.statements) == 0


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
        assert stmt.variable == "i"
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
        program = parse("Exit For")
        assert len(program.statements) == 1
        stmt = program.statements[0]
        assert isinstance(stmt, ExitStatement)
        assert stmt.exit_type == ExitType.FOR

    def test_parse_exit_do(self):
        program = parse("Exit Do")
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
        assert stmt.name == "SayHello"
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
        assert stmt.name == "Greet"
        assert len(stmt.parameters) == 1
        assert isinstance(stmt.parameters[0], Parameter)
        assert stmt.parameters[0].name == "name"

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
        assert stmt.name == "GetAnswer"
        assert len(stmt.parameters) == 0

    def test_parse_function_with_params(self):
        program = parse("""
            Function Add(a, b)
                Add = a + b
            End Function
        """)
        stmt = program.statements[0]
        assert isinstance(stmt, FunctionStatement)
        assert stmt.name == "Add"
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


class TestParserCallStatement:
    """Test parsing of Call statements."""

    def test_parse_call_with_parens(self):
        program = parse('Call MySub("Hello")')
        stmt = program.statements[0]
        assert isinstance(stmt, CallStatement)
        assert stmt.name == "MySub"
        assert len(stmt.arguments) == 1

    def test_parse_call_no_args(self):
        program = parse("Call MySub()")
        stmt = program.statements[0]
        assert isinstance(stmt, CallStatement)
        assert stmt.name == "MySub"
        assert len(stmt.arguments) == 0

    def test_parse_call_multiple_args(self):
        program = parse('Call MySub("Hello", "World", 42)')
        stmt = program.statements[0]
        assert isinstance(stmt, CallStatement)
        assert stmt.name == "MySub"
        assert len(stmt.arguments) == 3

    def test_parse_implicit_call(self):
        # Implicit call (without Call keyword) is parsed as expression statement
        program = parse('MySub "Hello"')
        stmt = program.statements[0]
        assert isinstance(stmt, ExpressionStatement)
