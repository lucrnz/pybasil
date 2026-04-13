"""VBScript Parser using Lark."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Iterator, List, Optional

from lark import Lark, Transformer, Token

from .ast_nodes import (
    ASTNode,
    Program,
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
    MemberAccess,
    FunctionCall,
    MethodCall,
    NewExpression,
    ArrayAccess,
    DimVariable,
    DimStatement,
    AssignmentStatement,
    SetStatement,
    PropertyAssignmentStatement,
    CallStatement,
    ExpressionStatement,
    IfStatement,
    ElseIfClause,
    ElseClause,
    CaseClause,
    CaseElseClause,
    SelectCaseStatement,
    ForStatement,
    ForEachStatement,
    WhileStatement,
    DoLoopStatement,
    LoopCondition,
    ExitStatement,
    SubStatement,
    FunctionStatement,
    Parameter,
    BinaryOp,
    UnaryOp,
    ComparisonOp,
    ExitType,
    LoopConditionType,
    OnErrorResumeNextStatement,
    OnErrorGoToStatement,
    ReDimStatement,
    EraseStatement,
)


class VBScriptTransformer(Transformer):
    """Transforms Lark parse tree into AST nodes."""

    def start(self, items: List[ASTNode]) -> Program:
        """Transform start rule."""
        # Filter out None values (from REM comments)
        items = [i for i in items if i is not None]
        return Program(statements=items)

    def statement(self, items: List[ASTNode]) -> ASTNode:
        # Filter out None values (from REM comments)
        items = [i for i in items if i is not None]
        if not items:
            return None
        return items[0]

    def block_statement(self, items: List[ASTNode]) -> ASTNode:
        """Transform a block statement."""
        items = [i for i in items if i is not None]
        if not items:
            return None
        return items[0]

    def dim_statement(self, items: List) -> DimStatement:
        """Transform Dim statement with optional array dimensions."""
        # items: [DIM_KW, dim_var, ("," dim_var)*]
        variables = []
        for item in items:
            if isinstance(item, DimVariable):
                variables.append(item)
        return DimStatement(variables=variables)

    def dim_var(self, items: List) -> DimVariable:
        """Transform a Dim variable declaration."""
        # items: [identifier] or [identifier, "(", dim_dimensions?, ")"]
        name = ''
        dimensions = None

        for item in items:
            if isinstance(item, Identifier):
                name = item.name
            elif isinstance(item, list):
                # dim_dimensions - list of dimension sizes
                dimensions = item
            elif item is None:
                # Empty parens - dynamic array
                dimensions = []

        return DimVariable(name=name, dimensions=dimensions)

    def dim_dimensions(self, items: List) -> List[ASTNode]:
        """Transform array dimension list."""
        return items

    def dim_size(self, items: List) -> ASTNode:
        """Transform a single dimension size."""
        return items[0]

    def assignment_statement(self, items: List) -> AssignmentStatement:
        """Transform assignment statement, supporting array element assignment."""
        # items: [LET_KW?, identifier, ("(" arg_list ")")?, "=", expression]
        # The optional part is either a list (indices) or None

        # Filter out the EQUAL token and LET_KW
        filtered = []
        for item in items:
            if isinstance(item, Token):
                if item.type == 'LET_KW':
                    continue
                # Skip the EQUAL token
                continue
            filtered.append(item)

        # Now filtered should be: [identifier, optional indices list, expression]
        if len(filtered) < 2:
            return None

        var_name = None
        indices = None
        expr = None

        # First item should be the identifier (variable name)
        if isinstance(filtered[0], Identifier):
            var_name = filtered[0].name

        # Check if there are indices (array assignment)
        if len(filtered) >= 3 and isinstance(filtered[1], list):
            indices = filtered[1]
            expr = filtered[2]
        elif len(filtered) >= 2:
            # No indices, so filtered[1] is the expression
            expr = filtered[1]

        return AssignmentStatement(variable=var_name, indices=indices, expression=expr)

    def set_statement(self, items: List) -> SetStatement:
        """Transform Set statement, supporting array element assignment."""
        # items: [SET_KW, identifier, ("(" arg_list ")")?, "=", expression]

        # Filter out the EQUAL token and SET_KW
        filtered = []
        for item in items:
            if isinstance(item, Token):
                if item.type == 'SET_KW':
                    continue
                continue
            filtered.append(item)

        if len(filtered) < 2:
            return None

        var_name = None
        indices = None
        expr = None

        # First item should be the identifier (variable name)
        if isinstance(filtered[0], Identifier):
            var_name = filtered[0].name

        # Check if there are indices (array assignment)
        if len(filtered) >= 3 and isinstance(filtered[1], list):
            indices = filtered[1]
            expr = filtered[2]
        elif len(filtered) >= 2:
            expr = filtered[1]

        return SetStatement(variable=var_name, indices=indices, expression=expr)

    def property_assignment_statement(self, items: List) -> PropertyAssignmentStatement:
        """Transform property assignment statement."""
        # items: [target, "=", expression]
        # Filter out the EQUAL token
        filtered = [item for item in items if not isinstance(item, Token)]

        if len(filtered) >= 2:
            target = filtered[0]
            expr = filtered[1]
            return PropertyAssignmentStatement(target=target, expression=expr)

        return None

    def expression_statement(self, items: List) -> ExpressionStatement:
        if len(items) == 1:
            return ExpressionStatement(expression=items[0])

        # Handle method call without parentheses: WScript.Echo "Hello"
        # Or implicit procedure call: MySub "Hello"
        # items: [expression, arg_list]
        expr = items[0]
        args = items[1] if len(items) > 1 and isinstance(items[1], list) else []

        # If the expression is a MemberAccess, convert it to a MethodCall
        if isinstance(expr, MemberAccess):
            return ExpressionStatement(
                expression=MethodCall(
                    object=expr.object, method=expr.member, arguments=args
                )
            )

        # If the expression is an Identifier, convert it to a FunctionCall (implicit procedure call)
        if isinstance(expr, Identifier):
            return ExpressionStatement(
                expression=FunctionCall(name=expr.name, arguments=args)
            )

        return ExpressionStatement(expression=expr)

    def identifier(self, items: List[Token]) -> Identifier:
        return Identifier(name=str(items[0]))

    def NUMBER(self, token: Token) -> NumberLiteral:
        value = float(token.value)
        if (
            value.is_integer()
            and '.' not in token.value
            and 'e' not in token.value.lower()
        ):
            value = int(value)
        return NumberLiteral(value=value)

    def STRING(self, token: Token) -> StringLiteral:
        # Remove surrounding quotes and unescape doubled quotes
        value = token.value[1:-1].replace('""', '"')
        return StringLiteral(value=value)

    def true_literal(self, items: List) -> BooleanLiteral:
        return BooleanLiteral(value=True)

    def false_literal(self, items: List) -> BooleanLiteral:
        return BooleanLiteral(value=False)

    def nothing_literal(self, items: List) -> NothingLiteral:
        return NothingLiteral()

    def empty_literal(self, items: List) -> EmptyLiteral:
        return EmptyLiteral()

    def null_literal(self, items: List) -> NullLiteral:
        return NullLiteral()

    def comparison_op(self, items: List[Token]) -> ComparisonOp:
        if not items:
            raise ValueError('comparison_op has no items')
        op_token = items[0]
        op_type = op_token.type if isinstance(op_token, Token) else str(op_token)
        op_map = {
            'EQUAL': ComparisonOp.EQ,
            'NOT_EQUAL': ComparisonOp.NE,
            'LESS_THAN': ComparisonOp.LT,
            'GREATER_THAN': ComparisonOp.GT,
            'LESS_EQUAL': ComparisonOp.LE,
            'GREATER_EQUAL': ComparisonOp.GE,
            'IS_KW': ComparisonOp.IS,
        }
        return op_map.get(op_type, ComparisonOp.EQ)

    def comparison_expr(self, items: List) -> ASTNode:
        if len(items) == 1:
            return items[0]
        left, op, right = items
        return ComparisonExpression(left=left, operator=op, right=right)

    def imp_expr(self, items: List) -> ASTNode:
        return self._build_tail_chain(items, BinaryOp.IMP)

    def imp_tail(self, items: List) -> ASTNode:
        # items: [IMP_KW, eqv_expr] - return the expression
        return items[1] if len(items) > 1 else items[0]

    def eqv_expr(self, items: List) -> ASTNode:
        return self._build_tail_chain(items, BinaryOp.EQV)

    def eqv_tail(self, items: List) -> ASTNode:
        # items: [EQV_KW, xor_expr] - return the expression
        return items[1] if len(items) > 1 else items[0]

    def xor_expr(self, items: List) -> ASTNode:
        return self._build_tail_chain(items, BinaryOp.XOR)

    def xor_tail(self, items: List) -> ASTNode:
        # items: [XOR_KW, or_expr] - return the expression
        return items[1] if len(items) > 1 else items[0]

    def or_expr(self, items: List) -> ASTNode:
        return self._build_tail_chain(items, BinaryOp.OR)

    def or_tail(self, items: List) -> ASTNode:
        # items: [OR_KW, and_expr] - return the expression
        return items[1] if len(items) > 1 else items[0]

    def and_expr(self, items: List) -> ASTNode:
        return self._build_tail_chain(items, BinaryOp.AND)

    def and_tail(self, items: List) -> ASTNode:
        # items: [AND_KW, not_expr] - return the expression
        return items[1] if len(items) > 1 else items[0]

    def concat_expr(self, items: List) -> ASTNode:
        return self._build_tail_chain(items, BinaryOp.CONCAT)

    def concat_tail(self, items: List) -> ASTNode:
        # items: [AMPERSAND, add_expr] - return the expression
        return items[1]

    def add_expr(self, items: List) -> ASTNode:
        if len(items) == 1:
            return items[0]
        result = items[0]
        for i in range(1, len(items)):
            tail = items[i]
            if isinstance(tail, tuple):
                op, expr = tail
                result = BinaryExpression(left=result, operator=op, right=expr)
        return result

    def add_plus(self, items: List) -> tuple:
        # items: [PLUS, mul_expr] - return the expression
        return (BinaryOp.ADD, items[1])

    def add_minus(self, items: List) -> tuple:
        # items: [MINUS, mul_expr] - return the expression
        return (BinaryOp.SUB, items[1])

    def mul_expr(self, items: List) -> ASTNode:
        if len(items) == 1:
            return items[0]
        result = items[0]
        for i in range(1, len(items)):
            tail = items[i]
            if isinstance(tail, tuple):
                op, expr = tail
                result = BinaryExpression(left=result, operator=op, right=expr)
        return result

    def mul_star(self, items: List) -> tuple:
        # items: [STAR, intdiv_expr] - return the expression
        return (BinaryOp.MUL, items[1])

    def mul_slash(self, items: List) -> tuple:
        # items: [SLASH, intdiv_expr] - return the expression
        return (BinaryOp.DIV, items[1])

    def intdiv_expr(self, items: List) -> ASTNode:
        return self._build_tail_chain(items, BinaryOp.INTDIV)

    def intdiv_tail(self, items: List) -> ASTNode:
        # items: [BACKSLASH, mod_expr] - return the expression
        return items[1]

    def mod_expr(self, items: List) -> ASTNode:
        return self._build_tail_chain(items, BinaryOp.MOD)

    def mod_tail(self, items: List) -> ASTNode:
        # items: [MOD_KW, pow_expr] - return the expression
        return items[1] if len(items) > 1 else items[0]

    def pow_expr(self, items: List) -> ASTNode:
        return self._build_tail_chain(items, BinaryOp.POW)

    def _build_tail_chain(self, items: List, op: BinaryOp) -> ASTNode:
        if len(items) == 1:
            return items[0]
        result = items[0]
        for i in range(1, len(items)):
            result = BinaryExpression(left=result, operator=op, right=items[i])
        return result

    def unary_expr(self, items: List) -> ASTNode:
        if len(items) == 1:
            return items[0]
        op_token = items[0]
        if isinstance(op_token, Token):
            op_str = op_token.value
            if op_str == '-':
                op = UnaryOp.NEG
            elif op_str == '+':
                op = UnaryOp.POS
            else:
                raise ValueError(f'Unknown unary operator: {op_str}')
            return UnaryExpression(operator=op, operand=items[1])
        return items[0]

    def not_expr(self, items: List) -> ASTNode:
        if len(items) == 1:
            return items[0]
        # Handle "Not" unary expression
        return UnaryExpression(
            operator=UnaryOp.NOT, operand=items[1] if len(items) > 1 else items[0]
        )

    def paren_expr(self, items: List) -> ASTNode:
        return items[0]

    def call_or_access(self, items: List) -> ASTNode:
        """Handle member access, function calls, and array access."""
        if len(items) == 1:
            return items[0]

        result = items[0]
        i = 1
        while i < len(items):
            item = items[i]
            if isinstance(item, Identifier):
                # Member access (e.g., WScript.Echo)
                result = MemberAccess(object=result, member=item.name)
            elif isinstance(item, list):
                # Function/method call or array access with arguments
                if isinstance(result, Identifier):
                    # Could be array access or function call
                    # We create ArrayAccess and let interpreter decide
                    result = ArrayAccess(name=result.name, indices=item)
                elif isinstance(result, MemberAccess):
                    result = MethodCall(
                        object=result.object, method=result.member, arguments=item
                    )
                elif isinstance(result, ArrayAccess):
                    # Chained array access like arr(i)(j) - not common but possible
                    result = ArrayAccess(
                        name=result.name, indices=result.indices + item
                    )
                else:
                    # Method call on an expression
                    result = MethodCall(object=result, method='', arguments=item)
            elif item is None:
                # Empty parentheses - function/method call with no arguments
                if isinstance(result, Identifier):
                    result = FunctionCall(name=result.name, arguments=[])
                elif isinstance(result, MemberAccess):
                    result = MethodCall(
                        object=result.object, method=result.member, arguments=[]
                    )
                else:
                    result = MethodCall(object=result, method='', arguments=[])
            i += 1
        return result

    def atom(self, items: List) -> ASTNode:
        return items[0]

    def arg_list(self, items: List) -> List[ASTNode]:
        return items

    def new_expr(self, items: List) -> NewExpression:
        class_name = (
            str(items[0]) if isinstance(items[0], Identifier) else str(items[0])
        )
        return NewExpression(class_name=class_name)

    # Control Flow transformers
    def block(self, items: List) -> List[ASTNode]:
        """Transform a block of statements."""
        # Filter out None values
        return [item for item in items if item is not None]

    def if_statement(self, items: List) -> IfStatement:
        """Transform if statement."""
        # items: [IF_KW, expression, THEN_KW, block, elseif_clauses..., else_clause?, END_KW, IF_KW]
        # Filter out keyword tokens
        filtered = [item for item in items if not isinstance(item, Token)]

        condition = filtered[0]
        then_body = filtered[1] if len(filtered) > 1 else []
        elseif_clauses = []
        else_clause = None

        for item in filtered[2:]:
            if isinstance(item, ElseIfClause):
                elseif_clauses.append(item)
            elif isinstance(item, ElseClause):
                else_clause = item
            elif isinstance(item, list):
                # This could be the then_body if it wasn't set
                if not then_body:
                    then_body = item

        return IfStatement(
            condition=condition,
            then_body=then_body if isinstance(then_body, list) else [],
            elseif_clauses=elseif_clauses,
            else_clause=else_clause,
        )

    def elseif_clause(self, items: List) -> ElseIfClause:
        """Transform elseif clause."""
        # items: [ELSEIF_KW, expression, THEN_KW, block]
        filtered = [item for item in items if not isinstance(item, Token)]
        condition = filtered[0]
        body = filtered[1] if len(filtered) > 1 else []
        return ElseIfClause(
            condition=condition, body=body if isinstance(body, list) else []
        )

    def else_clause(self, items: List) -> ElseClause:
        """Transform else clause."""
        # items: [ELSE_KW, block]
        filtered = [item for item in items if not isinstance(item, Token)]
        body = filtered[0] if filtered else []
        return ElseClause(body=body if isinstance(body, list) else [])

    def select_case_statement(self, items: List) -> SelectCaseStatement:
        """Transform Select Case statement."""
        # items: [SELECT_KW, CASE_KW, expression, case_clause*, case_else_clause?, END_KW, SELECT_KW]
        filtered = [item for item in items if not isinstance(item, Token)]

        expression = filtered[0] if filtered else None
        case_clauses = []
        case_else_clause = None

        for item in filtered[1:]:
            if isinstance(item, CaseClause):
                case_clauses.append(item)
            elif isinstance(item, CaseElseClause):
                case_else_clause = item

        return SelectCaseStatement(
            expression=expression,
            case_clauses=case_clauses,
            case_else_clause=case_else_clause,
        )

    def case_clause(self, items: List) -> CaseClause:
        """Transform a Case clause."""
        # items: [CASE_KW, case_values (list), block (list)]
        # case_values returns a list of expressions
        # block returns a list of statements

        values = []
        body = []

        # Expression types that indicate this is a value, not a statement
        expression_types = (
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
            MemberAccess,
            FunctionCall,
            MethodCall,
            NewExpression,
            ArrayAccess,
        )

        # Statement types that indicate this is the body
        statement_types = (
            ExpressionStatement,
            DimStatement,
            IfStatement,
            SelectCaseStatement,
            ForStatement,
            ForEachStatement,
            WhileStatement,
            DoLoopStatement,
            ExitStatement,
            SubStatement,
            FunctionStatement,
            CallStatement,
            OnErrorResumeNextStatement,
            OnErrorGoToStatement,
            ReDimStatement,
            EraseStatement,
            AssignmentStatement,
            SetStatement,
        )

        for item in items:
            if isinstance(item, Token):
                continue  # Skip CASE_KW
            elif isinstance(item, list):
                # Check if this is the body or values
                if item and isinstance(item[0], statement_types):
                    # This is the body
                    body = item
                elif item and isinstance(item[0], expression_types):
                    # This is the values list
                    values = item
                elif not item:
                    # Empty list - could be either, skip
                    pass
                else:
                    # Default: if first item looks like a statement, it's body
                    # Otherwise, it's values
                    if item and hasattr(item[0], '__class__'):
                        first_type = type(item[0]).__name__
                        if 'Statement' in first_type:
                            body = item
                        else:
                            values = item
            elif isinstance(item, ASTNode):
                # Single expression value
                values = [item]

        return CaseClause(values=values, body=body)

    def case_values(self, items: List) -> List[ASTNode]:
        """Transform case values list."""
        return items

    def case_else_clause(self, items: List) -> CaseElseClause:
        """Transform Case Else clause."""
        # items: [CASE_KW, ELSE_KW, block]
        filtered = [item for item in items if not isinstance(item, Token)]
        body = filtered[0] if filtered else []
        return CaseElseClause(body=body if isinstance(body, list) else [])

    def for_statement(self, items: List) -> ForStatement:
        """Transform for statement."""
        # items: [FOR_KW, identifier, "=", expression, TO_KW, expression, (STEP_KW, expression)?, block, NEXT_KW]
        filtered = [item for item in items if not isinstance(item, Token)]

        variable = filtered[0]
        if isinstance(variable, Identifier):
            var_name = variable.name
        else:
            var_name = str(variable)

        start = filtered[1]
        end = filtered[2]
        step = None
        body = []

        for item in filtered[3:]:
            if isinstance(item, list):
                body = item
            else:
                step = item

        return ForStatement(
            variable=var_name, start=start, end=end, step=step, body=body
        )

    def for_each_statement(self, items: List) -> ForEachStatement:
        """Transform For Each statement."""
        # items: [FOR_KW, EACH_KW, identifier, IN_KW, expression, block, NEXT_KW]
        filtered = [item for item in items if not isinstance(item, Token)]

        variable = filtered[0]
        if isinstance(variable, Identifier):
            var_name = variable.name
        else:
            var_name = str(variable)

        collection = filtered[1] if len(filtered) > 1 else None
        body = (
            filtered[2] if len(filtered) > 2 and isinstance(filtered[2], list) else []
        )

        return ForEachStatement(variable=var_name, collection=collection, body=body)

    def while_statement(self, items: List) -> WhileStatement:
        """Transform while statement."""
        # items: [WHILE_KW, expression, block, WEND_KW]
        filtered = [item for item in items if not isinstance(item, Token)]
        condition = filtered[0]
        body = filtered[1] if len(filtered) > 1 else []
        return WhileStatement(
            condition=condition, body=body if isinstance(body, list) else []
        )

    def do_while_pretest(self, items: List) -> DoLoopStatement:
        """Transform Do While ... Loop statement."""
        # items: [DO_KW, WHILE_KW, expression, block, LOOP_KW]
        filtered = [item for item in items if not isinstance(item, Token)]
        condition = filtered[0] if filtered else None
        body = filtered[1] if len(filtered) > 1 else []
        return DoLoopStatement(
            pre_condition=LoopCondition(
                condition_type=LoopConditionType.WHILE, condition=condition
            ),
            body=body if isinstance(body, list) else [],
            post_condition=None,
        )

    def do_until_pretest(self, items: List) -> DoLoopStatement:
        """Transform Do Until ... Loop statement."""
        # items: [DO_KW, UNTIL_KW, expression, block, LOOP_KW]
        filtered = [item for item in items if not isinstance(item, Token)]
        condition = filtered[0] if filtered else None
        body = filtered[1] if len(filtered) > 1 else []
        return DoLoopStatement(
            pre_condition=LoopCondition(
                condition_type=LoopConditionType.UNTIL, condition=condition
            ),
            body=body if isinstance(body, list) else [],
            post_condition=None,
        )

    def do_while_posttest(self, items: List) -> DoLoopStatement:
        """Transform Do ... Loop While statement."""
        # items: [DO_KW, block, LOOP_KW, WHILE_KW, expression]
        filtered = [item for item in items if not isinstance(item, Token)]
        body = filtered[0] if filtered else []
        condition = filtered[1] if len(filtered) > 1 else None
        return DoLoopStatement(
            pre_condition=None,
            body=body if isinstance(body, list) else [],
            post_condition=LoopCondition(
                condition_type=LoopConditionType.WHILE, condition=condition
            ),
        )

    def do_until_posttest(self, items: List) -> DoLoopStatement:
        """Transform Do ... Loop Until statement."""
        # items: [DO_KW, block, LOOP_KW, UNTIL_KW, expression]
        filtered = [item for item in items if not isinstance(item, Token)]
        body = filtered[0] if filtered else []
        condition = filtered[1] if len(filtered) > 1 else None
        return DoLoopStatement(
            pre_condition=None,
            body=body if isinstance(body, list) else [],
            post_condition=LoopCondition(
                condition_type=LoopConditionType.UNTIL, condition=condition
            ),
        )

    def do_infinite(self, items: List) -> DoLoopStatement:
        """Transform Do ... Loop statement (infinite loop)."""
        # items: [DO_KW, block, LOOP_KW]
        filtered = [item for item in items if not isinstance(item, Token)]
        body = filtered[0] if filtered else []
        return DoLoopStatement(
            pre_condition=None,
            body=body if isinstance(body, list) else [],
            post_condition=None,
        )

    def exit_statement(self, items: List) -> ExitStatement:
        """Transform exit statement."""
        # items: [EXIT_KW, FOR_KW | DO_KW | SUB_KW | FUNCTION_KW]
        for item in items:
            if isinstance(item, Token):
                if item.type == 'FOR_KW':
                    return ExitStatement(exit_type=ExitType.FOR)
                elif item.type == 'DO_KW':
                    return ExitStatement(exit_type=ExitType.DO)
                elif item.type == 'SUB_KW':
                    return ExitStatement(exit_type=ExitType.SUB)
                elif item.type == 'FUNCTION_KW':
                    return ExitStatement(exit_type=ExitType.FUNCTION)

        # Default to Exit For if we can't determine
        return ExitStatement(exit_type=ExitType.FOR)

    def redim_statement(self, items: List) -> ReDimStatement:
        """Transform ReDim statement."""
        # items: [REDIM_KW, PRESERVE_KW?, redim_var, ("," redim_var)*]
        preserve = False
        arrays = []

        for item in items:
            if isinstance(item, Token):
                if item.type == 'PRESERVE_KW':
                    preserve = True
            elif isinstance(item, tuple):
                arrays.append(item)

        return ReDimStatement(preserve=preserve, arrays=arrays)

    def redim_var(self, items: List) -> tuple:
        """Transform a ReDim variable declaration."""
        # items: [identifier, "(", dim_dimensions, ")"]
        name = ''
        dimensions = []

        for item in items:
            if isinstance(item, Identifier):
                name = item.name
            elif isinstance(item, list):
                dimensions = item

        return (name, dimensions)

    def erase_statement(self, items: List) -> EraseStatement:
        """Transform Erase statement."""
        # items: [ERASE_KW, identifier, ("," identifier)*]
        arrays = []
        for item in items:
            if isinstance(item, Identifier):
                arrays.append(item.name)
        return EraseStatement(arrays=arrays)

    def sub_statement(self, items: List) -> SubStatement:
        """Transform sub statement."""
        # items: [SUB_KW, identifier, "(", param_list?, ")", block, END_KW, SUB_KW]
        filtered = [item for item in items if not isinstance(item, Token)]

        name = ''
        parameters = []
        body = []

        for item in filtered:
            if isinstance(item, Identifier):
                name = item.name
            elif isinstance(item, list):
                if all(isinstance(p, Parameter) for p in item):
                    parameters = item
                else:
                    body = item
            elif isinstance(item, Parameter):
                parameters = [item]

        return SubStatement(name=name, parameters=parameters, body=body)

    def function_statement(self, items: List) -> FunctionStatement:
        """Transform function statement."""
        # items: [FUNCTION_KW, identifier, "(", param_list?, ")", block, END_KW, FUNCTION_KW]
        filtered = [item for item in items if not isinstance(item, Token)]

        name = ''
        parameters = []
        body = []

        for item in filtered:
            if isinstance(item, Identifier):
                name = item.name
            elif isinstance(item, list):
                if all(isinstance(p, Parameter) for p in item):
                    parameters = item
                else:
                    body = item
            elif isinstance(item, Parameter):
                parameters = [item]

        return FunctionStatement(name=name, parameters=parameters, body=body)

    def param_list(self, items: List) -> List[Parameter]:
        """Transform parameter list."""
        return [item for item in items if isinstance(item, Parameter)]

    def param_item(self, items: List) -> Parameter:
        """Transform a single parameter item."""
        # items: [BYREF_KW NAME] | [BYVAL_KW NAME] | [NAME]
        is_byref = True  # Default is ByRef in VBScript
        name = ''

        for item in items:
            if isinstance(item, Token):
                if item.type == 'BYREF_KW':
                    is_byref = True
                elif item.type == 'BYVAL_KW':
                    is_byref = False
                elif item.type == 'NAME':
                    name = str(item)

        return Parameter(name=name, is_byref=is_byref)

    def call_statement(self, items: List) -> CallStatement:
        """Transform call statement."""
        # items: [CALL_KW, identifier, "(" arg_list? ")" | arg_list?]
        filtered = [item for item in items if not isinstance(item, Token)]

        name = ''
        arguments = []

        for item in filtered:
            if isinstance(item, Identifier):
                name = item.name
            elif isinstance(item, list):
                arguments = item

        return CallStatement(name=name, arguments=arguments)

    def on_error_resume_next(self, items: List) -> OnErrorResumeNextStatement:
        """Transform On Error Resume Next statement."""
        return OnErrorResumeNextStatement()

    def on_error_goto(self, items: List) -> OnErrorGoToStatement:
        """Transform On Error GoTo statement."""
        # items: [ON_KW, ERROR_KW, GOTO_KW, NUMBER]
        # Find the NUMBER token
        for item in items:
            if isinstance(item, Token) and item.type == 'NUMBER':
                return OnErrorGoToStatement(label=int(item.value))
            elif isinstance(item, NumberLiteral):
                return OnErrorGoToStatement(label=int(item.value))
        # Default to 0 if not found
        return OnErrorGoToStatement(label=0)


class _VBScriptPostLexer:
    """Suppress newline tokens inside parentheses so that multi-line
    expressions (e.g. function calls spanning lines) parse correctly,
    while still treating top-level newlines as statement separators."""

    always_accept = ('_NL',)

    def process(self, stream: Iterator[Token]) -> Iterator[Token]:
        paren_depth = 0
        for token in stream:
            if token == '(':
                paren_depth += 1
            elif token == ')':
                paren_depth = max(0, paren_depth - 1)

            if token.type == '_NL':
                if paren_depth == 0:
                    yield token
                continue

            yield token


class VBScriptParser:
    """VBScript parser that produces AST from source code."""

    def __init__(self):
        self._lark_parser: Optional[Lark] = None
        self._transformer = VBScriptTransformer()

    @property
    def parser(self) -> Lark:
        if self._lark_parser is None:
            grammar_path = Path(__file__).parent / 'grammar' / 'vbscript.lark'
            with open(grammar_path, 'r') as f:
                grammar = f.read()
            self._lark_parser = Lark(
                grammar, parser='lalr', start='start',
                postlex=_VBScriptPostLexer(),
            )
        return self._lark_parser

    def _preprocess(self, source: str) -> str:
        """Pre-process VBScript source to handle REM comments."""
        lines = source.split('\n')
        processed_lines = []
        for line in lines:
            new_line = self._replace_rem_outside_strings(line)
            processed_lines.append(new_line)
        return '\n'.join(processed_lines)

    @staticmethod
    def _replace_rem_outside_strings(line: str) -> str:
        """Replace REM keywords with ' only when they are outside string literals."""
        result: list[str] = []
        i = 0
        while i < len(line):
            if line[i] == '"':
                # Skip over the entire string literal
                j = i + 1
                while j < len(line) and line[j] != '"':
                    j += 1
                # Include the closing quote (if present)
                result.append(line[i:j + 1])
                i = j + 1
            else:
                m = re.match(r'\b[Rr][Ee][Mm](?=\s|$)', line[i:])
                if m:
                    result.append("'")
                    i += m.end()
                else:
                    result.append(line[i])
                    i += 1
        return ''.join(result)

    def parse(self, source: str) -> Program:
        """Parse VBScript source code and return an AST."""
        source = self._preprocess(source)
        if not source.endswith('\n'):
            source += '\n'
        tree = self.parser.parse(source)
        return self._transformer.transform(tree)


def parse(source: str) -> Program:
    """Convenience function to parse VBScript source code."""
    parser = VBScriptParser()
    return parser.parse(source)
