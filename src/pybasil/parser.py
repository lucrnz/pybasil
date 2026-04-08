"""VBScript Parser using Lark."""

from __future__ import annotations
from pathlib import Path
from typing import List, Optional, Union

from lark import Lark, Transformer, Token, Tree

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
    DimStatement,
    AssignmentStatement,
    SetStatement,
    CallStatement,
    ExpressionStatement,
    IfStatement,
    ElseIfClause,
    ElseClause,
    ForStatement,
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
        variables = []
        for item in items:
            # Skip keyword tokens
            if isinstance(item, Token):
                continue
            if isinstance(item, Identifier):
                variables.append(item.name)
            else:
                variables.append(str(item))
        return DimStatement(variables=variables)

    def assignment_statement(self, items: List) -> AssignmentStatement:
        # items: [LET_KW?, identifier, expression]
        # Skip keyword tokens
        filtered = [item for item in items if not isinstance(item, Token)]
        if len(filtered) >= 2:
            variable, expr = filtered[0], filtered[1]
        else:
            variable, expr = filtered[0], items[-1]
        # Extract variable name from Identifier if needed
        var_name = variable.name if isinstance(variable, Identifier) else str(variable)
        return AssignmentStatement(variable=var_name, expression=expr)

    def set_statement(self, items: List) -> SetStatement:
        # items: [SET_KW, identifier, expression]
        # Skip keyword tokens
        filtered = [item for item in items if not isinstance(item, Token)]
        if len(filtered) >= 2:
            variable, expr = filtered[0], filtered[1]
        else:
            variable, expr = filtered[0], items[-1]
        # Extract variable name from Identifier if needed
        var_name = variable.name if isinstance(variable, Identifier) else str(variable)
        return SetStatement(variable=var_name, expression=expr)

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
            return ExpressionStatement(expression=MethodCall(
                object=expr.object,
                method=expr.member,
                arguments=args
            ))
        
        # If the expression is an Identifier, convert it to a FunctionCall (implicit procedure call)
        if isinstance(expr, Identifier):
            return ExpressionStatement(expression=FunctionCall(
                name=expr.name,
                arguments=args
            ))
        
        return ExpressionStatement(expression=expr)

    def identifier(self, items: List[Token]) -> Identifier:
        return Identifier(name=str(items[0]))

    def NUMBER(self, token: Token) -> NumberLiteral:
        value = float(token.value)
        if value.is_integer() and '.' not in token.value and 'e' not in token.value.lower():
            value = int(value)
        return NumberLiteral(value=value)

    def STRING(self, token: Token) -> StringLiteral:
        # Remove surrounding quotes
        value = token.value[1:-1]
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
            raise ValueError("comparison_op has no items")
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
                raise ValueError(f"Unknown unary operator: {op_str}")
            return UnaryExpression(operator=op, operand=items[1])
        return items[0]

    def not_expr(self, items: List) -> ASTNode:
        if len(items) == 1:
            return items[0]
        # Handle "Not" unary expression
        return UnaryExpression(operator=UnaryOp.NOT, operand=items[1] if len(items) > 1 else items[0])

    def paren_expr(self, items: List) -> ASTNode:
        return items[0]

    def call_or_access(self, items: List) -> ASTNode:
        """Handle member access and function calls."""
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
                # Function/method call with arguments
                if isinstance(result, Identifier):
                    result = FunctionCall(name=result.name, arguments=item)
                elif isinstance(result, MemberAccess):
                    result = MethodCall(object=result.object, method=result.member, arguments=item)
                else:
                    # Method call on an expression
                    result = MethodCall(object=result, method="", arguments=item)
            elif item is None:
                # Empty parentheses - function/method call with no arguments
                if isinstance(result, Identifier):
                    result = FunctionCall(name=result.name, arguments=[])
                elif isinstance(result, MemberAccess):
                    result = MethodCall(object=result.object, method=result.member, arguments=[])
                else:
                    result = MethodCall(object=result, method="", arguments=[])
            i += 1
        return result

    def _build_call_or_access(self, atom: ASTNode, identifiers: List[Identifier], args: Optional[List[ASTNode]]) -> ASTNode:
        """Build a call or access chain from components."""
        result = atom
        
        # Process member accesses
        for identifier in identifiers:
            # For the first member access, use the atom directly as object
            # For subsequent ones, use the previous result
            result = MemberAccess(object=result, member=identifier.name)
        
        # Process function/method call
        if args is not None:
            if isinstance(result, Identifier):
                result = FunctionCall(name=result.name, arguments=args)
            elif isinstance(result, MemberAccess):
                result = MethodCall(object=result.object, method=result.member, arguments=args)
            else:
                result = MethodCall(object=result, method="", arguments=args)
        
        return result

    def call_or_access(self, items: List) -> ASTNode:
        """Handle member access and function calls (fallback for old grammar)."""
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
                # Function/method call with arguments
                if isinstance(result, Identifier):
                    result = FunctionCall(name=result.name, arguments=item)
                elif isinstance(result, MemberAccess):
                    result = MethodCall(object=result.object, method=result.member, arguments=item)
                else:
                    # Method call on an expression
                    result = MethodCall(object=result, method="", arguments=item)
            elif item is None:
                # Empty parentheses - function/method call with no arguments
                if isinstance(result, Identifier):
                    result = FunctionCall(name=result.name, arguments=[])
                elif isinstance(result, MemberAccess):
                    result = MethodCall(object=result.object, method=result.member, arguments=[])
                else:
                    result = MethodCall(object=result, method="", arguments=[])
            i += 1
        return result

    def atom(self, items: List) -> ASTNode:
        return items[0]

    def arg_list(self, items: List) -> List[ASTNode]:
        return items

    def new_expr(self, items: List) -> NewExpression:
        class_name = str(items[0]) if isinstance(items[0], Identifier) else str(items[0])
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
            else_clause=else_clause
        )

    def elseif_clause(self, items: List) -> ElseIfClause:
        """Transform elseif clause."""
        # items: [ELSEIF_KW, expression, THEN_KW, block]
        filtered = [item for item in items if not isinstance(item, Token)]
        condition = filtered[0]
        body = filtered[1] if len(filtered) > 1 else []
        return ElseIfClause(condition=condition, body=body if isinstance(body, list) else [])

    def else_clause(self, items: List) -> ElseClause:
        """Transform else clause."""
        # items: [ELSE_KW, block]
        filtered = [item for item in items if not isinstance(item, Token)]
        body = filtered[0] if filtered else []
        return ElseClause(body=body if isinstance(body, list) else [])

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
            variable=var_name,
            start=start,
            end=end,
            step=step,
            body=body
        )

    def while_statement(self, items: List) -> WhileStatement:
        """Transform while statement."""
        # items: [WHILE_KW, expression, block, WEND_KW]
        filtered = [item for item in items if not isinstance(item, Token)]
        condition = filtered[0]
        body = filtered[1] if len(filtered) > 1 else []
        return WhileStatement(condition=condition, body=body if isinstance(body, list) else [])

    def do_while_pretest(self, items: List) -> DoLoopStatement:
        """Transform Do While ... Loop statement."""
        # items: [DO_KW, WHILE_KW, expression, block, LOOP_KW]
        filtered = [item for item in items if not isinstance(item, Token)]
        condition = filtered[0] if filtered else None
        body = filtered[1] if len(filtered) > 1 else []
        return DoLoopStatement(
            pre_condition=LoopCondition(condition_type=LoopConditionType.WHILE, condition=condition),
            body=body if isinstance(body, list) else [],
            post_condition=None
        )

    def do_until_pretest(self, items: List) -> DoLoopStatement:
        """Transform Do Until ... Loop statement."""
        # items: [DO_KW, UNTIL_KW, expression, block, LOOP_KW]
        filtered = [item for item in items if not isinstance(item, Token)]
        condition = filtered[0] if filtered else None
        body = filtered[1] if len(filtered) > 1 else []
        return DoLoopStatement(
            pre_condition=LoopCondition(condition_type=LoopConditionType.UNTIL, condition=condition),
            body=body if isinstance(body, list) else [],
            post_condition=None
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
            post_condition=LoopCondition(condition_type=LoopConditionType.WHILE, condition=condition)
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
            post_condition=LoopCondition(condition_type=LoopConditionType.UNTIL, condition=condition)
        )

    def do_infinite(self, items: List) -> DoLoopStatement:
        """Transform Do ... Loop statement (infinite loop)."""
        # items: [DO_KW, block, LOOP_KW]
        filtered = [item for item in items if not isinstance(item, Token)]
        body = filtered[0] if filtered else []
        return DoLoopStatement(
            pre_condition=None,
            body=body if isinstance(body, list) else [],
            post_condition=None
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

    # Procedure transformers
    def block_statement(self, items: List) -> ASTNode:
        """Transform a block statement (same as statement but for inside procedures)."""
        items = [i for i in items if i is not None]
        if not items:
            return None
        return items[0]

    def sub_statement(self, items: List) -> SubStatement:
        """Transform sub statement."""
        # items: [SUB_KW, identifier, "(", param_list?, ")", block, END_KW, SUB_KW]
        filtered = [item for item in items if not isinstance(item, Token)]
        
        name = ""
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
        
        name = ""
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
        name = ""
        
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
        
        name = ""
        arguments = []
        
        for item in filtered:
            if isinstance(item, Identifier):
                name = item.name
            elif isinstance(item, list):
                arguments = item
        
        return CallStatement(name=name, arguments=arguments)


import re

class VBScriptParser:
    """VBScript parser that produces AST from source code."""

    def __init__(self):
        self._lark_parser: Optional[Lark] = None
        self._transformer = VBScriptTransformer()

    @property
    def parser(self) -> Lark:
        if self._lark_parser is None:
            grammar_path = Path(__file__).parent / "grammar" / "vbscript.lark"
            with open(grammar_path, 'r') as f:
                grammar = f.read()
            self._lark_parser = Lark(grammar, parser='lalr', start='start')
        return self._lark_parser

    def _preprocess(self, source: str) -> str:
        """Pre-process VBScript source to handle REM comments."""
        # Convert REM comments to single-quote comments
        # REM must be followed by whitespace or be at end of line
        lines = source.split('\n')
        processed_lines = []
        for line in lines:
            # Find REM keyword (case-insensitive) followed by whitespace
            # and replace with single quote comment
            new_line = re.sub(r'\b[Rr][Ee][Mm](?=\s|$)', "'", line)
            processed_lines.append(new_line)
        return '\n'.join(processed_lines)

    def parse(self, source: str) -> Program:
        """Parse VBScript source code and return an AST."""
        source = self._preprocess(source)
        tree = self.parser.parse(source)
        return self._transformer.transform(tree)


def parse(source: str) -> Program:
    """Convenience function to parse VBScript source code."""
    parser = VBScriptParser()
    return parser.parse(source)

