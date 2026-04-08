"""VBScript Tree-Walking Interpreter."""

from __future__ import annotations
import math
import sys
from typing import Any, Callable, Dict, List, Optional
from dataclasses import dataclass

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


class VBScriptError(Exception):
    """Base exception for VBScript runtime errors."""
    pass


class VBScriptObject:
    """Base class for VBScript objects."""
    pass


@dataclass
class VBScriptNothing:
    """Represents VBScript Nothing value."""
    pass


@dataclass
class VBScriptEmpty:
    """Represents VBScript Empty value."""
    pass


@dataclass
class VBScriptNull:
    """Represents VBScript Null value."""
    pass


# Singleton instances
NOTHING = VBScriptNothing()
EMPTY = VBScriptEmpty()
NULL = VBScriptNull()


class ExitLoopException(Exception):
    """Exception raised when Exit For or Exit Do is encountered."""
    def __init__(self, exit_type: ExitType):
        self.exit_type = exit_type


class ExitProcedureException(Exception):
    """Exception raised when Exit Sub or Exit Function is encountered."""
    def __init__(self, exit_type: ExitType, return_value: Any = EMPTY):
        self.exit_type = exit_type
        self.return_value = return_value


class UndefinedVariableError(VBScriptError):
    """Raised when accessing an undefined variable."""
    pass


@dataclass
class Procedure:
    """Represents a user-defined procedure (Sub or Function)."""
    name: str
    parameters: List[Parameter]
    body: List[ASTNode]
    is_function: bool  # True for Function, False for Sub


class Environment:
    """Variable environment for VBScript execution."""

    def __init__(self, parent: Optional['Environment'] = None):
        self._variables: Dict[str, Any] = {}
        self._parent = parent

    def define(self, name: str, value: Any = EMPTY) -> None:
        """Define a new variable."""
        self._variables[name.lower()] = value

    def get(self, name: str) -> Any:
        """Get a variable value."""
        key = name.lower()
        if key in self._variables:
            return self._variables[key]
        if self._parent:
            return self._parent.get(name)
        # VBScript allows implicit variable creation
        # Return Empty for undefined variables
        return EMPTY

    def set(self, name: str, value: Any) -> None:
        """Set a variable value, checking parent scope if not in local scope."""
        key = name.lower()
        # If variable exists in local scope, set it there
        if key in self._variables:
            self._variables[key] = value
        # If variable exists in parent scope, set it there
        elif self._parent and self._parent.exists(key):
            self._parent.set(name, value)
        # Otherwise, create new variable in local scope (VBScript implicit declaration)
        else:
            self._variables[key] = value

    def exists(self, name: str) -> bool:
        """Check if a variable exists."""
        key = name.lower()
        if key in self._variables:
            return True
        if self._parent:
            return self._parent.exists(name)
        return False


class WScriptObject:
    """Simulates the WScript object for VBScript."""

    def __init__(self, output_stream=None):
        self._output = output_stream or sys.stdout

    def Echo(self, *args: Any) -> None:
        """WScript.Echo implementation - prints to stdout."""
        output_parts = []
        for arg in args:
            output_parts.append(self._format_value(arg))
        output = " ".join(output_parts)
        print(output, file=self._output)

    def _format_value(self, value: Any) -> str:
        """Format a value for output."""
        if value is True:
            return "True"
        elif value is False:
            return "False"
        elif isinstance(value, VBScriptNothing):
            return "Nothing"
        elif isinstance(value, VBScriptEmpty):
            return ""
        elif isinstance(value, VBScriptNull):
            return "Null"
        elif value is None:
            return "Nothing"
        elif isinstance(value, float):
            if value.is_integer():
                return str(int(value))
            return str(value)
        else:
            return str(value)

    def Quit(self, exit_code: int = 0) -> None:
        """WScript.Quit implementation."""
        sys.exit(exit_code)


class Interpreter:
    """Tree-walking interpreter for VBScript AST."""

    def __init__(self, output_stream=None):
        self._environment = Environment()
        self._output_stream = output_stream
        self._procedures: Dict[str, Procedure] = {}  # User-defined procedures
        self._setup_builtins()

    def _setup_builtins(self) -> None:
        """Set up built-in objects and functions."""
        # Create WScript object
        wscript = WScriptObject(self._output_stream)
        self._environment.define("WScript", wscript)

        # Built-in functions
        self._builtins: Dict[str, Callable] = {
            'msgbox': self._builtin_msgbox,
            'inputbox': self._builtin_inputbox,
            'len': self._builtin_len,
            'left': self._builtin_left,
            'right': self._builtin_right,
            'mid': self._builtin_mid,
            'trim': self._builtin_trim,
            'ltrim': self._builtin_ltrim,
            'rtrim': self._builtin_rtrim,
            'ucase': self._builtin_ucase,
            'lcase': self._builtin_lcase,
            'instr': self._builtin_instr,
            'replace': self._builtin_replace,
            'split': self._builtin_split,
            'join': self._builtin_join,
            'cstr': self._builtin_cstr,
            'cint': self._builtin_cint,
            'clng': self._builtin_clng,
            'cdbl': self._builtin_cdbl,
            'cbool': self._builtin_cbool,
            'cdate': self._builtin_cdate,
            'isnumeric': self._builtin_isnumeric,
            'isarray': self._builtin_isarray,
            'isdate': self._builtin_isdate,
            'isempty': self._builtin_isempty,
            'isnull': self._builtin_isnull,
            'isnumeric': self._builtin_isnumeric,
            'isobject': self._builtin_isobject,
            'typename': self._builtin_typename,
            'vartype': self._builtin_vartype,
            'abs': self._builtin_abs,
            'sqr': self._builtin_sqr,
            'int': self._builtin_int,
            'fix': self._builtin_fix,
            'round': self._builtin_round,
            'rnd': self._builtin_rnd,
            'randomize': self._builtin_randomize,
            'createobject': self._builtin_createobject,
            'getobject': self._builtin_getobject,
        }

    def interpret(self, program: Program) -> Any:
        """Interpret a VBScript program."""
        result = None
        for statement in program.statements:
            result = self._execute(statement)
        return result

    def _execute(self, node: ASTNode) -> Any:
        """Execute a statement node."""
        method_name = f'_execute_{type(node).__name__}'
        method = getattr(self, method_name, self._execute_default)
        return method(node)

    def _execute_default(self, node: ASTNode) -> Any:
        """Default execution handler."""
        raise VBScriptError(f"Unknown statement type: {type(node).__name__}")

    def _execute_DimStatement(self, node: DimStatement) -> None:
        """Execute a Dim statement."""
        for var_name in node.variables:
            self._environment.define(var_name, EMPTY)

    def _execute_AssignmentStatement(self, node: AssignmentStatement) -> None:
        """Execute an assignment statement."""
        value = self._evaluate(node.expression)
        self._environment.set(node.variable, value)

    def _execute_SetStatement(self, node: SetStatement) -> None:
        """Execute a Set statement."""
        value = self._evaluate(node.expression)
        self._environment.set(node.variable, value)

    def _execute_CallStatement(self, node: CallStatement) -> Any:
        """Execute a Call statement."""
        return self._call_procedure(node.name, node.arguments)

    def _execute_SubStatement(self, node: SubStatement) -> None:
        """Register a Sub procedure."""
        proc = Procedure(
            name=node.name.lower(),
            parameters=node.parameters,
            body=node.body,
            is_function=False
        )
        self._procedures[node.name.lower()] = proc

    def _execute_FunctionStatement(self, node: FunctionStatement) -> None:
        """Register a Function procedure."""
        proc = Procedure(
            name=node.name.lower(),
            parameters=node.parameters,
            body=node.body,
            is_function=True
        )
        self._procedures[node.name.lower()] = proc

    def _call_procedure(self, name: str, arguments: List[ASTNode]) -> Any:
        """Call a user-defined procedure or built-in function."""
        proc_name = name.lower()
        
        # Check for user-defined procedure
        if proc_name in self._procedures:
            proc = self._procedures[proc_name]
            return self._execute_procedure(proc, arguments)
        
        # Check for built-in function
        if proc_name in self._builtins:
            args = [self._evaluate(arg) for arg in arguments]
            return self._builtins[proc_name](*args)
        
        raise VBScriptError(f"Unknown procedure: {name}")

    def _execute_procedure(self, proc: Procedure, arguments: List[ASTNode]) -> Any:
        """Execute a user-defined procedure with proper scoping."""
        # Create a new environment for the procedure
        old_env = self._environment
        proc_env = Environment(parent=old_env)
        self._environment = proc_env
        
        try:
            # Bind parameters
            arg_values = [self._evaluate(arg) for arg in arguments]
            
            # Store references for ByRef parameters
            byref_bindings: Dict[str, tuple] = {}  # name -> (old_env, var_name)
            
            for i, param in enumerate(proc.parameters):
                if i < len(arg_values):
                    if param.is_byref:
                        # For ByRef, we need to store a reference to the original variable
                        # Check if the argument is a simple variable
                        if i < len(arguments) and isinstance(arguments[i], Identifier):
                            var_name = arguments[i].name
                            # Store binding info for later
                            byref_bindings[param.name.lower()] = (old_env, var_name)
                            # Copy current value to local scope
                            proc_env.define(param.name, old_env.get(var_name))
                        else:
                            # If argument is not a variable, treat as ByVal
                            proc_env.define(param.name, arg_values[i])
                    else:
                        # ByVal - copy the value
                        proc_env.define(param.name, arg_values[i])
                else:
                    # No argument provided, use Empty
                    proc_env.define(param.name, EMPTY)
            
            # For functions, initialize return value variable
            if proc.is_function:
                proc_env.define(proc.name, EMPTY)
            
            # Execute the procedure body
            try:
                for stmt in proc.body:
                    self._execute(stmt)
            except ExitProcedureException as e:
                # Exit Sub/Function was called
                if proc.is_function:
                    return proc_env.get(proc.name)
                return e.return_value
            
            # Handle Exit For/Do that propagated up
            except ExitLoopException:
                raise VBScriptError("Exit For/Do not valid outside of loops")
            
            # For functions, return the function's return value
            if proc.is_function:
                return proc_env.get(proc.name)
            
            return EMPTY
            
        finally:
            # Copy ByRef values back to the original scope
            for param_name, (orig_env, orig_var) in byref_bindings.items():
                orig_env.set(orig_var, proc_env.get(param_name))
            
            # Restore the original environment
            self._environment = old_env

    def _execute_ExpressionStatement(self, node: ExpressionStatement) -> Any:
        """Execute an expression statement."""
        # Check if this is a procedure call (FunctionCall to a Sub or Function)
        if isinstance(node.expression, FunctionCall):
            name = node.expression.name.lower()
            if name in self._procedures:
                return self._call_procedure(node.expression.name, node.expression.arguments)
        
        # Check if this is a procedure call (identifier without args)
        if isinstance(node.expression, Identifier):
            name = node.expression.name.lower()
            if name in self._procedures:
                return self._call_procedure(node.expression.name, [])
        
        # Check if this is a method call that should invoke a procedure
        if isinstance(node.expression, MethodCall):
            return self._evaluate(node.expression)
        
        return self._evaluate(node.expression)

    def _execute_IfStatement(self, node: IfStatement) -> Any:
        """Execute an If statement."""
        condition = self._evaluate(node.condition)
        if self._to_boolean(condition):
            for stmt in node.then_body:
                self._execute(stmt)
            return None
        
        # Check ElseIf clauses
        for elseif in node.elseif_clauses:
            condition = self._evaluate(elseif.condition)
            if self._to_boolean(condition):
                for stmt in elseif.body:
                    self._execute(stmt)
                return None
        
        # Execute Else clause if present
        if node.else_clause:
            for stmt in node.else_clause.body:
                self._execute(stmt)
        
        return None

    def _execute_ForStatement(self, node: ForStatement) -> Any:
        """Execute a For...Next statement."""
        start_val = self._to_number(self._evaluate(node.start))
        end_val = self._to_number(self._evaluate(node.end))
        
        # Determine step value
        if node.step:
            step_val = self._to_number(self._evaluate(node.step))
        else:
            # Default step is 1, or -1 if start > end
            step_val = 1 if start_val <= end_val else 1
        
        # Set the loop variable to start value
        self._environment.set(node.variable, start_val)
        
        try:
            while True:
                current_val = self._to_number(self._environment.get(node.variable))
                
                # Check loop condition
                if step_val > 0:
                    if current_val > end_val:
                        break
                else:
                    if current_val < end_val:
                        break
                
                # Execute body
                try:
                    for stmt in node.body:
                        self._execute(stmt)
                except ExitLoopException as e:
                    if e.exit_type == ExitType.FOR:
                        return None
                    raise
                
                # Increment loop variable
                self._environment.set(node.variable, current_val + step_val)
        except ExitLoopException as e:
            if e.exit_type == ExitType.FOR:
                return None
            raise
        
        return None

    def _execute_WhileStatement(self, node: WhileStatement) -> Any:
        """Execute a While...Wend statement."""
        try:
            while self._to_boolean(self._evaluate(node.condition)):
                try:
                    for stmt in node.body:
                        self._execute(stmt)
                except ExitLoopException as e:
                    if e.exit_type == ExitType.DO:
                        raise VBScriptError("Exit Do not valid in While loop")
                    return None
        except ExitLoopException as e:
            raise VBScriptError("Exit Do not valid in While loop")
        
        return None

    def _execute_DoLoopStatement(self, node: DoLoopStatement) -> Any:
        """Execute a Do...Loop statement."""
        try:
            while True:
                # Check pre-condition (Do While/Until)
                if node.pre_condition:
                    cond_result = self._to_boolean(self._evaluate(node.pre_condition.condition))
                    if node.pre_condition.condition_type == LoopConditionType.WHILE:
                        if not cond_result:
                            break
                    else:  # UNTIL
                        if cond_result:
                            break
                
                # Execute body
                try:
                    for stmt in node.body:
                        self._execute(stmt)
                except ExitLoopException as e:
                    if e.exit_type == ExitType.DO:
                        return None
                    raise
                
                # Check post-condition (Loop While/Until)
                if node.post_condition:
                    cond_result = self._to_boolean(self._evaluate(node.post_condition.condition))
                    if node.post_condition.condition_type == LoopConditionType.WHILE:
                        if not cond_result:
                            break
                    else:  # UNTIL
                        if cond_result:
                            break
                elif not node.pre_condition:
                    # No conditions - infinite loop protection
                    # In real VBScript this would loop forever
                    # We'll add a safety limit to prevent hanging
                    pass
        except ExitLoopException as e:
            if e.exit_type == ExitType.DO:
                return None
            raise
        
        return None

    def _execute_ExitStatement(self, node: ExitStatement) -> None:
        """Execute an Exit statement."""
        if node.exit_type in (ExitType.SUB, ExitType.FUNCTION):
            raise ExitProcedureException(node.exit_type)
        else:
            raise ExitLoopException(node.exit_type)

    def _evaluate(self, node: ASTNode) -> Any:
        """Evaluate an expression node."""
        method_name = f'_evaluate_{type(node).__name__}'
        method = getattr(self, method_name, self._evaluate_default)
        return method(node)

    def _evaluate_default(self, node: ASTNode) -> Any:
        """Default evaluation handler."""
        raise VBScriptError(f"Unknown expression type: {type(node).__name__}")

    def _evaluate_NumberLiteral(self, node: NumberLiteral) -> Any:
        """Evaluate a number literal."""
        return node.value

    def _evaluate_StringLiteral(self, node: StringLiteral) -> str:
        """Evaluate a string literal."""
        return node.value

    def _evaluate_BooleanLiteral(self, node: BooleanLiteral) -> bool:
        """Evaluate a boolean literal."""
        return node.value

    def _evaluate_NothingLiteral(self, node: NothingLiteral) -> VBScriptNothing:
        """Evaluate a Nothing literal."""
        return NOTHING

    def _evaluate_EmptyLiteral(self, node: EmptyLiteral) -> VBScriptEmpty:
        """Evaluate an Empty literal."""
        return EMPTY

    def _evaluate_NullLiteral(self, node: NullLiteral) -> VBScriptNull:
        """Evaluate a Null literal."""
        return NULL

    def _evaluate_Identifier(self, node: Identifier) -> Any:
        """Evaluate an identifier."""
        name = node.name.lower()
        
        # Check if this is a function call (function name without parentheses)
        if name in self._procedures:
            proc = self._procedures[name]
            if proc.is_function:
                return self._execute_procedure(proc, [])
        
        return self._environment.get(node.name)

    def _evaluate_BinaryExpression(self, node: BinaryExpression) -> Any:
        """Evaluate a binary expression."""
        left = self._evaluate(node.left)
        right = self._evaluate(node.right)
        return self._apply_binary_op(node.operator, left, right)

    def _evaluate_UnaryExpression(self, node: UnaryExpression) -> Any:
        """Evaluate a unary expression."""
        operand = self._evaluate(node.operand)
        return self._apply_unary_op(node.operator, operand)

    def _evaluate_ComparisonExpression(self, node: ComparisonExpression) -> bool:
        """Evaluate a comparison expression."""
        left = self._evaluate(node.left)
        right = self._evaluate(node.right)
        return self._apply_comparison_op(node.operator, left, right)

    def _evaluate_MemberAccess(self, node: MemberAccess) -> Any:
        """Evaluate member access (e.g., WScript.Echo)."""
        obj = self._evaluate(node.object)
        
        if obj is None or isinstance(obj, VBScriptNothing):
            raise VBScriptError(f"Object required: {node.member}")
        
        # Handle WScript object
        if isinstance(obj, WScriptObject):
            attr_name = node.member.lower()
            if attr_name == 'echo':
                return getattr(obj, 'Echo')
            elif attr_name == 'quit':
                return getattr(obj, 'Quit')
            else:
                raise VBScriptError(f"Unknown member: WScript.{node.member}")
        
        # Handle dictionary-like objects
        if isinstance(obj, dict):
            return obj.get(node.member.lower(), EMPTY)
        
        # Handle objects with attributes
        if hasattr(obj, node.member):
            return getattr(obj, node.member)
        
        raise VBScriptError(f"Object doesn't support this property or method: {node.member}")

    def _evaluate_FunctionCall(self, node: FunctionCall) -> Any:
        """Evaluate a function call."""
        func_name = node.name.lower()
        
        # Check for user-defined procedures first
        if func_name in self._procedures:
            proc = self._procedures[func_name]
            if not proc.is_function:
                raise VBScriptError(f"Cannot call Sub '{node.name}' as a function")
            return self._execute_procedure(proc, node.arguments)
        
        # Check built-in functions
        if func_name in self._builtins:
            args = [self._evaluate(arg) for arg in node.arguments]
            return self._builtins[func_name](*args)
        
        raise VBScriptError(f"Unknown function: {node.name}")

    def _evaluate_MethodCall(self, node: MethodCall) -> Any:
        """Evaluate a method call."""
        obj = self._evaluate(node.object)
        method = node.method
        args = [self._evaluate(arg) for arg in node.arguments]
        
        if callable(obj):
            return obj(*args)
        
        # Handle WScript object with case-insensitive method lookup
        if isinstance(obj, WScriptObject):
            method_lower = method.lower()
            if method_lower == 'echo':
                return obj.Echo(*args)
            elif method_lower == 'quit':
                return obj.Quit(*args)
            else:
                raise VBScriptError(f"Object doesn't support this property or method: {method}")
        
        # Try case-insensitive attribute lookup for other objects
        method_lower = method.lower()
        for attr_name in dir(obj):
            if attr_name.lower() == method_lower:
                func = getattr(obj, attr_name)
                if callable(func):
                    return func(*args)
        
        # Fall back to exact name
        if hasattr(obj, method):
            func = getattr(obj, method)
            if callable(func):
                return func(*args)
        
        raise VBScriptError(f"Object doesn't support this property or method: {method}")

    def _evaluate_NewExpression(self, node: NewExpression) -> Any:
        """Evaluate a New expression."""
        raise VBScriptError(f"CreateObject should be used instead of New for: {node.class_name}")

    def _apply_binary_op(self, op: BinaryOp, left: Any, right: Any) -> Any:
        """Apply a binary operator."""
        # Handle Empty values
        if isinstance(left, VBScriptEmpty):
            left = 0 if isinstance(right, (int, float)) else ""
        if isinstance(right, VBScriptEmpty):
            right = 0 if isinstance(left, (int, float)) else ""
        
        # Handle Null propagation
        if isinstance(left, VBScriptNull) or isinstance(right, VBScriptNull):
            if op in (BinaryOp.AND, BinaryOp.OR):
                pass  # Special handling for logical operators
            else:
                return NULL
        
        if op == BinaryOp.ADD:
            return self._add(left, right)
        elif op == BinaryOp.SUB:
            return self._to_number(left) - self._to_number(right)
        elif op == BinaryOp.MUL:
            return self._to_number(left) * self._to_number(right)
        elif op == BinaryOp.DIV:
            right_num = self._to_number(right)
            if right_num == 0:
                raise VBScriptError("Division by zero")
            return self._to_number(left) / right_num
        elif op == BinaryOp.INTDIV:
            right_num = self._to_number(right)
            if right_num == 0:
                raise VBScriptError("Division by zero")
            return int(self._to_number(left) // right_num)
        elif op == BinaryOp.MOD:
            right_num = self._to_number(right)
            if right_num == 0:
                raise VBScriptError("Division by zero")
            return self._to_number(left) % right_num
        elif op == BinaryOp.POW:
            return self._to_number(left) ** self._to_number(right)
        elif op == BinaryOp.CONCAT:
            return self._to_string(left) + self._to_string(right)
        elif op == BinaryOp.AND:
            return self._logical_and(left, right)
        elif op == BinaryOp.OR:
            return self._logical_or(left, right)
        elif op == BinaryOp.XOR:
            return bool(self._to_number(left) ^ self._to_number(right))
        elif op == BinaryOp.EQV:
            return not (bool(self._to_number(left)) ^ bool(self._to_number(right)))
        elif op == BinaryOp.IMP:
            return (not bool(self._to_number(left))) or bool(self._to_number(right))
        else:
            raise VBScriptError(f"Unknown binary operator: {op}")

    def _apply_unary_op(self, op: UnaryOp, operand: Any) -> Any:
        """Apply a unary operator."""
        if op == UnaryOp.NEG:
            return -self._to_number(operand)
        elif op == UnaryOp.POS:
            return self._to_number(operand)
        elif op == UnaryOp.NOT:
            return not self._to_boolean(operand)
        else:
            raise VBScriptError(f"Unknown unary operator: {op}")

    def _apply_comparison_op(self, op: ComparisonOp, left: Any, right: Any) -> bool:
        """Apply a comparison operator."""
        # Handle Empty values
        if isinstance(left, VBScriptEmpty):
            left = 0 if isinstance(right, (int, float)) else ""
        if isinstance(right, VBScriptEmpty):
            right = 0 if isinstance(left, (int, float)) else ""
        
        # Handle Nothing comparison
        if op == ComparisonOp.IS:
            return left is right or (isinstance(left, VBScriptNothing) and isinstance(right, VBScriptNothing))
        
        # Handle Null comparisons
        if isinstance(left, VBScriptNull) or isinstance(right, VBScriptNull):
            return False  # Null comparisons always return False in VBScript
        
        # Type coercion for comparison
        if isinstance(left, str) or isinstance(right, str):
            left = self._to_string(left)
            right = self._to_string(right)
        elif isinstance(left, bool) or isinstance(right, bool):
            left = self._to_boolean(left)
            right = self._to_boolean(right)
        else:
            try:
                left = self._to_number(left)
                right = self._to_number(right)
            except (ValueError, TypeError):
                left = self._to_string(left)
                right = self._to_string(right)
        
        if op == ComparisonOp.EQ:
            return left == right
        elif op == ComparisonOp.NE:
            return left != right
        elif op == ComparisonOp.LT:
            return left < right
        elif op == ComparisonOp.GT:
            return left > right
        elif op == ComparisonOp.LE:
            return left <= right
        elif op == ComparisonOp.GE:
            return left >= right
        else:
            raise VBScriptError(f"Unknown comparison operator: {op}")

    def _add(self, left: Any, right: Any) -> Any:
        """Handle addition with type coercion."""
        # If either is a string, concatenate
        if isinstance(left, str) or isinstance(right, str):
            return self._to_string(left) + self._to_string(right)
        # Otherwise, numeric addition
        return self._to_number(left) + self._to_number(right)

    def _logical_and(self, left: Any, right: Any) -> bool:
        """Logical AND with VBScript semantics."""
        # VBScript uses bitwise AND for numbers
        if isinstance(left, (int, float)) and isinstance(right, (int, float)):
            return bool(int(left) & int(right))
        return self._to_boolean(left) and self._to_boolean(right)

    def _logical_or(self, left: Any, right: Any) -> bool:
        """Logical OR with VBScript semantics."""
        # VBScript uses bitwise OR for numbers
        if isinstance(left, (int, float)) and isinstance(right, (int, float)):
            return bool(int(left) | int(right))
        return self._to_boolean(left) or self._to_boolean(right)

    def _to_number(self, value: Any) -> float:
        """Convert a value to a number."""
        if isinstance(value, VBScriptEmpty):
            return 0
        if isinstance(value, VBScriptNull):
            raise VBScriptError("Type mismatch: cannot convert Null to number")
        if isinstance(value, VBScriptNothing):
            raise VBScriptError("Type mismatch: cannot convert Nothing to number")
        if isinstance(value, bool):
            # In VBScript, True is -1 and False is 0
            return -1 if value else 0
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            if value == "":
                return 0
            try:
                return float(value)
            except ValueError:
                raise VBScriptError(f"Type mismatch: cannot convert '{value}' to number")
        raise VBScriptError(f"Type mismatch: cannot convert to number")

    def _to_string(self, value: Any) -> str:
        """Convert a value to a string."""
        if isinstance(value, VBScriptEmpty):
            return ""
        if isinstance(value, VBScriptNull):
            return "Null"
        if isinstance(value, VBScriptNothing):
            return "Nothing"
        if isinstance(value, bool):
            return "True" if value else "False"
        if isinstance(value, float):
            if value.is_integer():
                return str(int(value))
            return str(value)
        return str(value)

    def _to_boolean(self, value: Any) -> bool:
        """Convert a value to a boolean."""
        if isinstance(value, VBScriptEmpty):
            return False
        if isinstance(value, VBScriptNull):
            return False
        if isinstance(value, VBScriptNothing):
            return False
        if isinstance(value, bool):
            return value
        if isinstance(value, (int, float)):
            return value != 0
        if isinstance(value, str):
            if value == "":
                return False
            if value.lower() == "true":
                return True
            if value.lower() == "false":
                return False
            return True
        return bool(value)

    # Built-in functions
    def _builtin_msgbox(self, *args) -> int:
        """MsgBox function (simplified)."""
        if args:
            print(self._to_string(args[0]))
        return 1  # vbOK

    def _builtin_inputbox(self, prompt: str, title: str = "", default: str = "") -> str:
        """InputBox function (simplified)."""
        return default

    def _builtin_len(self, value: Any) -> int:
        """Len function."""
        if isinstance(value, str):
            return len(value)
        raise VBScriptError("Type mismatch: Len requires a string")

    def _builtin_left(self, string: str, length: int) -> str:
        """Left function."""
        return self._to_string(string)[:int(length)]

    def _builtin_right(self, string: str, length: int) -> str:
        """Right function."""
        s = self._to_string(string)
        n = int(length)
        return s[-n:] if n > 0 else ""

    def _builtin_mid(self, string: str, start: int, length: int = None) -> str:
        """Mid function."""
        s = self._to_string(string)
        start_idx = int(start) - 1  # VBScript is 1-indexed
        if length is None:
            return s[start_idx:]
        return s[start_idx:start_idx + int(length)]

    def _builtin_trim(self, string: str) -> str:
        """Trim function."""
        return self._to_string(string).strip()

    def _builtin_ltrim(self, string: str) -> str:
        """LTrim function."""
        return self._to_string(string).lstrip()

    def _builtin_rtrim(self, string: str) -> str:
        """RTrim function."""
        return self._to_string(string).rstrip()

    def _builtin_ucase(self, string: str) -> str:
        """UCase function."""
        return self._to_string(string).upper()

    def _builtin_lcase(self, string: str) -> str:
        """LCase function."""
        return self._to_string(string).lower()

    def _builtin_instr(self, *args) -> int:
        """InStr function."""
        if len(args) == 2:
            string1, string2 = args
            start = 1
        elif len(args) >= 3:
            start, string1, string2 = args[0], args[1], args[2]
        else:
            return 0
        
        s1 = self._to_string(string1)
        s2 = self._to_string(string2)
        idx = s1.lower().find(s2.lower()) if len(args) > 2 and isinstance(args[0], int) and args[0] == 1 else s1.find(s2)
        return idx + 1 if idx >= 0 else 0

    def _builtin_replace(self, string: str, find: str, replace_with: str, start: int = 1, count: int = -1, compare: int = 0) -> str:
        """Replace function."""
        s = self._to_string(string)
        f = self._to_string(find)
        r = self._to_string(replace_with)
        return s.replace(f, r, count if count > 0 else -1)

    def _builtin_split(self, string: str, delimiter: str = " ", count: int = -1, compare: int = 0) -> list:
        """Split function."""
        s = self._to_string(string)
        d = self._to_string(delimiter)
        return s.split(d)

    def _builtin_join(self, array: list, delimiter: str = " ") -> str:
        """Join function."""
        d = self._to_string(delimiter)
        return d.join(self._to_string(item) for item in array)

    def _builtin_cstr(self, value: Any) -> str:
        """CStr function."""
        return self._to_string(value)

    def _builtin_cint(self, value: Any) -> int:
        """CInt function."""
        return int(self._to_number(value))

    def _builtin_clng(self, value: Any) -> int:
        """CLng function."""
        return int(self._to_number(value))

    def _builtin_cdbl(self, value: Any) -> float:
        """CDbl function."""
        return self._to_number(value)

    def _builtin_cbool(self, value: Any) -> bool:
        """CBool function."""
        return self._to_boolean(value)

    def _builtin_cdate(self, value: Any) -> Any:
        """CDate function (simplified)."""
        return self._to_string(value)

    def _builtin_isnumeric(self, value: Any) -> bool:
        """IsNumeric function."""
        if isinstance(value, (int, float)):
            return True
        if isinstance(value, str):
            try:
                float(value)
                return True
            except ValueError:
                return False
        return False

    def _builtin_isarray(self, value: Any) -> bool:
        """IsArray function."""
        return isinstance(value, (list, tuple))

    def _builtin_isdate(self, value: Any) -> bool:
        """IsDate function (simplified)."""
        return False

    def _builtin_isempty(self, value: Any) -> bool:
        """IsEmpty function."""
        return isinstance(value, VBScriptEmpty)

    def _builtin_isnull(self, value: Any) -> bool:
        """IsNull function."""
        return isinstance(value, VBScriptNull)

    def _builtin_isobject(self, value: Any) -> bool:
        """IsObject function."""
        return isinstance(value, (WScriptObject, VBScriptObject)) or value is not None and not isinstance(value, (str, int, float, bool, VBScriptEmpty, VBScriptNull, VBScriptNothing))

    def _builtin_typename(self, value: Any) -> str:
        """TypeName function."""
        if isinstance(value, VBScriptEmpty):
            return "Empty"
        if isinstance(value, VBScriptNull):
            return "Null"
        if isinstance(value, VBScriptNothing):
            return "Nothing"
        if isinstance(value, bool):
            return "Boolean"
        if isinstance(value, int):
            return "Integer"
        if isinstance(value, float):
            return "Double"
        if isinstance(value, str):
            return "String"
        if isinstance(value, (list, tuple)):
            return "Variant()"
        return "Object"

    def _builtin_vartype(self, value: Any) -> int:
        """VarType function."""
        if isinstance(value, VBScriptEmpty):
            return 0  # vbEmpty
        if isinstance(value, VBScriptNull):
            return 1  # vbNull
        if isinstance(value, bool):
            return 11  # vbBoolean
        if isinstance(value, int):
            return 2  # vbInteger
        if isinstance(value, float):
            return 5  # vbDouble
        if isinstance(value, str):
            return 8  # vbString
        return 12  # vbVariant

    def _builtin_abs(self, value: Any) -> float:
        """Abs function."""
        return abs(self._to_number(value))

    def _builtin_sqr(self, value: Any) -> float:
        """Sqr function."""
        return math.sqrt(self._to_number(value))

    def _builtin_int(self, value: Any) -> int:
        """Int function."""
        return int(math.floor(self._to_number(value)))

    def _builtin_fix(self, value: Any) -> int:
        """Fix function."""
        return int(self._to_number(value))

    def _builtin_round(self, value: Any, decimals: int = 0) -> float:
        """Round function."""
        return round(self._to_number(value), int(decimals))

    def _builtin_rnd(self, number: float = 1) -> float:
        """Rnd function."""
        import random
        return random.random()

    def _builtin_randomize(self, seed: Any = None) -> None:
        """Randomize statement."""
        import random
        if seed is not None:
            random.seed(int(self._to_number(seed)))
        else:
            random.seed()

    def _builtin_createobject(self, class_name: str, server_name: str = None) -> Any:
        """CreateObject function (simplified)."""
        # This is a stub - in real VBScript this would create COM objects
        # For now, return a placeholder object
        return {"_class": class_name}

    def _builtin_getobject(self, path_name: str = None, class_name: str = None) -> Any:
        """GetObject function (simplified)."""
        # This is a stub - in real VBScript this would get COM objects
        return {"_path": path_name, "_class": class_name}


def run(source: str, output_stream=None) -> Any:
    """Parse and execute VBScript source code."""
    from .parser import parse
    program = parse(source)
    interpreter = Interpreter(output_stream=output_stream)
    return interpreter.interpret(program)
