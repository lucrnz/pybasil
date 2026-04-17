"""VBScript Tree-Walking Interpreter.

Sections
--------
- Dispatch setup      -- explicit dispatch tables, resolution to bound methods
- Execute handlers    -- ``_execute_*`` methods for statement nodes
- Evaluate handlers   -- ``_evaluate_*`` methods for expression nodes
- Binary / unary / comparison operators
- Coercion helpers    -- ``_to_number``, ``_to_string``, ``_to_boolean``

Runtime value types live in ``runtime.py``; built-in VBScript functions live
in ``builtins.py``.
"""

from __future__ import annotations
import math
from typing import Any, Callable, Dict, List, Union

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
    MeExpression,
    DimStatement,
    AssignmentStatement,
    SetStatement,
    PropertyAssignmentStatement,
    CallStatement,
    ExpressionStatement,
    IfStatement,
    SelectCaseStatement,
    CaseRange,
    CaseComparison,
    ForStatement,
    ForEachStatement,
    WhileStatement,
    DoLoopStatement,
    ExitStatement,
    SubStatement,
    FunctionStatement,
    ClassStatement,
    ClassMemberField,
    ClassMemberSub,
    ClassMemberFunction,
    PropertyGetStatement,
    PropertyLetStatement,
    PropertySetStatement,
    BinaryOp,
    UnaryOp,
    ComparisonOp,
    ExitType,
    LoopConditionType,
    OnErrorResumeNextStatement,
    OnErrorGoToStatement,
    ErrorHandlingMode,
    ReDimStatement,
    EraseStatement,
)
from .runtime import (
    VBScriptError,
    VBScriptNothing,
    VBScriptEmpty,
    VBScriptNull,
    NOTHING,
    EMPTY,
    NULL,
    _NOT_FOUND,
    VBScriptArray,
    _DictItemAccessor,
    _DictKeyAccessor,
    VBScriptDictionary,
    VBScriptClassDef,
    VBScriptClassInstance,
    ClassPropertyDef,
    ClassMethodDef,
    ClassFieldDef,
    ErrObject,
    WScriptObject,
    ExitLoopException,
    ExitProcedureException,
    Procedure,
    Environment,
)
from .builtins import get_builtin_table

# ---------------------------------------------------------------------------
#  VBScript built-in constants (VBScript 6.0)
# ---------------------------------------------------------------------------

VBSCRIPT_CONSTANTS: Dict[str, Any] = {
    # String constants
    'vbcr': '\r',
    'vblf': '\n',
    'vbcrlf': '\r\n',
    'vbnewline': '\r\n',
    'vbtab': '\t',
    'vbnullchar': '\x00',
    'vbnullstring': '',
    'vbformfeed': '\x0c',
    'vbverticaltab': '\x0b',
    'vbback': '\x08',
    # VarType constants
    'vbempty': 0,
    'vbnull': 1,
    'vbinteger': 2,
    'vblong': 3,
    'vbsingle': 4,
    'vbdouble': 5,
    'vbcurrency': 6,
    'vbdate': 7,
    'vbstring': 8,
    'vbobject': 9,
    'vberror': 10,
    'vbboolean': 11,
    'vbvariant': 12,
    'vbdataobject': 13,
    'vbdecimal': 14,
    'vbbyte': 17,
    'vbarray': 8192,
    # MsgBox button constants
    'vbokonly': 0,
    'vbokcancel': 1,
    'vbabortretryignore': 2,
    'vbyesnocancel': 3,
    'vbyesno': 4,
    'vbretrycancel': 5,
    'vbcritical': 16,
    'vbquestion': 32,
    'vbexclamation': 48,
    'vbinformation': 64,
    'vbdefaultbutton1': 0,
    'vbdefaultbutton2': 256,
    'vbdefaultbutton3': 512,
    'vbdefaultbutton4': 768,
    'vbapplicationmodal': 0,
    'vbsystemmodal': 4096,
    # MsgBox return value constants
    'vbok': 1,
    'vbcancel': 2,
    'vbabort': 3,
    'vbretry': 4,
    'vbignore': 5,
    'vbyes': 6,
    'vbno': 7,
    # Comparison constants
    'vbbinarycompare': 0,
    'vbtextcompare': 1,
    'vbdatabasecompare': 2,
    # Tristate constants
    'vbusedefault': -2,
    'vbtrue': -1,
    'vbfalse': 0,
    # Color constants
    'vbblack': 0x000000,
    'vbred': 0x0000FF,
    'vbgreen': 0x00FF00,
    'vbyellow': 0x00FFFF,
    'vbblue': 0xFF0000,
    'vbmagenta': 0xFF00FF,
    'vbcyan': 0xFFFF00,
    'vbwhite': 0xFFFFFF,
    # Miscellaneous constants
    'vbobjecterror': -2147221504,
    'vbgeneraldate': 0,
    'vblongdate': 1,
    'vbshortdate': 2,
    'vblongtime': 3,
    'vbshorttime': 4,
}


def _is_numeric_not_bool(value: Any) -> bool:
    """True when *value* is int or float but not bool."""
    return isinstance(value, (int, float)) and not isinstance(value, bool)


_STATEMENT_ONLY_BUILTINS = {'execute', 'executeglobal'}


class Interpreter:
    """Tree-walking interpreter for VBScript AST."""

    def __init__(self, output_stream=None):
        self._environment = Environment()
        self._global_environment = self._environment
        self._output_stream = output_stream
        self._procedures: Dict[str, Procedure] = {}  # User-defined procedures
        self._class_defs: Dict[str, VBScriptClassDef] = {}  # User-defined classes
        self._current_instance: VBScriptClassInstance | None = None  # Me reference
        self._error_mode: ErrorHandlingMode = ErrorHandlingMode.DEFAULT
        self._err: ErrObject = ErrObject()
        self._setup_builtins()
        self._resolve_dispatch_tables()

    def _setup_builtins(self) -> None:
        """Set up built-in objects and functions."""
        # Create WScript object
        wscript = WScriptObject(self._output_stream)
        self._environment.define('WScript', wscript)

        # Create Err object
        self._environment.define('Err', self._err)

        # Built-in constants (seeded into the global environment)
        for name, value in VBSCRIPT_CONSTANTS.items():
            self._environment.define(name, value)

        # Built-in functions (defined in builtins.py)
        self._builtins: Dict[str, Callable] = get_builtin_table(self)

    def interpret(self, program: Program) -> Any:
        """Interpret a VBScript program."""
        result = None
        for statement in program.statements:
            result = self._execute_with_error_handling(statement)
        return result

    def _execute_with_error_handling(self, node: ASTNode) -> Any:
        """Execute a statement with error handling based on current mode."""
        try:
            return self._execute(node)
        except VBScriptError as e:
            # Check if this error was raised by Err.Raise
            # Err.Raise sets the properties before raising, so we shouldn't overwrite
            if self._err._number == 0 or not self._err._description:
                # Set the Err object with error information
                self._err._number = self._get_error_number(e)
                self._err._description = str(e)
                self._err._source = 'Microsoft VBScript runtime error'

            if self._error_mode == ErrorHandlingMode.RESUME_NEXT:
                # Continue to next statement
                return None
            else:
                # Re-raise the error
                raise
        except ExitLoopException:
            # These should always propagate
            raise
        except ExitProcedureException:
            # These should always propagate
            raise

    def _get_error_number(self, error: VBScriptError) -> int:
        """Extract or generate an error number from a VBScriptError."""
        # Try to parse error number from the error message
        msg = str(error)
        if msg.startswith('Error ') and ':' in msg:
            try:
                num_part = msg.split(':')[0].replace('Error ', '').strip()
                return int(num_part)
            except ValueError:
                pass
        # Default error numbers for common errors
        if 'Type mismatch' in msg:
            return 13  # Type mismatch
        elif 'Division by zero' in msg:
            return 11  # Division by zero
        elif 'Object required' in msg:
            return 424  # Object required
        elif 'Unknown procedure' in msg or 'Unknown function' in msg:
            return 438  # Object doesn't support this property or method
        elif 'Subscript out of range' in msg:
            return 9  # Subscript out of range
        elif 'Syntax error' in msg:
            return 1002  # Syntax error
        return 1000  # Generic runtime error

    # Explicit dispatch tables: {ASTNode subclass -> method name}.
    # Resolved to bound methods once in __init__ via _resolve_dispatch_tables.
    _EXECUTE_DISPATCH = {
        DimStatement: '_execute_DimStatement',
        AssignmentStatement: '_execute_AssignmentStatement',
        SetStatement: '_execute_SetStatement',
        PropertyAssignmentStatement: '_execute_PropertyAssignmentStatement',
        CallStatement: '_execute_CallStatement',
        ExpressionStatement: '_execute_ExpressionStatement',
        IfStatement: '_execute_IfStatement',
        SelectCaseStatement: '_execute_SelectCaseStatement',
        ForStatement: '_execute_ForStatement',
        ForEachStatement: '_execute_ForEachStatement',
        WhileStatement: '_execute_WhileStatement',
        DoLoopStatement: '_execute_DoLoopStatement',
        ExitStatement: '_execute_ExitStatement',
        SubStatement: '_execute_SubStatement',
        FunctionStatement: '_execute_FunctionStatement',
        OnErrorResumeNextStatement: '_execute_OnErrorResumeNextStatement',
        OnErrorGoToStatement: '_execute_OnErrorGoToStatement',
        ReDimStatement: '_execute_ReDimStatement',
        EraseStatement: '_execute_EraseStatement',
        ClassStatement: '_execute_ClassStatement',
    }

    _EVALUATE_DISPATCH = {
        NumberLiteral: '_evaluate_NumberLiteral',
        StringLiteral: '_evaluate_StringLiteral',
        BooleanLiteral: '_evaluate_BooleanLiteral',
        NothingLiteral: '_evaluate_NothingLiteral',
        EmptyLiteral: '_evaluate_EmptyLiteral',
        NullLiteral: '_evaluate_NullLiteral',
        Identifier: '_evaluate_Identifier',
        BinaryExpression: '_evaluate_BinaryExpression',
        UnaryExpression: '_evaluate_UnaryExpression',
        ComparisonExpression: '_evaluate_ComparisonExpression',
        MemberAccess: '_evaluate_MemberAccess',
        FunctionCall: '_evaluate_FunctionCall',
        MethodCall: '_evaluate_MethodCall',
        NewExpression: '_evaluate_NewExpression',
        ArrayAccess: '_evaluate_ArrayAccess',
        MeExpression: '_evaluate_MeExpression',
    }

    def _resolve_dispatch_tables(self) -> None:
        """Resolve method-name strings to bound methods once at init time."""
        self._execute_dispatch: Dict[type, Callable] = {
            cls: getattr(self, name)
            for cls, name in self._EXECUTE_DISPATCH.items()
        }
        self._evaluate_dispatch: Dict[type, Callable] = {
            cls: getattr(self, name)
            for cls, name in self._EVALUATE_DISPATCH.items()
        }
        self._binop_dispatch: Dict[BinaryOp, Callable] = {
            op: getattr(self, name)
            for op, name in self._BINOP_DISPATCH_NAMES.items()
        }

    def _execute(self, node: ASTNode) -> Any:
        """Execute a statement node."""
        handler = self._execute_dispatch.get(type(node))
        if handler is not None:
            return handler(node)
        raise VBScriptError(f'Unknown statement type: {type(node).__name__}')

    # ------------------------------------------------------------------
    #  Execute handlers
    # ------------------------------------------------------------------

    def _execute_DimStatement(self, node: DimStatement) -> None:
        """Execute a Dim statement."""
        for dim_var in node.variables:
            if dim_var.dimensions is not None:
                # Array declaration
                if len(dim_var.dimensions) == 0:
                    # Dynamic array: Dim arr()
                    arr = VBScriptArray([], is_dynamic=True)
                else:
                    # Fixed-size array: Dim arr(5) or Dim arr(2, 3)
                    dims = [int(self._evaluate(d)) for d in dim_var.dimensions]
                    arr = VBScriptArray(dims, is_dynamic=False)
                self._environment.define(dim_var.name, arr)
            else:
                # Simple variable
                self._environment.define(dim_var.name, EMPTY)

    def _execute_AssignmentStatement(self, node: AssignmentStatement) -> None:
        """Execute an assignment statement."""
        value = self._evaluate(node.expression)

        if node.indices:
            # Array or dictionary element assignment
            obj = self._environment.get(node.variable)
            if isinstance(obj, VBScriptArray):
                indices = [int(self._evaluate(idx)) for idx in node.indices]
                obj.set_element(indices, value)
            elif isinstance(obj, VBScriptDictionary):
                if len(node.indices) == 1:
                    key = self._evaluate(node.indices[0])
                    obj.set_item(key, value)
                else:
                    raise VBScriptError(
                        'Wrong number of arguments or invalid property assignment'
                    )
            else:
                raise VBScriptError('Type mismatch: expected array or dictionary')
        else:
            # Simple variable assignment
            self._environment.set(node.variable, value)

    def _execute_SetStatement(self, node: SetStatement) -> None:
        """Execute a Set statement."""
        value = self._evaluate(node.expression)

        if node.indices:
            # Array or dictionary element assignment
            obj = self._environment.get(node.variable)
            if isinstance(obj, VBScriptArray):
                indices = [int(self._evaluate(idx)) for idx in node.indices]
                obj.set_element(indices, value)
            elif isinstance(obj, VBScriptDictionary):
                if len(node.indices) == 1:
                    key = self._evaluate(node.indices[0])
                    obj.set_item(key, value)
                else:
                    raise VBScriptError(
                        'Wrong number of arguments or invalid property assignment'
                    )
            else:
                raise VBScriptError('Type mismatch: expected array or dictionary')
        else:
            # Simple variable assignment
            self._environment.set(node.variable, value)

    def _execute_PropertyAssignmentStatement(self, node) -> None:
        """Execute a property assignment statement like obj.Prop = value or obj.Prop("key") = value."""
        value = self._evaluate(node.expression)
        target = node.target

        if isinstance(target, MemberAccess):
            # obj.Property = value or obj.Property(args) = value
            obj = self._evaluate(target.object)

            if isinstance(obj, VBScriptClassInstance):
                self._class_instance_member_set(obj, target.member, value)
                return

            if isinstance(obj, VBScriptDictionary):
                # Handle dictionary property assignment
                prop_name = target.member.lower()
                if prop_name == 'item':
                    # This should be handled by MethodCall, not MemberAccess
                    # But if we get here, it's a property assignment without args
                    raise VBScriptError(
                        'Wrong number of arguments or invalid property assignment'
                    )
                elif prop_name == 'comparemode':
                    obj.CompareMode = int(value)
                    return
                elif prop_name == 'key':
                    raise VBScriptError(
                        'Wrong number of arguments or invalid property assignment'
                    )
                else:
                    raise VBScriptError(
                        f"Object doesn't support this property or method: {target.member}"
                    )
            else:
                # Try to set the attribute
                raise VBScriptError(
                    f"Object doesn't support this property or method: {target.member}"
                )

        elif isinstance(target, MethodCall):
            # obj.Method(args) = value - this is property assignment with arguments
            obj = self._evaluate(target.object)

            if isinstance(obj, VBScriptClassInstance):
                member = target.method.lower()
                prop = obj._class_def.properties.get(member)
                if prop and prop.let_body:
                    args = [self._evaluate(a) for a in target.arguments] + [value]
                    self._call_class_property_let(obj, prop, args, args_evaluated=True)
                    return
                raise VBScriptError(
                    f"Object doesn't support this property or method: {target.method}"
                )

            if isinstance(obj, VBScriptDictionary):
                method_name = target.method.lower()
                if method_name == 'item':
                    # dict.Item("key") = value
                    if len(target.arguments) == 1:
                        key = self._evaluate(target.arguments[0])
                        obj.set_item(key, value)
                        return
                    else:
                        raise VBScriptError(
                            'Wrong number of arguments or invalid property assignment'
                        )
                elif method_name == 'key':
                    # dict.Key("oldkey") = "newkey" - change a key
                    if len(target.arguments) == 1:
                        old_key = self._evaluate(target.arguments[0])
                        obj.set_key(old_key, value)
                        return
                    else:
                        raise VBScriptError(
                            'Wrong number of arguments or invalid property assignment'
                        )
                else:
                    raise VBScriptError(
                        f"Object doesn't support this property or method: {target.method}"
                    )
            else:
                raise VBScriptError("Object doesn't support this property or method")

        elif isinstance(target, ArrayAccess):
            # arr(index) = value or dict("key") = value (default property)
            obj = self._environment.get(target.name)

            if isinstance(obj, VBScriptDictionary):
                # dict("key") = value - default property (Item) assignment
                if len(target.indices) == 1:
                    key = self._evaluate(target.indices[0])
                    obj.set_item(key, value)
                    return
                else:
                    raise VBScriptError(
                        'Wrong number of arguments or invalid property assignment'
                    )
            elif isinstance(obj, VBScriptArray):
                # Array element assignment - this should be handled by AssignmentStatement
                indices = [int(self._evaluate(idx)) for idx in target.indices]
                obj.set_element(indices, value)
                return
            else:
                raise VBScriptError('Type mismatch: expected array or dictionary')

        else:
            raise VBScriptError(f'Invalid assignment target: {type(target).__name__}')

    def _execute_CallStatement(self, node: CallStatement) -> Any:
        """Execute a Call statement."""
        return self._call_procedure(node.name, node.arguments)

    def _execute_SubStatement(self, node: SubStatement) -> None:
        """Register a Sub procedure."""
        proc = Procedure(
            name=node.name.lower(),
            parameters=node.parameters,
            body=node.body,
            is_function=False,
        )
        self._procedures[node.name.lower()] = proc

    def _execute_FunctionStatement(self, node: FunctionStatement) -> None:
        """Register a Function procedure."""
        proc = Procedure(
            name=node.name.lower(),
            parameters=node.parameters,
            body=node.body,
            is_function=True,
        )
        self._procedures[node.name.lower()] = proc

    def _call_procedure(self, name: str, arguments: List[ASTNode]) -> Any:
        """Call a user-defined procedure or built-in function."""
        proc_name = name.lower()

        # Check for user-defined procedure
        if proc_name in self._procedures:
            proc = self._procedures[proc_name]
            return self._execute_procedure(proc, arguments)

        if proc_name in _STATEMENT_ONLY_BUILTINS:
            raise VBScriptError('Syntax error')

        # Check for built-in function
        if proc_name in self._builtins:
            args = [self._evaluate(arg) for arg in arguments]
            return self._builtins[proc_name](*args)

        raise VBScriptError(f'Unknown procedure: {name}')

    def _execute_procedure(self, proc: Procedure, arguments: List[ASTNode]) -> Any:
        """Execute a user-defined procedure with proper scoping."""
        # Create a new environment for the procedure
        old_env = self._environment
        old_error_mode = self._error_mode
        proc_env = Environment(parent=old_env)
        self._environment = proc_env
        # Reset error mode to default at procedure entry
        self._error_mode = ErrorHandlingMode.DEFAULT
        self._err.Clear()

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
                    self._execute_with_error_handling(stmt)
            except ExitProcedureException as e:
                # Exit Sub/Function was called
                if proc.is_function:
                    return proc_env.get(proc.name)
                return e.return_value

            # Handle Exit For/Do that propagated up
            except ExitLoopException:
                raise VBScriptError('Exit For/Do not valid outside of loops')

            # For functions, return the function's return value
            if proc.is_function:
                return proc_env.get(proc.name)

            return EMPTY

        finally:
            # Copy ByRef values back to the original scope
            for param_name, (orig_env, orig_var) in byref_bindings.items():
                orig_env.set(orig_var, proc_env.get(param_name))

            # Restore the original environment and error mode
            self._environment = old_env
            self._error_mode = old_error_mode

    def _execute_ExpressionStatement(self, node: ExpressionStatement) -> Any:
        """Execute an expression statement."""
        # Handle misparse of implicit call with unary minus/plus argument.
        # The parser sees "WScript.Echo -1" as BinaryExpression(SUB,
        # MemberAccess(WScript, Echo), 1) instead of a method call with
        # argument -1.  Re-interpret as MethodCall when the left operand
        # is a MemberAccess (i.e. looks like a callable).
        if isinstance(node.expression, BinaryExpression) and node.expression.operator in (BinaryOp.ADD, BinaryOp.SUB):
            left = node.expression.left
            if isinstance(left, MemberAccess):
                arg = node.expression.right
                if node.expression.operator == BinaryOp.SUB:
                    arg = UnaryExpression(operator=UnaryOp.NEG, operand=arg)
                rewritten = MethodCall(
                    object=left.object, method=left.member, arguments=[arg]
                )
                return self._evaluate(rewritten)

        # Check if this is a procedure call (FunctionCall to a Sub or Function)
        if isinstance(node.expression, FunctionCall):
            name = node.expression.name.lower()
            if name in _STATEMENT_ONLY_BUILTINS:
                args = [self._evaluate(arg) for arg in node.expression.arguments]
                return self._builtins[name](*args)
            if name in self._procedures:
                return self._call_procedure(
                    node.expression.name, node.expression.arguments
                )

        if isinstance(node.expression, ArrayAccess):
            name = node.expression.name.lower()
            if name in _STATEMENT_ONLY_BUILTINS:
                args = [self._evaluate(arg) for arg in node.expression.indices]
                return self._builtins[name](*args)

        # Check if this is a procedure call (identifier without args)
        if isinstance(node.expression, Identifier):
            name = node.expression.name.lower()
            if name in self._procedures:
                return self._call_procedure(node.expression.name, [])

        # Check for Err.Clear() or Err.Raise() as MemberAccess (with parentheses)
        if isinstance(node.expression, MemberAccess):
            if isinstance(node.expression.object, Identifier):
                if node.expression.object.name.lower() == 'err':
                    method = node.expression.member.lower()
                    if method == 'clear':
                        self._err.Clear()
                        return None
                    elif method == 'raise':
                        # Err.Raise without arguments - raise generic error
                        self._err.Raise(0)
                        return None

        # Check if this is a method call that should invoke a procedure
        if isinstance(node.expression, MethodCall):
            # Handle Err.Raise specially (Err.Clear is handled in interpret())
            if isinstance(node.expression.object, Identifier):
                if node.expression.object.name.lower() == 'err':
                    method = node.expression.method.lower()
                    if method == 'raise':
                        args = [
                            self._evaluate(arg) for arg in node.expression.arguments
                        ]
                        self._err.Raise(*args) if args else self._err.Raise(0)
                        return None
            return self._evaluate(node.expression)

        return self._evaluate(node.expression)

    def _execute_IfStatement(self, node: IfStatement) -> Any:
        """Execute an If statement."""
        condition = self._evaluate(node.condition)
        if self._to_boolean(condition):
            for stmt in node.then_body:
                self._execute_with_error_handling(stmt)
            return None

        # Check ElseIf clauses
        for elseif in node.elseif_clauses:
            condition = self._evaluate(elseif.condition)
            if self._to_boolean(condition):
                for stmt in elseif.body:
                    self._execute_with_error_handling(stmt)
                return None

        # Execute Else clause if present
        if node.else_clause:
            for stmt in node.else_clause.body:
                self._execute_with_error_handling(stmt)

        return None

    def _execute_SelectCaseStatement(self, node: SelectCaseStatement) -> Any:
        """Execute a Select Case statement."""
        # Evaluate the select expression
        select_value = self._evaluate(node.expression)

        # Check each Case clause
        for case_clause in node.case_clauses:
            # Check if any of the case values match
            for case_value_node in case_clause.values:
                matched = False

                if isinstance(case_value_node, CaseRange):
                    low = self._evaluate(case_value_node.low)
                    high = self._evaluate(case_value_node.high)
                    matched = (self._to_number(low) <= self._to_number(select_value)
                               <= self._to_number(high))
                elif isinstance(case_value_node, CaseComparison):
                    comp_value = self._evaluate(case_value_node.expression)
                    matched = self._apply_comparison_op(
                        case_value_node.operator, select_value, comp_value)
                else:
                    case_value = self._evaluate(case_value_node)
                    if isinstance(select_value, bool):
                        matched = self._to_boolean(case_value) == select_value
                    else:
                        matched = self._values_equal(select_value, case_value)

                if matched:
                    for stmt in case_clause.body:
                        self._execute_with_error_handling(stmt)
                    return None

        # Execute Case Else if no match found
        if node.case_else_clause:
            for stmt in node.case_else_clause.body:
                self._execute_with_error_handling(stmt)

        return None

    def _values_equal(self, left: Any, right: Any) -> bool:
        """Check if two values are equal for Select Case comparison."""
        # Handle Empty values
        if isinstance(left, VBScriptEmpty) and isinstance(right, VBScriptEmpty):
            return True
        if isinstance(left, VBScriptEmpty):
            left = 0 if isinstance(right, (int, float)) else ''
        if isinstance(right, VBScriptEmpty):
            right = 0 if isinstance(left, (int, float)) else ''

        # Handle Null - Null only equals Null
        if isinstance(left, VBScriptNull) and isinstance(right, VBScriptNull):
            return True
        if isinstance(left, VBScriptNull) or isinstance(right, VBScriptNull):
            return False

        # Handle Nothing
        if isinstance(left, VBScriptNothing) and isinstance(right, VBScriptNothing):
            return True

        # Handle strings (case-insensitive comparison in VBScript)
        if isinstance(left, str) and isinstance(right, str):
            return left.lower() == right.lower()

        # Handle numbers
        if isinstance(left, (int, float)) and isinstance(right, (int, float)):
            return left == right

        # Handle booleans
        if isinstance(left, bool) and isinstance(right, bool):
            return left == right

        # Mixed type comparison
        if isinstance(left, str) and isinstance(right, (int, float)):
            try:
                return float(left) == right
            except ValueError:
                return False
        if isinstance(left, (int, float)) and isinstance(right, str):
            try:
                return left == float(right)
            except ValueError:
                return False

        # Default comparison
        return left == right

    def _execute_ForStatement(self, node: ForStatement) -> Any:
        """Execute a For...Next statement."""
        start_val = self._to_number(self._evaluate(node.start))
        end_val = self._to_number(self._evaluate(node.end))

        # Determine step value
        if node.step:
            step_val = self._to_number(self._evaluate(node.step))
        else:
            # Default step is 1, or -1 if start > end
            step_val = 1 if start_val <= end_val else -1

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
                        self._execute_with_error_handling(stmt)
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

    def _execute_ForEachStatement(self, node: ForEachStatement) -> Any:
        """Execute a For Each...Next statement."""
        # Evaluate the collection
        collection = self._evaluate(node.collection)

        # Get an iterable from the collection
        if isinstance(collection, VBScriptArray):
            iterable = list(collection)
        elif isinstance(collection, VBScriptDictionary):
            iterable = list(collection)  # VBScriptDictionary.__iter__ yields items
        elif isinstance(collection, list):
            iterable = collection
        elif isinstance(collection, str):
            # Strings are iterable character by character in VBScript
            iterable = list(collection)
        else:
            raise VBScriptError("Object doesn't support this property or method")

        try:
            for item in iterable:
                # Set the loop variable
                self._environment.set(node.variable, item)

                # Execute body
                try:
                    for stmt in node.body:
                        self._execute_with_error_handling(stmt)
                except ExitLoopException as e:
                    if e.exit_type == ExitType.FOR:
                        return None
                    raise
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
                        self._execute_with_error_handling(stmt)
                except ExitLoopException as e:
                    if e.exit_type == ExitType.DO:
                        raise VBScriptError('Exit Do not valid in While loop')
                    elif e.exit_type == ExitType.FOR:
                        raise VBScriptError('Exit For not valid in While loop')
                    raise
        except ExitLoopException as e:
            if e.exit_type == ExitType.DO:
                raise VBScriptError('Exit Do not valid in While loop')
            elif e.exit_type == ExitType.FOR:
                raise VBScriptError('Exit For not valid in While loop')
            raise

        return None

    def _execute_DoLoopStatement(self, node: DoLoopStatement) -> Any:
        """Execute a Do...Loop statement."""
        try:
            while True:
                # Check pre-condition (Do While/Until)
                if node.pre_condition:
                    cond_result = self._to_boolean(
                        self._evaluate(node.pre_condition.condition)
                    )
                    if node.pre_condition.condition_type == LoopConditionType.WHILE:
                        if not cond_result:
                            break
                    else:  # UNTIL
                        if cond_result:
                            break

                # Execute body
                try:
                    for stmt in node.body:
                        self._execute_with_error_handling(stmt)
                except ExitLoopException as e:
                    if e.exit_type == ExitType.DO:
                        return None
                    raise

                # Check post-condition (Loop While/Until)
                if node.post_condition:
                    cond_result = self._to_boolean(
                        self._evaluate(node.post_condition.condition)
                    )
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
        if node.exit_type in (ExitType.SUB, ExitType.FUNCTION, ExitType.PROPERTY):
            raise ExitProcedureException(node.exit_type)
        else:
            raise ExitLoopException(node.exit_type)

    def _execute_OnErrorResumeNextStatement(
        self, node: OnErrorResumeNextStatement
    ) -> None:
        """Execute On Error Resume Next statement."""
        self._error_mode = ErrorHandlingMode.RESUME_NEXT
        self._err.Clear()

    def _execute_OnErrorGoToStatement(self, node: OnErrorGoToStatement) -> None:
        """Execute On Error GoTo statement."""
        if node.label == 0:
            # On Error GoTo 0 - reset to default error handling
            self._error_mode = ErrorHandlingMode.DEFAULT
        else:
            # On Error GoTo label - set goto mode (line number not fully supported)
            self._error_mode = ErrorHandlingMode.GOTO
        self._err.Clear()

    def _execute_ReDimStatement(self, node: ReDimStatement) -> None:
        """Execute a ReDim statement."""
        for name, dimensions in node.arrays:
            arr = self._environment.get(name)

            # If the variable doesn't exist or is not an array, create a new dynamic array
            if not isinstance(arr, VBScriptArray):
                arr = VBScriptArray([], is_dynamic=True)
                self._environment.set(name, arr)

            # Check if it's a dynamic array
            if not arr._is_dynamic:
                raise VBScriptError('This array is fixed or temporarily locked')

            # Evaluate dimensions
            dims = [int(self._evaluate(d)) for d in dimensions]

            # Resize the array
            arr.redim(dims, preserve=node.preserve)

    def _execute_EraseStatement(self, node: EraseStatement) -> None:
        """Execute an Erase statement."""
        for name in node.arrays:
            arr = self._environment.get(name)
            if isinstance(arr, VBScriptArray):
                arr.erase()

    # ------------------------------------------------------------------
    #  Class handler
    # ------------------------------------------------------------------

    def _execute_ClassStatement(self, node: ClassStatement) -> None:
        """Register a Class definition."""
        class_def = VBScriptClassDef(node.name)

        for member in node.members:
            if isinstance(member, ClassMemberField):
                if member.dim:
                    for dv in member.dim.variables:
                        class_def.fields.append(
                            ClassFieldDef(
                                name=dv.name,
                                is_public=member.is_public,
                                dimensions=dv.dimensions,
                            )
                        )
            elif isinstance(member, ClassMemberSub):
                proc = Procedure(
                    name=member.sub.name.lower(),
                    parameters=member.sub.parameters,
                    body=member.sub.body,
                    is_function=False,
                )
                method_def = ClassMethodDef(
                    proc=proc, is_public=member.is_public,
                    is_default=member.is_default,
                )
                class_def.methods[proc.name] = method_def
                if member.is_default:
                    class_def.default_member = proc.name
            elif isinstance(member, ClassMemberFunction):
                proc = Procedure(
                    name=member.function.name.lower(),
                    parameters=member.function.parameters,
                    body=member.function.body,
                    is_function=True,
                )
                method_def = ClassMethodDef(
                    proc=proc, is_public=member.is_public,
                    is_default=member.is_default,
                )
                class_def.methods[proc.name] = method_def
                if member.is_default:
                    class_def.default_member = proc.name
            elif isinstance(member, PropertyGetStatement):
                prop_name = member.name.lower()
                prop = class_def.properties.setdefault(
                    prop_name, ClassPropertyDef()
                )
                prop.get_params = member.parameters
                prop.get_body = member.body
                prop.is_public = member.is_public
                if member.is_default:
                    prop.is_default = True
                    class_def.default_member = prop_name
            elif isinstance(member, PropertyLetStatement):
                prop_name = member.name.lower()
                prop = class_def.properties.setdefault(
                    prop_name, ClassPropertyDef()
                )
                prop.let_params = member.parameters
                prop.let_body = member.body
                prop.is_public = member.is_public
            elif isinstance(member, PropertySetStatement):
                prop_name = member.name.lower()
                prop = class_def.properties.setdefault(
                    prop_name, ClassPropertyDef()
                )
                prop.set_params = member.parameters
                prop.set_body = member.body
                prop.is_public = member.is_public

        # Build field_names lookup dict
        for fld in class_def.fields:
            class_def.field_names[fld.name.lower()] = fld.name

        # Pre-cache Procedure objects for properties
        for prop_name, prop in class_def.properties.items():
            if prop.get_body is not None:
                prop._get_proc = Procedure(
                    name=prop_name,
                    parameters=prop.get_params or [],
                    body=prop.get_body,
                    is_function=True,
                )
            if prop.let_body is not None:
                prop._let_proc = Procedure(
                    name='__property_let__',
                    parameters=prop.let_params or [],
                    body=prop.let_body,
                    is_function=False,
                )
            if prop.set_body is not None:
                prop._set_proc = Procedure(
                    name='__property_set__',
                    parameters=prop.set_params or [],
                    body=prop.set_body,
                    is_function=False,
                )

        self._class_defs[node.name.lower()] = class_def

    def _instantiate_class(self, class_def: VBScriptClassDef) -> VBScriptClassInstance:
        """Create a new instance of a user-defined class."""
        inst_env = Environment(parent=self._environment)

        for fld in class_def.fields:
            if fld.dimensions is not None:
                if len(fld.dimensions) == 0:
                    inst_env.define(fld.name, VBScriptArray([], is_dynamic=True))
                else:
                    dims = [int(self._evaluate(d)) for d in fld.dimensions]
                    inst_env.define(fld.name, VBScriptArray(dims, is_dynamic=False))
            else:
                inst_env.define(fld.name, EMPTY)

        instance = VBScriptClassInstance(class_def, inst_env)

        # Run Class_Initialize if present
        init_method = class_def.methods.get('class_initialize')
        if init_method:
            self._call_class_method(instance, init_method.proc, [])

        return instance

    def _call_class_method(
        self, instance: VBScriptClassInstance, proc: Procedure,
        arguments: List[ASTNode], args_evaluated: bool = False,
    ) -> Any:
        """Execute a method on a class instance."""
        old_env = self._environment
        old_instance = self._current_instance
        old_error_mode = self._error_mode
        proc_env = Environment(parent=instance._env)
        self._environment = proc_env
        self._current_instance = instance
        self._error_mode = ErrorHandlingMode.DEFAULT
        self._err.Clear()

        try:
            if args_evaluated:
                # Fast path: args already evaluated, no ByRef handling needed
                for i, param in enumerate(proc.parameters):
                    proc_env.define(
                        param.name,
                        arguments[i] if i < len(arguments) else EMPTY,
                    )
            else:
                arg_values = [self._evaluate(arg) for arg in arguments]
                byref_bindings: Dict[str, tuple] = {}
                for i, param in enumerate(proc.parameters):
                    if i < len(arg_values):
                        if param.is_byref and isinstance(arguments[i], Identifier):
                            var_name = arguments[i].name
                            byref_bindings[param.name.lower()] = (old_env, var_name)
                            proc_env.define(param.name, old_env.get(var_name))
                        else:
                            proc_env.define(param.name, arg_values[i])
                    else:
                        proc_env.define(param.name, EMPTY)

            if proc.is_function:
                proc_env.define(proc.name, EMPTY)

            try:
                for stmt in proc.body:
                    self._execute_with_error_handling(stmt)
            except ExitProcedureException:
                if proc.is_function:
                    return proc_env.get(proc.name)
                return EMPTY

            if proc.is_function:
                return proc_env.get(proc.name)
            return EMPTY
        finally:
            if not args_evaluated:
                for param_name, (orig_env, orig_var) in byref_bindings.items():
                    orig_env.set(orig_var, proc_env.get(param_name))
            self._environment = old_env
            self._current_instance = old_instance
            self._error_mode = old_error_mode

    def _call_class_property_get(
        self, instance: VBScriptClassInstance, prop: ClassPropertyDef,
        arguments: List = None, args_evaluated: bool = False,
        prop_name: str = '',
    ) -> Any:
        """Execute a Property Get on a class instance."""
        proc = prop._get_proc
        if proc is None:
            raise VBScriptError("Object doesn't support this property or method")
        return self._call_class_method(
            instance, proc, arguments or [], args_evaluated=args_evaluated,
        )

    def _call_class_property_let(
        self, instance: VBScriptClassInstance, prop: ClassPropertyDef,
        arguments: List, args_evaluated: bool = False,
    ) -> None:
        """Execute a Property Let on a class instance."""
        proc = prop._let_proc
        if proc is None:
            raise VBScriptError("Object doesn't support this property or method")
        self._call_class_method(
            instance, proc, arguments, args_evaluated=args_evaluated,
        )

    def _call_class_property_set(
        self, instance: VBScriptClassInstance, prop: ClassPropertyDef,
        arguments: List, args_evaluated: bool = False,
    ) -> None:
        """Execute a Property Set on a class instance."""
        proc = prop._set_proc
        if proc is None:
            raise VBScriptError("Object doesn't support this property or method")
        self._call_class_method(
            instance, proc, arguments, args_evaluated=args_evaluated,
        )

    # ------------------------------------------------------------------
    #  Evaluate handlers
    # ------------------------------------------------------------------

    def _evaluate(self, node: ASTNode) -> Any:
        """Evaluate an expression node."""
        handler = self._evaluate_dispatch.get(type(node))
        if handler is not None:
            return handler(node)
        raise VBScriptError(f'Unknown expression type: {type(node).__name__}')

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
        """Evaluate an identifier.

        The scope-chain walk below is intentionally inlined rather than
        delegating to ``Environment.get()``.  Profiling showed a ~9% speedup
        because:
          1. ``Identifier._lower`` already holds the lowered key, so we skip
             the ``.lower()`` call that ``Environment.get()`` performs.
          2. The tight ``while`` loop avoids per-level method-call overhead.

        ``_NOT_FOUND`` (defined in ``runtime.py``) is a unique sentinel object
        used to distinguish "key absent from this scope" from "key is present
        but its value is ``None`` or ``Empty``".

        **Maintenance note:** if ``Environment.get()`` logic changes (e.g. a
        new scope rule), the inline loop here must be updated to match.
        """
        name_lower = node._lower

        # Check if this is a function call (function name without parentheses)
        if name_lower in self._procedures:
            proc = self._procedures[name_lower]
            if proc.is_function:
                return self._execute_procedure(proc, [])

        # Inline fast path -- see docstring above for rationale.
        env = self._environment
        _miss = _NOT_FOUND
        while env is not None:
            val = env._variables.get(name_lower, _miss)
            if val is not _miss:
                return val
            env = env._parent

        # Implicit member resolution: inside a class method, bare identifiers
        # that aren't found in the scope chain resolve to properties/methods
        # on the current instance (equivalent to Me.<name>).
        inst = self._current_instance
        if inst is not None:
            class_def = inst._class_def
            prop = class_def.properties.get(name_lower)
            if prop is not None and prop.get_body is not None:
                return self._call_class_property_get(
                    inst, prop, prop_name=name_lower,
                )
            method_def = class_def.methods.get(name_lower)
            if method_def is not None and not method_def.proc.parameters:
                return self._call_class_method(
                    inst, method_def.proc, [], args_evaluated=True,
                )

        return EMPTY

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
            raise VBScriptError(f'Object required: {node.member}')

        # Handle WScript object
        if isinstance(obj, WScriptObject):
            attr_name = node.member.lower()
            if attr_name == 'echo':
                return getattr(obj, 'Echo')
            elif attr_name == 'quit':
                return getattr(obj, 'Quit')
            else:
                raise VBScriptError(f'Unknown member: WScript.{node.member}')

        # Handle Err object with case-insensitive access
        if isinstance(obj, ErrObject):
            attr_name = node.member.lower()
            attr_map = {
                'number': 'Number',
                'source': 'Source',
                'description': 'Description',
                'helpfile': 'HelpFile',
                'helpcontext': 'HelpContext',
                'clear': 'Clear',
                'raise': 'Raise',
            }
            if attr_name in attr_map:
                return getattr(obj, attr_map[attr_name])
            raise VBScriptError(f'Unknown member: Err.{node.member}')

        # Handle VBScriptDictionary
        if isinstance(obj, VBScriptDictionary):
            attr_name = node.member.lower()
            # Properties
            if attr_name == 'count':
                return obj.Count
            elif attr_name == 'comparemode':
                return obj.CompareMode
            # Methods - return bound method
            elif attr_name == 'add':
                return obj.Add
            elif attr_name == 'exists':
                return obj.Exists
            elif attr_name == 'items':
                return obj.Items
            elif attr_name == 'keys':
                return obj.Keys
            elif attr_name == 'remove':
                return obj.Remove
            elif attr_name == 'removeall':
                return obj.RemoveAll
            elif attr_name == 'item':
                # Item is the default property - return a callable wrapper
                return _DictItemAccessor(obj)
            elif attr_name == 'key':
                # Key property
                return _DictKeyAccessor(obj)
            else:
                raise VBScriptError(
                    f"Object doesn't support this property or method: {node.member}"
                )

        # Handle user-defined class instances
        if isinstance(obj, VBScriptClassInstance):
            return self._class_instance_member_get(obj, node.member)

        # Handle dictionary-like objects
        if isinstance(obj, dict):
            return obj.get(node.member.lower(), EMPTY)

        # Handle objects with attributes
        if hasattr(obj, node.member):
            return getattr(obj, node.member)

        raise VBScriptError(
            f"Object doesn't support this property or method: {node.member}"
        )

    def _evaluate_FunctionCall(self, node: FunctionCall) -> Any:
        """Evaluate a function call."""
        func_name = node.name.lower()

        # Check for user-defined procedures first
        if func_name in self._procedures:
            proc = self._procedures[func_name]
            if not proc.is_function:
                self._execute_procedure(proc, node.arguments)
                return EMPTY
            return self._execute_procedure(proc, node.arguments)

        if func_name in _STATEMENT_ONLY_BUILTINS:
            raise VBScriptError('Syntax error')

        # Check built-in functions
        if func_name in self._builtins:
            args = [self._evaluate(arg) for arg in node.arguments]
            return self._builtins[func_name](*args)

        # Implicit member resolution: inside a class method, bare function
        # calls resolve to methods on the current instance.
        inst = self._current_instance
        if inst is not None:
            class_def = inst._class_def
            method_def = class_def.methods.get(func_name)
            if method_def is not None:
                args = [self._evaluate(arg) for arg in node.arguments]
                return self._call_class_method(
                    inst, method_def.proc, args, args_evaluated=True,
                )
            prop = class_def.properties.get(func_name)
            if prop is not None and prop.get_body is not None:
                args = [self._evaluate(arg) for arg in node.arguments]
                return self._call_class_property_get(
                    inst, prop, args, args_evaluated=True,
                    prop_name=func_name,
                )

        raise VBScriptError(f'Unknown function: {node.name}')

    def _evaluate_ArrayAccess(self, node: ArrayAccess) -> Any:
        """Evaluate an array access expression."""
        # First check if this is actually an array or a function call
        var = self._environment.get(node.name)

        if isinstance(var, VBScriptArray):
            # Array access
            indices = [int(self._evaluate(idx)) for idx in node.indices]
            return var.get_element(indices)
        elif isinstance(var, VBScriptDictionary):
            # Dictionary default property access (Item)
            if len(node.indices) == 1:
                key = self._evaluate(node.indices[0])
                return var.get_item(key)
            else:
                raise VBScriptError(
                    'Wrong number of arguments or invalid property assignment'
                )
        else:
            # This might be a function call - check builtins and procedures
            func_name = node.name.lower()

            if func_name in self._procedures:
                proc = self._procedures[func_name]
                if not proc.is_function:
                    self._execute_procedure(proc, node.indices)
                    return EMPTY
                return self._execute_procedure(proc, node.indices)

            if func_name in _STATEMENT_ONLY_BUILTINS:
                raise VBScriptError('Syntax error')

            if func_name in self._builtins:
                args = [self._evaluate(arg) for arg in node.indices]
                return self._builtins[func_name](*args)

            raise VBScriptError(f'Unknown function or array: {node.name}')

    def _evaluate_MethodCall(self, node: MethodCall) -> Any:
        """Evaluate a method call."""
        obj = self._evaluate(node.object)
        method = node.method
        args = [self._evaluate(arg) for arg in node.arguments]

        if callable(obj):
            return obj(*args)

        # Handle user-defined class instances
        if isinstance(obj, VBScriptClassInstance):
            return self._class_instance_method_call(obj, method, args)

        # Handle WScript object with case-insensitive method lookup
        if isinstance(obj, WScriptObject):
            method_lower = method.lower()
            if method_lower == 'echo':
                return obj.Echo(*args)
            elif method_lower == 'quit':
                return obj.Quit(*args)
            else:
                raise VBScriptError(
                    f"Object doesn't support this property or method: {method}"
                )

        # Handle VBScriptDictionary methods
        if isinstance(obj, VBScriptDictionary):
            method_lower = method.lower()
            if method_lower == 'add':
                if len(args) != 2:
                    raise VBScriptError(
                        'Wrong number of arguments or invalid property assignment'
                    )
                obj.Add(args[0], args[1])
                return None
            elif method_lower == 'exists':
                if len(args) != 1:
                    raise VBScriptError(
                        'Wrong number of arguments or invalid property assignment'
                    )
                return obj.Exists(args[0])
            elif method_lower == 'items':
                return obj.Items()
            elif method_lower == 'keys':
                return obj.Keys()
            elif method_lower == 'remove':
                if len(args) != 1:
                    raise VBScriptError(
                        'Wrong number of arguments or invalid property assignment'
                    )
                obj.Remove(args[0])
                return None
            elif method_lower == 'removeall':
                obj.RemoveAll()
                return None
            elif method_lower == 'item':
                if len(args) != 1:
                    raise VBScriptError(
                        'Wrong number of arguments or invalid property assignment'
                    )
                return obj.get_item(args[0])
            elif method_lower == 'key':
                if len(args) != 1:
                    raise VBScriptError(
                        'Wrong number of arguments or invalid property assignment'
                    )
                return obj.get_key(args[0])
            else:
                raise VBScriptError(
                    f"Object doesn't support this property or method: {method}"
                )

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
        class_name_lower = node.class_name.lower() if isinstance(node.class_name, str) else node.class_name.name.lower()
        class_def = self._class_defs.get(class_name_lower)
        if class_def is None:
            raise VBScriptError(f'Class not defined: {node.class_name}')
        return self._instantiate_class(class_def)

    def _evaluate_MeExpression(self, node: MeExpression) -> Any:
        """Evaluate the Me keyword."""
        if self._current_instance is None:
            raise VBScriptError('Invalid use of Me keyword')
        return self._current_instance

    # ------------------------------------------------------------------
    #  Class instance helpers
    # ------------------------------------------------------------------

    def _class_instance_member_get(self, obj: VBScriptClassInstance, member: str) -> Any:
        """Read a member (field, property, or zero-arg method) from a class instance."""
        name_lower = member.lower()
        class_def = obj._class_def

        # Check properties first (Property Get)
        prop = class_def.properties.get(name_lower)
        if prop is not None:
            return self._call_class_property_get(obj, prop, prop_name=name_lower)

        # Check fields (O(1) dict lookup)
        orig_name = class_def.field_names.get(name_lower)
        if orig_name is not None:
            return obj._env.get(orig_name)

        # Check methods - invoke zero-arg methods immediately
        method_def = class_def.methods.get(name_lower)
        if method_def is not None:
            if not method_def.proc.parameters:
                return self._call_class_method(obj, method_def.proc, [], args_evaluated=True)
            return method_def

        raise VBScriptError(
            f"Object doesn't support this property or method: {member}"
        )

    def _class_instance_member_set(self, obj: VBScriptClassInstance, member: str, value: Any) -> None:
        """Write to a member (field or Property Let/Set) on a class instance."""
        name_lower = member.lower()
        class_def = obj._class_def

        # Check Property Let/Set first
        prop = class_def.properties.get(name_lower)
        if prop is not None:
            if prop.let_body is not None:
                self._call_class_property_let(obj, prop, [value], args_evaluated=True)
                return
            if prop.set_body is not None:
                self._call_class_property_set(obj, prop, [value], args_evaluated=True)
                return

        # Check fields (O(1) dict lookup)
        orig_name = class_def.field_names.get(name_lower)
        if orig_name is not None:
            obj._env.set(orig_name, value)
            return

        raise VBScriptError(
            f"Object doesn't support this property or method: {member}"
        )

    def _class_instance_method_call(self, obj: VBScriptClassInstance, method: str, args: list) -> Any:
        """Call a method (Sub/Function) or Property Get with args on a class instance."""
        name_lower = method.lower()
        class_def = obj._class_def

        # Check methods
        method_def = class_def.methods.get(name_lower)
        if method_def is not None:
            return self._call_class_method(
                obj, method_def.proc, args, args_evaluated=True,
            )

        # Check Property Get with parameters (e.g. indexed property)
        prop = class_def.properties.get(name_lower)
        if prop is not None and prop.get_body is not None:
            return self._call_class_property_get(
                obj, prop, args, args_evaluated=True, prop_name=name_lower,
            )

        # Default member invocation with args (e.g. obj(args))
        if not method and class_def.default_member:
            default_name = class_def.default_member
            # Try method first
            dm = class_def.methods.get(default_name)
            if dm is not None:
                return self._call_class_method(
                    obj, dm.proc, args, args_evaluated=True,
                )
            dp = class_def.properties.get(default_name)
            if dp is not None and dp.get_body is not None:
                return self._call_class_property_get(
                    obj, dp, args, args_evaluated=True, prop_name=default_name,
                )

        raise VBScriptError(
            f"Object doesn't support this property or method: {method}"
        )

    # ------------------------------------------------------------------
    #  Binary, unary, and comparison operators
    # ------------------------------------------------------------------

    def _binop_sub(self, left: Any, right: Any) -> Any:
        return self._to_number(left) - self._to_number(right)

    def _binop_mul(self, left: Any, right: Any) -> Any:
        return self._to_number(left) * self._to_number(right)

    def _binop_div(self, left: Any, right: Any) -> Any:
        right_num = self._to_number(right)
        if right_num == 0:
            raise VBScriptError('Division by zero')
        return self._to_number(left) / right_num

    def _binop_intdiv(self, left: Any, right: Any) -> Any:
        right_num = self._to_number(right)
        if right_num == 0:
            raise VBScriptError('Division by zero')
        return int(self._to_number(left) / right_num)

    def _binop_mod(self, left: Any, right: Any) -> Any:
        right_num = self._to_number(right)
        if right_num == 0:
            raise VBScriptError('Division by zero')
        return int(math.fmod(self._to_number(left), right_num))

    def _binop_pow(self, left: Any, right: Any) -> Any:
        return self._to_number(left) ** self._to_number(right)

    def _binop_concat(self, left: Any, right: Any) -> Any:
        return self._to_string(left) + self._to_string(right)

    def _binop_xor(self, left: Any, right: Any) -> Any:
        return int(self._to_number(left)) ^ int(self._to_number(right))

    def _binop_eqv(self, left: Any, right: Any) -> Any:
        if _is_numeric_not_bool(left) and _is_numeric_not_bool(right):
            return ~(int(left) ^ int(right))
        return not (self._to_boolean(left) ^ self._to_boolean(right))

    def _binop_imp(self, left: Any, right: Any) -> Any:
        if _is_numeric_not_bool(left) and _is_numeric_not_bool(right):
            return (~int(left)) | int(right)
        return (not self._to_boolean(left)) or self._to_boolean(right)

    _BINOP_DISPATCH_NAMES = {
        BinaryOp.ADD: '_add',
        BinaryOp.SUB: '_binop_sub',
        BinaryOp.MUL: '_binop_mul',
        BinaryOp.DIV: '_binop_div',
        BinaryOp.INTDIV: '_binop_intdiv',
        BinaryOp.MOD: '_binop_mod',
        BinaryOp.POW: '_binop_pow',
        BinaryOp.CONCAT: '_binop_concat',
        BinaryOp.AND: '_logical_and',
        BinaryOp.OR: '_logical_or',
        BinaryOp.XOR: '_binop_xor',
        BinaryOp.EQV: '_binop_eqv',
        BinaryOp.IMP: '_binop_imp',
    }

    # Operators where both-numeric operands can use direct Python arithmetic,
    # skipping _to_number coercion and Empty/Null checks entirely.
    _NUMERIC_FAST = {
        BinaryOp.ADD: int.__add__,
        BinaryOp.SUB: int.__sub__,
        BinaryOp.MUL: int.__mul__,
        BinaryOp.POW: int.__pow__,
    }

    def _apply_binary_op(self, op: BinaryOp, left: Any, right: Any) -> Any:
        """Apply a binary operator."""
        # Fast path: both operands are int (not bool) -- skip Empty/Null checks
        # and coercion entirely.  Covers the hot arithmetic in For loops.
        if type(left) is int and type(right) is int:
            fast = self._NUMERIC_FAST.get(op)
            if fast is not None:
                return fast(left, right)
            if op is BinaryOp.DIV:
                if right == 0:
                    raise VBScriptError('Division by zero')
                return left / right
            if op is BinaryOp.INTDIV:
                if right == 0:
                    raise VBScriptError('Division by zero')
                return int(left / right)
            if op is BinaryOp.MOD:
                if right == 0:
                    raise VBScriptError('Division by zero')
                return int(math.fmod(left, right))

        # Handle Empty values
        if isinstance(left, VBScriptEmpty) and isinstance(right, VBScriptEmpty):
            left, right = 0, 0
        elif isinstance(left, VBScriptEmpty):
            left = 0 if isinstance(right, (int, float)) else ''
        elif isinstance(right, VBScriptEmpty):
            right = 0 if isinstance(left, (int, float)) else ''

        # Handle Null propagation (AND/OR have their own Null logic)
        if isinstance(left, VBScriptNull) or isinstance(right, VBScriptNull):
            if op is not BinaryOp.AND and op is not BinaryOp.OR:
                return NULL

        handler = self._binop_dispatch.get(op)
        if handler is not None:
            return handler(left, right)
        raise VBScriptError(f'Unknown binary operator: {op}')

    def _apply_unary_op(self, op: UnaryOp, operand: Any) -> Any:
        """Apply a unary operator."""
        if op == UnaryOp.NEG:
            return -self._to_number(operand)
        elif op == UnaryOp.POS:
            return self._to_number(operand)
        elif op == UnaryOp.NOT:
            if isinstance(operand, VBScriptNull):
                return NULL
            return not self._to_boolean(operand)
        else:
            raise VBScriptError(f'Unknown unary operator: {op}')

    def _apply_comparison_op(self, op: ComparisonOp, left: Any, right: Any) -> bool:
        """Apply a comparison operator."""
        # Handle Empty values
        if isinstance(left, VBScriptEmpty):
            left = 0 if isinstance(right, (int, float)) else ''
        if isinstance(right, VBScriptEmpty):
            right = 0 if isinstance(left, (int, float)) else ''

        # Handle Nothing comparison
        if op == ComparisonOp.IS:
            return left is right or (
                isinstance(left, VBScriptNothing) and isinstance(right, VBScriptNothing)
            )

        # Handle Null comparisons - three-valued logic: any comparison with Null returns Null
        if isinstance(left, VBScriptNull) or isinstance(right, VBScriptNull):
            return NULL

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
            raise VBScriptError(f'Unknown comparison operator: {op}')

    def _add(self, left: Any, right: Any) -> Any:
        """Handle addition with type coercion."""
        # In VBScript, + operator behavior depends on operand types:
        # - Both numbers: arithmetic addition
        # - Both strings: concatenation
        # - String and number: try to convert string to number and add
        #   (raises error if string is not numeric)
        # Use & operator for guaranteed string concatenation

        if isinstance(left, str) and isinstance(right, str):
            # Both strings - concatenate
            return left + right
        elif isinstance(left, str):
            # Left is string, right is not - try to convert string to number
            try:
                left_num = self._to_number(left)
                return left_num + self._to_number(right)
            except VBScriptError:
                raise VBScriptError('Type mismatch')
        elif isinstance(right, str):
            # Right is string, left is not - try to convert string to number
            try:
                right_num = self._to_number(right)
                return self._to_number(left) + right_num
            except VBScriptError:
                raise VBScriptError('Type mismatch')
        # Otherwise, numeric addition
        return self._to_number(left) + self._to_number(right)

    def _logical_and(self, left: Any, right: Any) -> Any:
        """Logical AND with VBScript semantics including Null propagation."""
        if isinstance(left, VBScriptNull):
            if isinstance(right, VBScriptNull):
                return NULL
            return False if not self._to_boolean(right) else NULL
        if isinstance(right, VBScriptNull):
            return False if not self._to_boolean(left) else NULL
        # VBScript uses bitwise AND for numbers
        if _is_numeric_not_bool(left) and _is_numeric_not_bool(right):
            return int(left) & int(right)
        return self._to_boolean(left) and self._to_boolean(right)

    def _logical_or(self, left: Any, right: Any) -> Any:
        """Logical OR with VBScript semantics including Null propagation."""
        if isinstance(left, VBScriptNull):
            if isinstance(right, VBScriptNull):
                return NULL
            return True if self._to_boolean(right) else NULL
        if isinstance(right, VBScriptNull):
            return True if self._to_boolean(left) else NULL
        # VBScript uses bitwise OR for numbers
        if _is_numeric_not_bool(left) and _is_numeric_not_bool(right):
            return int(left) | int(right)
        return self._to_boolean(left) or self._to_boolean(right)

    # ------------------------------------------------------------------
    #  Coercion helpers
    # ------------------------------------------------------------------

    def _to_number(self, value: Any) -> Union[int, float]:
        """Convert a value to a number."""
        if isinstance(value, int):
            if isinstance(value, bool):
                return -1 if value else 0
            return value
        if isinstance(value, float):
            return value
        if isinstance(value, VBScriptEmpty):
            return 0
        if isinstance(value, str):
            if value == '':
                return 0
            try:
                return float(value)
            except ValueError:
                raise VBScriptError(
                    f"Type mismatch: cannot convert '{value}' to number"
                )
        if isinstance(value, VBScriptNull):
            raise VBScriptError('Type mismatch: cannot convert Null to number')
        if isinstance(value, VBScriptNothing):
            raise VBScriptError('Type mismatch: cannot convert Nothing to number')
        raise VBScriptError('Type mismatch: cannot convert to number')

    def _to_string(self, value: Any) -> str:
        """Convert a value to a string."""
        if isinstance(value, str):
            return value
        if isinstance(value, int):
            if isinstance(value, bool):
                return 'True' if value else 'False'
            return str(value)
        if isinstance(value, float):
            if value.is_integer():
                return str(int(value))
            return str(value)
        if isinstance(value, VBScriptEmpty):
            return ''
        if isinstance(value, VBScriptNull):
            return 'Null'
        if isinstance(value, VBScriptNothing):
            return 'Nothing'
        return str(value)

    def _to_boolean(self, value: Any) -> bool:
        """Convert a value to a boolean."""
        if isinstance(value, bool):
            return value
        if isinstance(value, int):
            return value != 0
        if isinstance(value, float):
            return value != 0
        if isinstance(value, str):
            if value == '':
                return False
            if value.lower() == 'true':
                return True
            if value.lower() == 'false':
                return False
            return True
        if isinstance(value, VBScriptEmpty):
            return False
        if isinstance(value, VBScriptNull):
            return False
        if isinstance(value, VBScriptNothing):
            return False
        return bool(value)

def run(source: str, output_stream=None) -> Any:
    """Parse and execute VBScript source code."""
    from .parser import parse

    program = parse(source)
    interpreter = Interpreter(output_stream=output_stream)
    return interpreter.interpret(program)
