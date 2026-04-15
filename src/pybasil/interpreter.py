"""VBScript Tree-Walking Interpreter."""

from __future__ import annotations
import math
import sys
from typing import Any, Callable, Dict, List, Optional, Union
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
    ArrayAccess,
    DimStatement,
    AssignmentStatement,
    SetStatement,
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
    Parameter,
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


class VBScriptArray:
    """VBScript array implementation supporting multi-dimensional arrays."""

    def __init__(self, dimensions: List[int], is_dynamic: bool = False):
        """
        Initialize a VBScript array.

        Args:
            dimensions: List of upper bounds for each dimension (0-based)
                       e.g., [5] for arr(5), [2, 2] for arr(2, 2)
            is_dynamic: Whether this is a dynamic array (can be ReDim'd)
        """
        self._dimensions = dimensions
        self._is_dynamic = is_dynamic
        self._is_erased = False

        # Create the data structure
        if dimensions:
            self._data = self._create_array(dimensions)
        else:
            self._data = None  # Dynamic array not yet dimensioned

    def _create_array(self, dimensions: List[int]) -> list:
        """Recursively create a multi-dimensional array."""
        if len(dimensions) == 1:
            # Single dimension - create list with Empty values
            return [EMPTY for _ in range(dimensions[0] + 1)]
        else:
            # Multi-dimensional - create nested lists
            return [
                self._create_array(dimensions[1:]) for _ in range(dimensions[0] + 1)
            ]

    def get_element(self, indices: List[int]) -> Any:
        """Get an element by indices."""
        if self._is_erased:
            raise VBScriptError('Subscript out of range')
        if self._data is None:
            raise VBScriptError('Subscript out of range')

        # Navigate to the element
        current = self._data
        for i, idx in enumerate(indices):
            if not isinstance(current, list):
                raise VBScriptError('Subscript out of range')
            if idx < 0 or idx >= len(current):
                raise VBScriptError('Subscript out of range')
            current = current[idx]

        return current

    def set_element(self, indices: List[int], value: Any) -> None:
        """Set an element by indices."""
        if self._is_erased:
            raise VBScriptError('Subscript out of range')
        if self._data is None:
            raise VBScriptError('Subscript out of range')

        # Navigate to the parent of the element
        current = self._data
        for i, idx in enumerate(indices[:-1]):
            if not isinstance(current, list):
                raise VBScriptError('Subscript out of range')
            if idx < 0 or idx >= len(current):
                raise VBScriptError('Subscript out of range')
            current = current[idx]

        # Set the final element
        final_idx = indices[-1]
        if not isinstance(current, list):
            raise VBScriptError('Subscript out of range')
        if final_idx < 0 or final_idx >= len(current):
            raise VBScriptError('Subscript out of range')
        current[final_idx] = value

    def redim(self, dimensions: List[int], preserve: bool = False) -> None:
        """Resize the array, optionally preserving existing values."""
        if not self._is_dynamic:
            raise VBScriptError('This array is fixed or temporarily locked')

        old_data = self._data
        old_dimensions = self._dimensions

        self._dimensions = dimensions
        self._data = self._create_array(dimensions)
        self._is_erased = False

        if preserve and old_data is not None:
            # Copy existing values
            self._copy_data(old_data, old_dimensions, self._data, dimensions)

    def _copy_data(
        self, old_data: Any, old_dims: List[int], new_data: Any, new_dims: List[int]
    ) -> None:
        """Copy data from old array to new array during ReDim Preserve."""
        if len(old_dims) == 1:
            # Single dimension - copy elements
            min_len = min(len(old_data), len(new_data))
            for i in range(min_len):
                new_data[i] = old_data[i]
        else:
            # Multi-dimensional - recursively copy
            min_len = min(len(old_data), len(new_data))
            for i in range(min_len):
                self._copy_data(old_data[i], old_dims[1:], new_data[i], new_dims[1:])

    def erase(self) -> None:
        """Erase the array (deallocate dynamic arrays)."""
        if self._is_dynamic:
            self._data = None
            self._dimensions = []
            self._is_erased = True
        else:
            # Fixed-size array: reset all elements to Empty
            if self._data:
                self._reset_array(self._data)

    def _reset_array(self, data: Any) -> None:
        """Reset all elements of a fixed array to Empty."""
        if isinstance(data, list):
            for i, item in enumerate(data):
                if isinstance(item, list):
                    self._reset_array(item)
                else:
                    data[i] = EMPTY

    def ubound(self, dimension: int = 1) -> int:
        """Get the upper bound of a dimension (1-indexed)."""
        if self._is_erased or self._data is None:
            raise VBScriptError('Subscript out of range')
        if dimension < 1 or dimension > len(self._dimensions):
            raise VBScriptError('Subscript out of range')
        return self._dimensions[dimension - 1]

    def lbound(self, dimension: int = 1) -> int:
        """Get the lower bound of a dimension (always 0 in VBScript)."""
        if self._is_erased or self._data is None:
            raise VBScriptError('Subscript out of range')
        if dimension < 1 or dimension > len(self._dimensions):
            raise VBScriptError('Subscript out of range')
        return 0

    @property
    def dimensions(self) -> int:
        """Return the number of dimensions."""
        return len(self._dimensions)

    @property
    def is_erased(self) -> bool:
        """Check if the array has been erased."""
        return self._is_erased

    def __iter__(self):
        """Iterate over array elements (for For Each)."""
        if self._is_erased or self._data is None:
            return iter([])
        return self._iterate(self._data)

    def _iterate(self, data: Any):
        """Recursively iterate over array elements."""
        if isinstance(data, list):
            for item in data:
                yield from self._iterate(item)
        else:
            yield data


class _DictItemAccessor:
    """Helper class for dictionary Item property access."""

    def __init__(self, dictionary: 'VBScriptDictionary'):
        self._dict = dictionary

    def __call__(self, key: Any) -> Any:
        """Get item by key."""
        return self._dict.get_item(key)


class _DictKeyAccessor:
    """Helper class for dictionary Key property access."""

    def __init__(self, dictionary: 'VBScriptDictionary'):
        self._dict = dictionary

    def __call__(self, key: Any) -> Any:
        """Get key by key (returns the normalized key)."""
        return self._dict.get_key(key)


class VBScriptDictionary:
    """VBScript Scripting.Dictionary implementation."""

    def __init__(self):
        self._data: Dict[str, Any] = {}  # normalized key -> value
        self._key_order: List[str] = []  # normalized keys in insertion order
        self._original_keys: Dict[str, str] = {}  # normalized key -> original key
        self._compare_mode: int = (
            0  # 0 = binary (case-sensitive), 1 = text (case-insensitive)
        )

    def _normalize_key(self, key: Any) -> str:
        """Convert key to string and normalize based on compare mode."""
        if isinstance(key, str):
            if self._compare_mode == 1:  # Text mode - case insensitive
                return key.lower()
            return key
        else:
            return str(key)

    @property
    def Count(self) -> int:
        """Returns the number of key-item pairs."""
        return len(self._data)

    def _to_str_key(self, key: Any) -> str:
        """Convert key to its string representation (without normalization)."""
        return key if isinstance(key, str) else str(key)

    def Add(self, key: Any, item: Any) -> None:
        """Add a key-item pair to the dictionary."""
        norm_key = self._normalize_key(key)
        if norm_key in self._data:
            raise VBScriptError(
                'This key is already associated with an element of this collection'
            )
        self._data[norm_key] = item
        self._key_order.append(norm_key)
        self._original_keys[norm_key] = self._to_str_key(key)

    def Exists(self, key: Any) -> bool:
        """Returns True if the key exists in the dictionary."""
        norm_key = self._normalize_key(key)
        return norm_key in self._data

    def Items(self) -> VBScriptArray:
        """Returns an array containing all items."""
        items = [self._data[k] for k in self._key_order]
        if len(items) == 0:
            return VBScriptArray([-1], is_dynamic=True)
        arr = VBScriptArray([len(items) - 1], is_dynamic=False)
        for i, item in enumerate(items):
            arr.set_element([i], item)
        return arr

    def Keys(self) -> VBScriptArray:
        """Returns an array containing all keys."""
        keys = [self._original_keys.get(k, k) for k in self._key_order]
        if len(keys) == 0:
            return VBScriptArray([-1], is_dynamic=True)
        arr = VBScriptArray([len(keys) - 1], is_dynamic=False)
        for i, key in enumerate(keys):
            arr.set_element([i], key)
        return arr

    def Remove(self, key: Any) -> None:
        """Remove a key-item pair from the dictionary."""
        norm_key = self._normalize_key(key)
        if norm_key not in self._data:
            raise VBScriptError(
                'This key is not associated with an element of this collection'
            )
        del self._data[norm_key]
        self._key_order.remove(norm_key)
        self._original_keys.pop(norm_key, None)

    def RemoveAll(self) -> None:
        """Remove all key-item pairs from the dictionary."""
        self._data.clear()
        self._key_order.clear()
        self._original_keys.clear()

    @property
    def CompareMode(self) -> int:
        """Get or set the comparison mode (0=binary, 1=text)."""
        return self._compare_mode

    @CompareMode.setter
    def CompareMode(self, value: int):
        if len(self._data) > 0:
            raise VBScriptError('Invalid procedure call or argument')
        self._compare_mode = value

    def get_item(self, key: Any) -> Any:
        """Get an item by key (for default property access)."""
        norm_key = self._normalize_key(key)
        if norm_key not in self._data:
            # VBScript creates empty entry for non-existent key access
            self._data[norm_key] = EMPTY
            self._key_order.append(norm_key)
            self._original_keys[norm_key] = self._to_str_key(key)
            return EMPTY
        return self._data[norm_key]

    def set_item(self, key: Any, value: Any) -> None:
        """Set an item by key (for default property access)."""
        norm_key = self._normalize_key(key)
        if norm_key not in self._data:
            self._key_order.append(norm_key)
            self._original_keys[norm_key] = self._to_str_key(key)
        self._data[norm_key] = value

    def get_key(self, key: Any) -> Any:
        """Get the key value (for Key property)."""
        norm_key = self._normalize_key(key)
        if norm_key not in self._data:
            raise VBScriptError(
                'This key is not associated with an element of this collection'
            )
        return self._original_keys.get(norm_key, norm_key)

    def set_key(self, old_key: Any, new_key: Any) -> None:
        """Change a key value."""
        norm_old = self._normalize_key(old_key)
        norm_new = self._normalize_key(new_key)

        if norm_old not in self._data:
            raise VBScriptError(
                'This key is not associated with an element of this collection'
            )
        if norm_new in self._data and norm_new != norm_old:
            raise VBScriptError(
                'This key is already associated with an element of this collection'
            )

        # Move the item to the new key
        item = self._data[norm_old]
        del self._data[norm_old]
        self._data[norm_new] = item

        # Update key order and original key mapping
        idx = self._key_order.index(norm_old)
        self._key_order[idx] = norm_new
        self._original_keys.pop(norm_old, None)
        self._original_keys[norm_new] = self._to_str_key(new_key)

    def __iter__(self):
        """Iterate over keys (for For Each)."""
        for key in self._key_order:
            yield self._original_keys.get(key, key)


class ErrObject:
    """VBScript Err object for error information."""

    def __init__(self):
        self._number: int = 0
        self._source: str = ''
        self._description: str = ''
        self._helpfile: str = ''
        self._helpcontext: int = 0

    @property
    def Number(self) -> int:
        """Error number (default property)."""
        return self._number

    @Number.setter
    def Number(self, value: int):
        self._number = value

    @property
    def Source(self) -> str:
        """Source of the error."""
        return self._source

    @Source.setter
    def Source(self, value: str):
        self._source = value

    @property
    def Description(self) -> str:
        """Error description."""
        return self._description

    @Description.setter
    def Description(self, value: str):
        self._description = value

    @property
    def HelpFile(self) -> str:
        """Help file path."""
        return self._helpfile

    @HelpFile.setter
    def HelpFile(self, value: str):
        self._helpfile = value

    @property
    def HelpContext(self) -> int:
        """Help context ID."""
        return self._helpcontext

    @HelpContext.setter
    def HelpContext(self, value: int):
        self._helpcontext = value

    def Clear(self):
        """Clear the error information."""
        self._number = 0
        self._source = ''
        self._description = ''
        self._helpfile = ''
        self._helpcontext = 0

    def Raise(
        self,
        number: int,
        source: str = '',
        description: str = '',
        helpfile: str = '',
        helpcontext: int = 0,
    ):
        """Raise a runtime error."""
        self._number = number
        self._source = source
        self._description = description
        self._helpfile = helpfile
        self._helpcontext = helpcontext
        raise VBScriptError(
            f'Error {number}: {description}' if description else f'Error {number}'
        )


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
        output = ' '.join(output_parts)
        print(output, file=self._output)

    def _format_value(self, value: Any) -> str:
        """Format a value for output."""
        if value is True:
            return 'True'
        elif value is False:
            return 'False'
        elif isinstance(value, VBScriptNothing):
            return 'Nothing'
        elif isinstance(value, VBScriptEmpty):
            return ''
        elif isinstance(value, VBScriptNull):
            return 'Null'
        elif isinstance(value, VBScriptArray):
            return 'Variant()'
        elif value is None:
            return 'Nothing'
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
        self._error_mode: ErrorHandlingMode = ErrorHandlingMode.DEFAULT
        self._err: ErrObject = ErrObject()
        self._setup_builtins()

    def _setup_builtins(self) -> None:
        """Set up built-in objects and functions."""
        # Create WScript object
        wscript = WScriptObject(self._output_stream)
        self._environment.define('WScript', wscript)

        # Create Err object
        self._environment.define('Err', self._err)

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
            'ubound': self._builtin_ubound,
            'lbound': self._builtin_lbound,
            'array': self._builtin_array,
        }

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
        return 1000  # Generic runtime error

    def _execute(self, node: ASTNode) -> Any:
        """Execute a statement node."""
        method_name = f'_execute_{type(node).__name__}'
        method = getattr(self, method_name, self._execute_default)
        return method(node)

    def _execute_default(self, node: ASTNode) -> Any:
        """Default execution handler."""
        raise VBScriptError(f'Unknown statement type: {type(node).__name__}')

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
            if name in self._procedures:
                return self._call_procedure(
                    node.expression.name, node.expression.arguments
                )

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
        if node.exit_type in (ExitType.SUB, ExitType.FUNCTION):
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

    def _evaluate(self, node: ASTNode) -> Any:
        """Evaluate an expression node."""
        method_name = f'_evaluate_{type(node).__name__}'
        method = getattr(self, method_name, self._evaluate_default)
        return method(node)

    def _evaluate_default(self, node: ASTNode) -> Any:
        """Default evaluation handler."""
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

        # Check built-in functions
        if func_name in self._builtins:
            args = [self._evaluate(arg) for arg in node.arguments]
            return self._builtins[func_name](*args)

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
        raise VBScriptError(
            f'CreateObject should be used instead of New for: {node.class_name}'
        )

    def _apply_binary_op(self, op: BinaryOp, left: Any, right: Any) -> Any:
        """Apply a binary operator."""
        # Handle Empty values
        if isinstance(left, VBScriptEmpty) and isinstance(right, VBScriptEmpty):
            left, right = 0, 0
        elif isinstance(left, VBScriptEmpty):
            left = 0 if isinstance(right, (int, float)) else ''
        elif isinstance(right, VBScriptEmpty):
            right = 0 if isinstance(left, (int, float)) else ''

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
                raise VBScriptError('Division by zero')
            return self._to_number(left) / right_num
        elif op == BinaryOp.INTDIV:
            right_num = self._to_number(right)
            if right_num == 0:
                raise VBScriptError('Division by zero')
            return int(self._to_number(left) / right_num)
        elif op == BinaryOp.MOD:
            right_num = self._to_number(right)
            if right_num == 0:
                raise VBScriptError('Division by zero')
            return int(math.fmod(self._to_number(left), right_num))
        elif op == BinaryOp.POW:
            return self._to_number(left) ** self._to_number(right)
        elif op == BinaryOp.CONCAT:
            return self._to_string(left) + self._to_string(right)
        elif op == BinaryOp.AND:
            return self._logical_and(left, right)
        elif op == BinaryOp.OR:
            return self._logical_or(left, right)
        elif op == BinaryOp.XOR:
            return int(self._to_number(left)) ^ int(self._to_number(right))
        elif op == BinaryOp.EQV:
            if isinstance(left, (int, float)) and isinstance(right, (int, float)) and not isinstance(left, bool) and not isinstance(right, bool):
                return ~(int(left) ^ int(right))
            return not (self._to_boolean(left) ^ self._to_boolean(right))
        elif op == BinaryOp.IMP:
            if isinstance(left, (int, float)) and isinstance(right, (int, float)) and not isinstance(left, bool) and not isinstance(right, bool):
                return (~int(left)) | int(right)
            return (not self._to_boolean(left)) or self._to_boolean(right)
        else:
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
        if isinstance(left, (int, float)) and isinstance(right, (int, float)) and not isinstance(left, bool) and not isinstance(right, bool):
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
        if isinstance(left, (int, float)) and isinstance(right, (int, float)) and not isinstance(left, bool) and not isinstance(right, bool):
            return int(left) | int(right)
        return self._to_boolean(left) or self._to_boolean(right)

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

    # Built-in functions
    def _builtin_msgbox(self, *args) -> int:
        """MsgBox function (simplified)."""
        if args:
            print(self._to_string(args[0]))
        return 1  # vbOK

    def _builtin_inputbox(self, prompt: str, title: str = '', default: str = '') -> str:
        """InputBox function (simplified)."""
        return default

    def _builtin_len(self, value: Any) -> int:
        """Len function."""
        if isinstance(value, str):
            return len(value)
        raise VBScriptError('Type mismatch: Len requires a string')

    def _builtin_left(self, string: str, length: int) -> str:
        """Left function."""
        n = int(length)
        if n < 0:
            raise VBScriptError('Invalid procedure call or argument')
        return self._to_string(string)[:n]

    def _builtin_right(self, string: str, length: int) -> str:
        """Right function."""
        s = self._to_string(string)
        n = int(length)
        if n < 0:
            raise VBScriptError('Invalid procedure call or argument')
        return s[-n:] if n > 0 else ''

    def _builtin_mid(self, string: str, start: int, length: int = None) -> str:
        """Mid function."""
        s = self._to_string(string)
        start_val = int(start)
        if start_val < 1:
            raise VBScriptError('Invalid procedure call or argument')
        start_idx = start_val - 1  # VBScript is 1-indexed
        if length is None:
            return s[start_idx:]
        return s[start_idx : start_idx + int(length)]

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
            start = 1
            string1, string2 = args
            compare = 0
        elif len(args) == 3:
            start = int(args[0])
            string1, string2 = args[1], args[2]
            compare = 0
        elif len(args) >= 4:
            start = int(args[0])
            string1, string2 = args[1], args[2]
            compare = int(args[3])
        else:
            return 0

        s1 = string1 if isinstance(string1, str) else self._to_string(string1)
        s2 = string2 if isinstance(string2, str) else self._to_string(string2)
        start_idx = start - 1  # VBScript is 1-indexed
        if compare == 1:
            idx = s1.lower().find(s2.lower(), start_idx)
        else:
            idx = s1.find(s2, start_idx)
        return idx + 1 if idx >= 0 else 0

    def _builtin_replace(
        self,
        string: str,
        find: str,
        replace_with: str,
        start: int = 1,
        count: int = -1,
        compare: int = 0,
    ) -> str:
        """Replace function."""
        s = self._to_string(string)
        f = self._to_string(find)
        r = self._to_string(replace_with)
        start_val = int(start)
        count_val = int(count)
        compare_val = int(compare)

        # VBScript Replace returns the substring starting at 'start'
        s = s[start_val - 1:]

        if compare_val == 1:
            # Case-insensitive replace
            result = []
            lower_s = s.lower()
            lower_f = f.lower()
            i = 0
            replacements = 0
            while i < len(s):
                if lower_s[i:i + len(lower_f)] == lower_f and (count_val < 0 or replacements < count_val):
                    result.append(r)
                    i += len(f)
                    replacements += 1
                else:
                    result.append(s[i])
                    i += 1
            return ''.join(result)
        else:
            if count_val >= 0:
                return s.replace(f, r, count_val)
            else:
                return s.replace(f, r)

    def _builtin_split(
        self, string: str, delimiter: str = ' ', count: int = -1, compare: int = 0
    ) -> VBScriptArray:
        """Split function."""
        s = self._to_string(string)
        d = self._to_string(delimiter)
        if count > 0:
            parts = s.split(d, count - 1)
        else:
            parts = s.split(d)
        arr = VBScriptArray([len(parts) - 1], is_dynamic=True)
        for i, part in enumerate(parts):
            arr.set_element([i], part)
        return arr

    def _builtin_join(self, array: list, delimiter: str = ' ') -> str:
        """Join function."""
        d = self._to_string(delimiter)
        return d.join(self._to_string(item) for item in array)

    def _builtin_cstr(self, value: Any) -> str:
        """CStr function."""
        return self._to_string(value)

    def _builtin_cint(self, value: Any) -> int:
        """CInt function."""
        if isinstance(value, int) and not isinstance(value, bool):
            return value
        return round(self._to_number(value))

    def _builtin_clng(self, value: Any) -> int:
        """CLng function."""
        if isinstance(value, int) and not isinstance(value, bool):
            return value
        return round(self._to_number(value))

    def _builtin_cdbl(self, value: Any) -> float:
        """CDbl function."""
        if isinstance(value, float):
            return value
        if isinstance(value, int) and not isinstance(value, bool):
            return float(value)
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
        return isinstance(value, (VBScriptArray, list, tuple))

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
        return (
            isinstance(value, (WScriptObject, VBScriptObject))
            or value is not None
            and not isinstance(
                value,
                (str, int, float, bool, VBScriptEmpty, VBScriptNull, VBScriptNothing),
            )
        )

    def _builtin_typename(self, value: Any) -> str:
        """TypeName function."""
        if isinstance(value, VBScriptEmpty):
            return 'Empty'
        if isinstance(value, VBScriptNull):
            return 'Null'
        if isinstance(value, VBScriptNothing):
            return 'Nothing'
        if isinstance(value, VBScriptArray):
            return 'Variant()'
        if isinstance(value, bool):
            return 'Boolean'
        if isinstance(value, int):
            return 'Integer'
        if isinstance(value, float):
            return 'Double'
        if isinstance(value, str):
            return 'String'
        if isinstance(value, (list, tuple)):
            return 'Variant()'
        return 'Object'

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
        if isinstance(value, (VBScriptArray, list, tuple)):
            return 8204  # vbArray + vbVariant
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
        """CreateObject function - creates COM objects (simplified)."""
        # Handle known object types
        class_lower = class_name.lower()

        if class_lower == 'scripting.dictionary':
            return VBScriptDictionary()

        # For unknown objects, return a placeholder
        return {'_class': class_name}

    def _builtin_getobject(self, path_name: str = None, class_name: str = None) -> Any:
        """GetObject function (simplified)."""
        # This is a stub - in real VBScript this would get COM objects
        return {'_path': path_name, '_class': class_name}

    def _builtin_ubound(self, array: Any, dimension: int = 1) -> int:
        """UBound function - returns the upper bound of an array dimension."""
        if isinstance(array, VBScriptArray):
            return array.ubound(dimension)
        elif isinstance(array, list):
            if dimension != 1:
                raise VBScriptError('Subscript out of range')
            return len(array) - 1
        else:
            raise VBScriptError('Type mismatch: UBound requires an array')

    def _builtin_lbound(self, array: Any, dimension: int = 1) -> int:
        """LBound function - returns the lower bound of an array dimension."""
        if isinstance(array, VBScriptArray):
            return array.lbound(dimension)
        elif isinstance(array, list):
            if dimension != 1:
                raise VBScriptError('Subscript out of range')
            return 0
        else:
            raise VBScriptError('Type mismatch: LBound requires an array')

    def _builtin_array(self, *args) -> VBScriptArray:
        """Array function - creates a variant array from the given values."""
        if len(args) == 0:
            # Empty array (UBound = -1)
            arr = VBScriptArray([-1], is_dynamic=True)
            return arr

        # Create array with upper bound = len(args) - 1
        arr = VBScriptArray([len(args) - 1], is_dynamic=False)
        for i, val in enumerate(args):
            arr.set_element([i], val)
        return arr


def run(source: str, output_stream=None) -> Any:
    """Parse and execute VBScript source code."""
    from .parser import parse

    program = parse(source)
    interpreter = Interpreter(output_stream=output_stream)
    return interpreter.interpret(program)
