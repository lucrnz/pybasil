"""VBScript runtime value types, environment, and control-flow exceptions."""

from __future__ import annotations
import sys
from typing import Any, Dict, List, Optional
from dataclasses import dataclass

from .ast_nodes import ExitType, Parameter, ASTNode


# ---------------------------------------------------------------------------
#  Sentinel singletons
# ---------------------------------------------------------------------------

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

# Sentinel used by the inlined scope-walk in _evaluate_Identifier to
# distinguish "key absent" from "value is None/Empty".  See the comment
# in interpreter.py for why the lookup is inlined.
_NOT_FOUND = object()


# ---------------------------------------------------------------------------
#  VBScriptArray
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
#  VBScriptDictionary
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
#  ErrObject
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
#  WScriptObject
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
#  Control-flow exceptions
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
#  Environment & Procedure
# ---------------------------------------------------------------------------

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
