"""VBScript runtime value types, environment, and control-flow exceptions."""

from __future__ import annotations
import os
import sys
import shutil
import tempfile
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional
from dataclasses import dataclass

from ..ast_nodes import ExitType, Parameter, ASTNode


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
#  VBScriptDate  (OLE Automation date)
# ---------------------------------------------------------------------------

# OLE Automation epoch: 1899-12-30 00:00:00
_OLE_EPOCH = datetime(1899, 12, 30)
_SECONDS_PER_DAY = 86400.0


class VBScriptDate:
    """VBScript Date value backed by an OLE Automation serial number.

    The integer part counts days from 1899-12-30 (day 0).
    The fractional part represents the time of day.
    """

    __slots__ = ('_serial',)

    def __init__(self, serial: float = 0.0):
        self._serial = serial

    # -- construction helpers ------------------------------------------------

    @classmethod
    def from_datetime(cls, dt: datetime) -> 'VBScriptDate':
        delta = dt - _OLE_EPOCH
        serial = delta.days + delta.seconds / _SECONDS_PER_DAY
        return cls(serial)

    @classmethod
    def from_date_parts(cls, year: int, month: int, day: int) -> 'VBScriptDate':
        dt = datetime(year, month, day)
        return cls.from_datetime(dt)

    @classmethod
    def from_time_parts(cls, hour: int, minute: int, second: int) -> 'VBScriptDate':
        serial = (hour * 3600 + minute * 60 + second) / _SECONDS_PER_DAY
        return cls(serial)

    @classmethod
    def from_string(cls, s: str) -> 'VBScriptDate':
        """Parse common date/time string formats."""
        s = s.strip()
        if not s:
            raise VBScriptError('Type mismatch')

        # Try datetime formats
        for fmt in (
            '%m/%d/%Y %I:%M:%S %p',
            '%m/%d/%Y %H:%M:%S',
            '%m/%d/%Y %H:%M',
            '%m/%d/%Y',
            '%Y-%m-%d %H:%M:%S',
            '%Y-%m-%d',
            '%m-%d-%Y',
            '%d-%b-%Y',
            '%d-%b-%y',
            '%B %d, %Y',
            '%b %d, %Y',
            '%I:%M:%S %p',
            '%H:%M:%S',
            '%H:%M',
        ):
            try:
                dt = datetime.strptime(s, fmt)
                return cls.from_datetime(dt)
            except ValueError:
                continue

        raise VBScriptError(f'Type mismatch: cannot convert \'{s}\' to Date')

    # -- conversion ----------------------------------------------------------

    def to_datetime(self) -> datetime:
        serial = self._serial
        days = int(serial)
        frac = abs(serial - days)
        total_seconds = round(frac * _SECONDS_PER_DAY)
        return _OLE_EPOCH + timedelta(days=days, seconds=total_seconds)

    @property
    def serial(self) -> float:
        return self._serial

    # -- component accessors -------------------------------------------------

    @property
    def year(self) -> int:
        return self.to_datetime().year

    @property
    def month(self) -> int:
        return self.to_datetime().month

    @property
    def day(self) -> int:
        return self.to_datetime().day

    @property
    def hour(self) -> int:
        return self.to_datetime().hour

    @property
    def minute(self) -> int:
        return self.to_datetime().minute

    @property
    def second(self) -> int:
        return self.to_datetime().second

    @property
    def weekday(self) -> int:
        """VBScript weekday: 1=Sunday, 2=Monday, ..., 7=Saturday."""
        # Python: Monday=0 ... Sunday=6
        py_wd = self.to_datetime().weekday()
        return (py_wd + 2) % 7 or 7

    # -- formatting ----------------------------------------------------------

    def _format_date(self, dt: datetime) -> str:
        return f'{dt.month}/{dt.day}/{dt.year}'

    def _format_time(self, dt: datetime) -> str:
        h = dt.hour % 12 or 12
        ampm = 'AM' if dt.hour < 12 else 'PM'
        return f'{h}:{dt.minute:02d}:{dt.second:02d} {ampm}'

    def __str__(self) -> str:
        dt = self.to_datetime()
        has_date = int(self._serial) != 0
        frac = round((self._serial % 1) * _SECONDS_PER_DAY)
        has_time = frac != 0
        if has_date and has_time:
            return f'{self._format_date(dt)} {self._format_time(dt)}'
        elif has_time:
            return self._format_time(dt)
        else:
            return self._format_date(dt)

    def __repr__(self) -> str:
        return f'VBScriptDate({self._serial})'

    # -- comparison / arithmetic (used by interpreter operators) -------------

    def __eq__(self, other: object) -> bool:
        if isinstance(other, VBScriptDate):
            return self._serial == other._serial
        return NotImplemented

    def __lt__(self, other: 'VBScriptDate') -> bool:
        if isinstance(other, VBScriptDate):
            return self._serial < other._serial
        return NotImplemented

    def __le__(self, other: 'VBScriptDate') -> bool:
        if isinstance(other, VBScriptDate):
            return self._serial <= other._serial
        return NotImplemented

    def __gt__(self, other: 'VBScriptDate') -> bool:
        if isinstance(other, VBScriptDate):
            return self._serial > other._serial
        return NotImplemented

    def __ge__(self, other: 'VBScriptDate') -> bool:
        if isinstance(other, VBScriptDate):
            return self._serial >= other._serial
        return NotImplemented

    def __hash__(self) -> int:
        return hash(self._serial)


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
#  Scripting.FileSystemObject
# ---------------------------------------------------------------------------

# I/O mode constants
_FOR_READING = 1
_FOR_WRITING = 2
_FOR_APPENDING = 8

# Tristate constants
_TRISTATE_TRUE = -1      # Unicode
_TRISTATE_FALSE = 0      # ASCII
_TRISTATE_USE_DEFAULT = -2

# SpecialFolder constants
_WINDOWS_FOLDER = 0
_SYSTEM_FOLDER = 1
_TEMP_FOLDER = 2

# File attribute constants
_ATTR_NORMAL = 0
_ATTR_READONLY = 1
_ATTR_HIDDEN = 2
_ATTR_SYSTEM = 4
_ATTR_DIRECTORY = 16
_ATTR_ARCHIVE = 32

# File.Type descriptions matching cscript.exe / Windows Shell
_FILE_TYPE_MAP: Dict[str, str] = {
    '.txt': 'Text Document',
    '.log': 'Text Document',
    '.ini': 'Configuration settings',
    '.cfg': 'Configuration Source File',
    '.csv': 'Comma Separated Values Source File',
    '.xml': 'XML Source File',
    '.json': 'JSON Source File',
    '.yaml': 'Yaml Source File',
    '.yml': 'Yaml Source File',
    '.sql': 'SQL Source File',
    '.md': 'Markdown Source File',
    '.rtf': 'RTF File',
    '.htm': 'HTML Document',
    '.html': 'HTML Document',
    '.bat': 'Windows Batch File',
    '.cmd': 'Windows Command Script',
    '.vbs': 'VBScript Script File',
    '.js': 'JSFile',
    '.exe': 'Application',
    '.dll': 'Application extension',
    '.sys': 'System file',
    '.reg': 'Registration Entries',
    '.zip': 'Compressed (zipped) Folder',
    '.py': 'Python Source File',
    '.rb': 'Ruby Source File',
    '.pl': 'Perl Source File',
    '.sh': 'SH Source File',
    '.c': 'C Source File',
    '.cpp': 'C++ Source File',
    '.h': 'C Header Source File',
    '.cs': 'C# Source File',
    '.java': 'Java Source File',
}


def _file_type_description(ext: str) -> str:
    if not ext:
        return 'File'
    desc = _FILE_TYPE_MAP.get(ext.lower())
    if desc:
        return desc
    return ext.lstrip('.').upper() + ' File'


class VBScriptTextStream(VBScriptObject):
    """TextStream object for reading/writing text files."""

    def __init__(self, file_handle, mode: int):
        self._handle = file_handle
        self._mode = mode
        self._line = 1
        self._column = 1
        self._closed = False

    def _check_open(self) -> None:
        if self._closed:
            raise VBScriptError('Bad file mode')

    # -- Reading -------------------------------------------------------------

    def Read(self, characters: int) -> str:
        self._check_open()
        if self._mode != _FOR_READING:
            raise VBScriptError('Bad file mode')
        data = self._handle.read(int(characters))
        self._update_position(data)
        return data

    def ReadLine(self) -> str:
        self._check_open()
        if self._mode != _FOR_READING:
            raise VBScriptError('Bad file mode')
        line = self._handle.readline()
        if line.endswith('\n'):
            line = line[:-1]
            if line.endswith('\r'):
                line = line[:-1]
        self._line += 1
        self._column = 1
        return line

    def ReadAll(self) -> str:
        self._check_open()
        if self._mode != _FOR_READING:
            raise VBScriptError('Bad file mode')
        data = self._handle.read()
        self._update_position(data)
        return data

    # -- Writing -------------------------------------------------------------

    def Write(self, text: str) -> None:
        self._check_open()
        if self._mode == _FOR_READING:
            raise VBScriptError('Bad file mode')
        s = str(text) if not isinstance(text, str) else text
        self._handle.write(s)
        self._update_position(s)

    def WriteLine(self, text: str = '') -> None:
        self._check_open()
        if self._mode == _FOR_READING:
            raise VBScriptError('Bad file mode')
        s = str(text) if not isinstance(text, str) else text
        self._handle.write(s + '\r\n')
        self._line += 1
        self._column = 1

    def WriteBlankLines(self, lines: int) -> None:
        self._check_open()
        if self._mode == _FOR_READING:
            raise VBScriptError('Bad file mode')
        for _ in range(int(lines)):
            self._handle.write('\r\n')
        self._line += int(lines)
        self._column = 1

    def Close(self) -> None:
        if not self._closed:
            self._handle.close()
            self._closed = True

    # -- Properties ----------------------------------------------------------

    @property
    def AtEndOfStream(self) -> bool:
        self._check_open()
        if self._mode != _FOR_READING:
            raise VBScriptError('Bad file mode')
        pos = self._handle.tell()
        ch = self._handle.read(1)
        if ch == '':
            return True
        self._handle.seek(pos)
        return False

    @property
    def AtEndOfLine(self) -> bool:
        self._check_open()
        if self._mode != _FOR_READING:
            raise VBScriptError('Bad file mode')
        pos = self._handle.tell()
        ch = self._handle.read(1)
        if ch == '':
            return True
        self._handle.seek(pos)
        return ch in ('\n', '\r')

    @property
    def Line(self) -> int:
        return self._line

    @property
    def Column(self) -> int:
        return self._column

    def _update_position(self, text: str) -> None:
        for ch in text:
            if ch == '\n':
                self._line += 1
                self._column = 1
            else:
                self._column += 1


class VBScriptFile(VBScriptObject):
    """File object representing a file on disk."""

    def __init__(self, path: str):
        self._path = os.path.abspath(path)

    @property
    def Name(self) -> str:
        return os.path.basename(self._path)

    @Name.setter
    def Name(self, value: str) -> None:
        new_path = os.path.join(os.path.dirname(self._path), value)
        os.rename(self._path, new_path)
        self._path = new_path

    @property
    def Path(self) -> str:
        return self._path

    @property
    def ShortPath(self) -> str:
        return self._path

    @property
    def ShortName(self) -> str:
        return self.Name

    @property
    def Size(self) -> int:
        return os.path.getsize(self._path)

    @property
    def Type(self) -> str:
        _, ext = os.path.splitext(self._path)
        return _file_type_description(ext)

    @property
    def DateCreated(self) -> 'VBScriptDate':
        ts = os.path.getctime(self._path)
        return VBScriptDate.from_datetime(datetime.fromtimestamp(ts))

    @property
    def DateLastModified(self) -> 'VBScriptDate':
        ts = os.path.getmtime(self._path)
        return VBScriptDate.from_datetime(datetime.fromtimestamp(ts))

    @property
    def DateLastAccessed(self) -> 'VBScriptDate':
        ts = os.path.getatime(self._path)
        return VBScriptDate.from_datetime(datetime.fromtimestamp(ts))

    @property
    def Attributes(self) -> int:
        attrs = _ATTR_NORMAL
        if not os.access(self._path, os.W_OK):
            attrs |= _ATTR_READONLY
        return attrs

    @property
    def ParentFolder(self) -> 'VBScriptFolder':
        return VBScriptFolder(os.path.dirname(self._path))

    def Delete(self, force: bool = False) -> None:
        os.remove(self._path)

    def Copy(self, destination: str, overwrite: bool = True) -> None:
        dest = str(destination)
        if os.path.isdir(dest):
            dest = os.path.join(dest, self.Name)
        if not overwrite and os.path.exists(dest):
            raise VBScriptError('File already exists')
        shutil.copy2(self._path, dest)

    def Move(self, destination: str) -> None:
        dest = str(destination)
        if os.path.isdir(dest):
            dest = os.path.join(dest, self.Name)
        shutil.move(self._path, dest)
        self._path = os.path.abspath(dest)

    def OpenAsTextStream(self, iomode: int = _FOR_READING, _format: int = _TRISTATE_FALSE) -> VBScriptTextStream:
        if iomode == _FOR_READING:
            fh = open(self._path, 'r', encoding='utf-8', newline='')
        elif iomode == _FOR_WRITING:
            fh = open(self._path, 'w', encoding='utf-8', newline='')
        elif iomode == _FOR_APPENDING:
            fh = open(self._path, 'a', encoding='utf-8', newline='')
        else:
            raise VBScriptError('Bad file mode')
        return VBScriptTextStream(fh, iomode)


class VBScriptFolder(VBScriptObject):
    """Folder object representing a directory on disk."""

    def __init__(self, path: str):
        self._path = os.path.abspath(path)

    @property
    def Name(self) -> str:
        return os.path.basename(self._path) or self._path

    @Name.setter
    def Name(self, value: str) -> None:
        new_path = os.path.join(os.path.dirname(self._path), value)
        os.rename(self._path, new_path)
        self._path = new_path

    @property
    def Path(self) -> str:
        return self._path

    @property
    def ShortPath(self) -> str:
        return self._path

    @property
    def ShortName(self) -> str:
        return self.Name

    @property
    def Size(self) -> int:
        total = 0
        for dirpath, _dirnames, filenames in os.walk(self._path):
            for f in filenames:
                fp = os.path.join(dirpath, f)
                try:
                    total += os.path.getsize(fp)
                except OSError:
                    pass
        return total

    @property
    def Type(self) -> str:
        return 'File folder'

    @property
    def DateCreated(self) -> 'VBScriptDate':
        ts = os.path.getctime(self._path)
        return VBScriptDate.from_datetime(datetime.fromtimestamp(ts))

    @property
    def DateLastModified(self) -> 'VBScriptDate':
        ts = os.path.getmtime(self._path)
        return VBScriptDate.from_datetime(datetime.fromtimestamp(ts))

    @property
    def DateLastAccessed(self) -> 'VBScriptDate':
        ts = os.path.getatime(self._path)
        return VBScriptDate.from_datetime(datetime.fromtimestamp(ts))

    @property
    def Attributes(self) -> int:
        return _ATTR_DIRECTORY

    @property
    def IsRootFolder(self) -> bool:
        return os.path.dirname(self._path) == self._path

    @property
    def ParentFolder(self) -> 'VBScriptFolder':
        parent = os.path.dirname(self._path)
        if parent == self._path:
            raise VBScriptError('Path not found')
        return VBScriptFolder(parent)

    @property
    def SubFolders(self) -> 'VBScriptFolderCollection':
        return VBScriptFolderCollection(self._path)

    @property
    def Files(self) -> 'VBScriptFileCollection':
        return VBScriptFileCollection(self._path)

    def Delete(self, force: bool = False) -> None:
        shutil.rmtree(self._path)

    def Copy(self, destination: str, overwrite: bool = True) -> None:
        dest = str(destination)
        if os.path.exists(dest) and not overwrite:
            raise VBScriptError('File already exists')
        shutil.copytree(self._path, dest, dirs_exist_ok=overwrite)

    def Move(self, destination: str) -> None:
        dest = str(destination)
        shutil.move(self._path, dest)
        self._path = os.path.abspath(dest)

    def CreateTextFile(self, filename: str, overwrite: bool = True) -> VBScriptTextStream:
        path = os.path.join(self._path, filename)
        if not overwrite and os.path.exists(path):
            raise VBScriptError('File already exists')
        fh = open(path, 'w', encoding='utf-8', newline='')
        return VBScriptTextStream(fh, _FOR_WRITING)


class VBScriptFileCollection(VBScriptObject):
    """Collection of File objects in a folder."""

    def __init__(self, folder_path: str):
        self._folder_path = folder_path

    @property
    def Count(self) -> int:
        try:
            return sum(1 for e in os.scandir(self._folder_path) if e.is_file())
        except OSError:
            return 0

    def Item(self, name: str) -> VBScriptFile:
        path = os.path.join(self._folder_path, name)
        if not os.path.isfile(path):
            raise VBScriptError('File not found')
        return VBScriptFile(path)

    def __iter__(self):
        try:
            for entry in os.scandir(self._folder_path):
                if entry.is_file():
                    yield VBScriptFile(entry.path)
        except OSError:
            return


class VBScriptFolderCollection(VBScriptObject):
    """Collection of Folder objects (subfolders)."""

    def __init__(self, folder_path: str):
        self._folder_path = folder_path

    @property
    def Count(self) -> int:
        try:
            return sum(1 for e in os.scandir(self._folder_path) if e.is_dir())
        except OSError:
            return 0

    def Item(self, name: str) -> VBScriptFolder:
        path = os.path.join(self._folder_path, name)
        if not os.path.isdir(path):
            raise VBScriptError('Path not found')
        return VBScriptFolder(path)

    def __iter__(self):
        try:
            for entry in os.scandir(self._folder_path):
                if entry.is_dir():
                    yield VBScriptFolder(entry.path)
        except OSError:
            return


class VBScriptDrive(VBScriptObject):
    """Drive object."""

    def __init__(self, path: str):
        self._path = path

    @property
    def DriveLetter(self) -> str:
        if len(self._path) >= 1 and self._path[1:2] == ':':
            return self._path[0].upper()
        return ''

    @property
    def Path(self) -> str:
        return self._path

    @property
    def RootFolder(self) -> VBScriptFolder:
        return VBScriptFolder(self._path)

    @property
    def DriveType(self) -> int:
        return 2  # Fixed

    @property
    def IsReady(self) -> bool:
        return os.path.exists(self._path)

    @property
    def FileSystem(self) -> str:
        return 'Unknown'

    @property
    def TotalSize(self) -> int:
        try:
            usage = shutil.disk_usage(self._path)
            return usage.total
        except OSError:
            return 0

    @property
    def AvailableSpace(self) -> int:
        try:
            usage = shutil.disk_usage(self._path)
            return usage.free
        except OSError:
            return 0

    @property
    def FreeSpace(self) -> int:
        return self.AvailableSpace

    @property
    def VolumeName(self) -> str:
        return ''

    @property
    def SerialNumber(self) -> int:
        return 0


class VBScriptDriveCollection(VBScriptObject):
    """Collection of Drive objects."""

    def __init__(self):
        pass

    @property
    def Count(self) -> int:
        return len(self._get_drives())

    def Item(self, spec: str) -> VBScriptDrive:
        s = str(spec).rstrip(':').rstrip('\\').rstrip('/')
        if len(s) == 1:
            s = s.upper() + ':' + os.sep
        return VBScriptDrive(s)

    def _get_drives(self) -> list:
        if sys.platform == 'win32':
            import string
            return [f'{d}:\\' for d in string.ascii_uppercase if os.path.exists(f'{d}:\\')]
        return ['/']

    def __iter__(self):
        for d in self._get_drives():
            yield VBScriptDrive(d)


class VBScriptFileSystemObject(VBScriptObject):
    """Scripting.FileSystemObject implementation."""

    # -- File operations -----------------------------------------------------

    def FileExists(self, filespec: str) -> bool:
        return os.path.isfile(str(filespec))

    def FolderExists(self, folderspec: str) -> bool:
        return os.path.isdir(str(folderspec))

    def DriveExists(self, drivespec: str) -> bool:
        s = str(drivespec)
        if len(s) == 1:
            s = s + ':' + os.sep
        return os.path.exists(s)

    def GetFile(self, filespec: str) -> VBScriptFile:
        path = str(filespec)
        if not os.path.isfile(path):
            raise VBScriptError('File not found')
        return VBScriptFile(path)

    def GetFolder(self, folderspec: str) -> VBScriptFolder:
        path = str(folderspec)
        if not os.path.isdir(path):
            raise VBScriptError('Path not found')
        return VBScriptFolder(path)

    def GetDrive(self, drivespec: str) -> VBScriptDrive:
        s = str(drivespec)
        if len(s) == 1:
            s = s + ':' + os.sep
        return VBScriptDrive(s)

    @property
    def Drives(self) -> VBScriptDriveCollection:
        return VBScriptDriveCollection()

    def CreateTextFile(self, filename: str, overwrite: bool = True, unicode: bool = False) -> VBScriptTextStream:
        path = str(filename)
        if not overwrite and os.path.exists(path):
            raise VBScriptError('File already exists')
        fh = open(path, 'w', encoding='utf-8', newline='')
        return VBScriptTextStream(fh, _FOR_WRITING)

    def OpenTextFile(self, filename: str, iomode: int = _FOR_READING, create: bool = False, _format: int = _TRISTATE_FALSE) -> VBScriptTextStream:
        path = str(filename)
        mode = int(iomode)
        if mode == _FOR_READING:
            if not os.path.exists(path):
                if create:
                    open(path, 'w', encoding='utf-8', newline='').close()
                else:
                    raise VBScriptError('File not found')
            fh = open(path, 'r', encoding='utf-8', newline='')
        elif mode == _FOR_WRITING:
            fh = open(path, 'w', encoding='utf-8', newline='')
        elif mode == _FOR_APPENDING:
            fh = open(path, 'a', encoding='utf-8', newline='')
        else:
            raise VBScriptError('Bad file mode')
        return VBScriptTextStream(fh, mode)

    def DeleteFile(self, filespec: str, force: bool = False) -> None:
        path = str(filespec)
        if not os.path.isfile(path):
            raise VBScriptError('File not found')
        os.remove(path)

    def DeleteFolder(self, folderspec: str, force: bool = False) -> None:
        path = str(folderspec)
        if not os.path.isdir(path):
            raise VBScriptError('Path not found')
        shutil.rmtree(path)

    def CopyFile(self, source: str, destination: str, overwrite: bool = True) -> None:
        src = str(source)
        dest = str(destination)
        if os.path.isdir(dest):
            dest = os.path.join(dest, os.path.basename(src))
        if not overwrite and os.path.exists(dest):
            raise VBScriptError('File already exists')
        shutil.copy2(src, dest)

    def CopyFolder(self, source: str, destination: str, overwrite: bool = True) -> None:
        src = str(source)
        dest = str(destination)
        shutil.copytree(src, dest, dirs_exist_ok=overwrite)

    def MoveFile(self, source: str, destination: str) -> None:
        src = str(source)
        dest = str(destination)
        if os.path.isdir(dest):
            dest = os.path.join(dest, os.path.basename(src))
        shutil.move(src, dest)

    def MoveFolder(self, source: str, destination: str) -> None:
        shutil.move(str(source), str(destination))

    def CreateFolder(self, foldername: str) -> VBScriptFolder:
        path = str(foldername)
        if os.path.exists(path):
            raise VBScriptError('File already exists')
        os.makedirs(path)
        return VBScriptFolder(path)

    # -- Path helpers --------------------------------------------------------

    def BuildPath(self, path: str, name: str) -> str:
        return os.path.join(str(path), str(name))

    def GetFileName(self, pathspec: str) -> str:
        return os.path.basename(str(pathspec))

    def GetBaseName(self, pathspec: str) -> str:
        name = os.path.basename(str(pathspec))
        root, _ = os.path.splitext(name)
        return root

    def GetExtensionName(self, pathspec: str) -> str:
        _, ext = os.path.splitext(str(pathspec))
        return ext.lstrip('.')

    def GetParentFolderName(self, pathspec: str) -> str:
        return os.path.dirname(str(pathspec))

    def GetAbsolutePathName(self, pathspec: str) -> str:
        return os.path.abspath(str(pathspec))

    def GetTempName(self) -> str:
        return os.path.basename(tempfile.mktemp())

    def GetSpecialFolder(self, folderspec: int) -> VBScriptFolder:
        spec = int(folderspec)
        if spec == _TEMP_FOLDER:
            return VBScriptFolder(tempfile.gettempdir())
        elif spec == _WINDOWS_FOLDER:
            if sys.platform == 'win32':
                return VBScriptFolder(os.environ.get('WINDIR', 'C:\\Windows'))
            return VBScriptFolder('/tmp')
        elif spec == _SYSTEM_FOLDER:
            if sys.platform == 'win32':
                return VBScriptFolder(os.environ.get('SYSTEMROOT', 'C:\\Windows') + '\\System32')
            return VBScriptFolder('/usr')
        raise VBScriptError('Invalid procedure call or argument')


# ---------------------------------------------------------------------------
#  VBScriptClassDef / VBScriptClassInstance
# ---------------------------------------------------------------------------

@dataclass
class ClassPropertyDef:
    """Definition of a Property Get/Let/Set triplet."""

    get_params: Optional[List[Parameter]] = None
    get_body: Optional[List[ASTNode]] = None
    let_params: Optional[List[Parameter]] = None
    let_body: Optional[List[ASTNode]] = None
    set_params: Optional[List[Parameter]] = None
    set_body: Optional[List[ASTNode]] = None
    is_public: bool = True
    is_default: bool = False
    # Cached Procedure objects (built once during class registration)
    _get_proc: Optional['Procedure'] = None
    _let_proc: Optional['Procedure'] = None
    _set_proc: Optional['Procedure'] = None


@dataclass
class ClassMethodDef:
    """Definition of a Sub or Function in a class."""

    proc: 'Procedure'
    is_public: bool = True
    is_default: bool = False


@dataclass
class ClassFieldDef:
    """Definition of a field in a class."""

    name: str
    is_public: bool = True
    dimensions: Optional[List[ASTNode]] = None


class VBScriptClassDef:
    """Blueprint for a user-defined VBScript class."""

    def __init__(self, name: str):
        self.name = name
        self.fields: List[ClassFieldDef] = []
        self.field_names: Dict[str, str] = {}  # lowercase name -> original name
        self.methods: Dict[str, ClassMethodDef] = {}  # lowercase name -> def
        self.properties: Dict[str, ClassPropertyDef] = {}  # lowercase name -> def
        self.default_member: Optional[str] = None  # lowercase name of default member


class VBScriptClassInstance(VBScriptObject):
    """A live instance of a user-defined VBScript class."""

    def __init__(self, class_def: VBScriptClassDef, env: 'Environment'):
        self._class_def = class_def
        self._env = env  # instance-level environment (holds fields)

    @property
    def class_name(self) -> str:
        return self._class_def.name


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
        elif isinstance(value, VBScriptDate):
            return str(value)
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
