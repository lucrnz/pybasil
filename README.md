# pybasil

A VBScript parser and interpreter written in Python, aiming for full compatibility with VBScript 6.0 and beyond.

## Current Support

- **Tree-walking interpreter** targeting full VBScript 6.0 compatibility
- **Variables & literals**:
  - Variables are case-insensitive
  - Implicit variable creation is supported (`Empty` default)
  - `Dim` declarations (single and multiple variables)
  - Literals: numbers (including scientific notation), hex (`&HFF`), octal (`&O77`), strings, booleans, `Nothing`, `Empty`, `Null`
- **Statements**: `Dim`, assignments (`Let` optional), `Set`, `Call`, property assignments (`obj.Prop = value`), and expression statements (for things like `WScript.Echo`)
- **Operators**:
  - Arithmetic: `+`, `-`, `*`, `/`, `\` (integer division), `Mod`, `^`
  - String: `&`
  - Comparison: `=`, `<>`, `<`, `>`, `<=`, `>=`, `Is`
  - Logical: `And`, `Or`, `Not`, `Xor`, `Eqv`, `Imp`
- **Control flow**:
  - `If ... Then ... ElseIf ... Else ... End If`
  - `Select Case ... Case ... Case Else ... End Select` (including `Case x To y` ranges and `Case Is > x` comparisons)
  - `For ... To ... [Step ...] ... Next`
  - `For Each ... In ... Next`
  - `While ... Wend`
  - `Do While/Until ... Loop`, `Do ... Loop While/Until`, and `Do ... Loop`
  - `Exit For` and `Exit Do`
- **Arrays**:
  - Fixed-size arrays: `Dim arr(5)`
  - Dynamic arrays: `Dim arr()`, `ReDim arr(5)`, `ReDim Preserve arr(10)`
  - Multi-dimensional arrays: `Dim matrix(2, 2)`
  - `UBound`, `LBound`, `Array`, `Erase`
- **Error handling**:
  - `On Error Resume Next` - continue execution after errors
  - `On Error GoTo 0` - reset error handling to default
  - `Err` object with `Number`, `Source`, `Description`, `HelpFile`, `HelpContext` properties
  - `Err.Clear` - clear error information
  - `Err.Raise` - raise custom errors
- **Procedures**:
  - `Sub ... End Sub` and `Function ... End Function`
  - `Call` statements and implicit procedure calls
  - `ByRef` / `ByVal` parameters (`ByRef` default)
  - `Exit Sub` and `Exit Function`
  - Recursive function/procedure calls
- **Built-in runtime**:
  - `WScript.Echo`, `WScript.Quit`
  - String helpers (`Len`, `Left`, `Right`, `Mid`, `Trim`, `LTrim`, `RTrim`, `UCase`, `LCase`, `InStr`, `Replace`, `Split`, `Join`)
  - Conversion/type helpers (`CStr`, `CInt`, `CLng`, `CDbl`, `CBool`, `CDate`, `IsNumeric`, `IsArray`, `IsDate`, `IsEmpty`, `IsNull`, `IsObject`, `TypeName`, `VarType`)
  - Math/random helpers (`Abs`, `Sqr`, `Int`, `Fix`, `Round`, `Rnd`, `Randomize`)
  - `MsgBox`, `InputBox`, `CreateObject`, `GetObject` (simplified behavior)
- **Objects**:
  - `Scripting.Dictionary` via `CreateObject` (`Add`, `Exists`, `Items`, `Keys`, `Remove`, `RemoveAll`, `Item`, `Key`, `Count`, `CompareMode`)
- **Comments**: single quote (`'`) and `Rem`
- **CLI**: execute code from files, stdin, or `-c/--code`

## Installation

```bash
uv add pybasil
```

## Quick Start

### Command Line Usage

```bash
# Run a VBScript file
pybasil script.vbs

# Pipe VBScript code
echo 'WScript.Echo "Hello, World!"' | pybasil

# Execute code directly
pybasil -c 'WScript.Echo 2 + 2'
```

### Python API

```python
from pybasil import run, parse

# Execute VBScript code directly
run('WScript.Echo "Hello, World!"')

# Parse and interpret separately
program = parse("""
    Dim x, y
    x = 10
    y = 20
    WScript.Echo x + y
""")
from pybasil import Interpreter
interpreter = Interpreter()
interpreter.interpret(program)
```

## Usage Examples

### Variables and Assignments

```python
from pybasil import run

# Implicit variable creation
run('x = 42')
run('WScript.Echo x')  # Output: 42

# Dim statement
run("""
    Dim name, age
    name = "John"
    age = 30
    WScript.Echo name & " is " & age & " years old"
""")

# Let style (optional keyword)
run('Let x = 100')

# Set style for objects
run('Set obj = CreateObject("Scripting.FileSystemObject")')
```

### Arithmetic Operations

```python
from pybasil import run

run('WScript.Echo 5 + 3')      # Output: 8
run('WScript.Echo 10 - 4')     # Output: 6
run('WScript.Echo 6 * 7')      # Output: 42
run('WScript.Echo 15 / 3')     # Output: 5
run('WScript.Echo 17 \\ 5')    # Output: 3 (integer division)
run('WScript.Echo 17 Mod 5')   # Output: 2
run('WScript.Echo 2 ^ 10')     # Output: 1024
```

### String Operations

```python
from pybasil import run

# Concatenation with &
run('WScript.Echo "Hello" & " " & "World"')

# String functions
run('WScript.Echo Len("Hello")')           # Output: 5
run('WScript.Echo Left("Hello", 3)')       # Output: Hel
run('WScript.Echo Right("Hello", 3)')      # Output: llo
run('WScript.Echo UCase("hello")')         # Output: HELLO
run('WScript.Echo LCase("HELLO")')         # Output: hello
```

### Comparison and Logical Operators

```python
from pybasil import run

# Comparisons
run('WScript.Echo 5 = 5')      # Output: True
run('WScript.Echo 5 <> 6')     # Output: True
run('WScript.Echo 3 < 5')      # Output: True

# Logical operators
run('WScript.Echo True And False')   # Output: False
run('WScript.Echo True Or False')    # Output: True
run('WScript.Echo Not True')         # Output: False
```

### Special Values

```python
from pybasil import run

# Nothing - for object references
run('x = Nothing')

# Empty - uninitialized variable
run('x = Empty')

# Null - database null value
run('x = Null')
```

### Error Handling

```python
from pybasil import run

# On Error Resume Next - continue after errors
run("""
    On Error Resume Next
    x = 1 / 0
    WScript.Echo "After error"
    WScript.Echo Err.Number
""")

# Err.Raise - raise custom errors
run("""
    On Error Resume Next
    Err.Raise 100, "MyApp", "Custom error message"
    WScript.Echo Err.Number
    WScript.Echo Err.Description
""")

# On Error GoTo 0 - reset error handling
run("""
    On Error Resume Next
    x = 1 / 0
    On Error GoTo 0
    y = 1 / 0  ' This will raise an error
""")
```

## API Reference

### `run(source: str, output_stream=None)`

Parse and execute VBScript source code.

```python
from pybasil import run

run('WScript.Echo "Hello"')
```

### `parse(source: str) -> Program`

Parse VBScript source code and return an AST.

```python
from pybasil import parse

program = parse('x = 42')
```

### `Interpreter(output_stream=None)`

Tree-walking interpreter for VBScript AST.

```python
from pybasil import Interpreter, parse

interpreter = Interpreter()
program = parse('x = 42')
interpreter.interpret(program)
```

### `VBScriptParser`

Parser class that preprocesses source (including `Rem` comments) and returns an AST.

```python
from pybasil import VBScriptParser

parser = VBScriptParser()
program = parser.parse('x = 42')
```

## Development

### Setup

```bash
git clone https://github.com/lucrnz/pybasil.git
cd pybasil
uv sync
```

### Running Tests

```bash
uv run pytest
```

## License

MIT License
