# pyvbscript

A VBScript parser and interpreter written in Python. This library enables companies running legacy VBScript projects to migrate to cloud-native scenarios using Python instead of Windows.

## Features

- **Tree-walking interpreter** for VBScript source code
- **Variables**: Implicit creation allowed, Dim statement (single & multiple variables)
- **Literals**: Numbers, Strings, Booleans, Nothing, Empty, Null
- **Assignments**: Let style (`x = 5`) and Set style (`Set obj = CreateObject()`)
- **Operators**:
  - Arithmetic: `+`, `-`, `*`, `/`, `\` (integer division), `Mod`, `^` (exponentiation)
  - String: `&` (concatenation)
  - Comparison: `=`, `<>`, `<`, `>`, `<=`, `>=`, `Is`
  - Logical: `And`, `Or`, `Not`, `Xor`, `Eqv`, `Imp`
- **WScript.Echo**: Basic output for testing (prints to stdout)
- **Built-in functions**: Len, Left, Right, Mid, Trim, UCase, LCase, CStr, CInt, and more
- **CLI**: Run VBScript code from files, stdin, or command-line arguments

## Installation

```bash
uv add pyvbscript
```

Or with pip:

```bash
pip install pyvbscript
```

## Quick Start

### Command Line Usage

```bash
# Run a VBScript file
pyvbscript script.vbs

# Pipe VBScript code
echo 'WScript.Echo "Hello, World!"' | pyvbscript

# Execute code directly
pyvbscript -c 'WScript.Echo 2 + 2'
```

### Python API

```python
from pyvbscript import run, parse

# Execute VBScript code directly
run('WScript.Echo "Hello, World!"')

# Parse and interpret separately
program = parse("""
    Dim x, y
    x = 10
    y = 20
    WScript.Echo x + y
""")
from pyvbscript import Interpreter
interpreter = Interpreter()
interpreter.interpret(program)
```

## Usage Examples

### Variables and Assignments

```python
from pyvbscript import run

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
from pyvbscript import run

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
from pyvbscript import run

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
from pyvbscript import run

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
from pyvbscript import run

# Nothing - for object references
run('x = Nothing')

# Empty - uninitialized variable
run('x = Empty')

# Null - database null value
run('x = Null')
```

## API Reference

### `run(source: str, output_stream=None)`

Parse and execute VBScript source code.

```python
from pyvbscript import run

run('WScript.Echo "Hello"')
```

### `parse(source: str) -> Program`

Parse VBScript source code and return an AST.

```python
from pyvbscript import parse

program = parse('x = 42')
```

### `Interpreter`

Tree-walking interpreter for VBScript AST.

```python
from pyvbscript import Interpreter, parse

interpreter = Interpreter()
program = parse('x = 42')
interpreter.interpret(program)
```

## Development

### Setup

```bash
git clone https://github.com/your-org/pyvbscript.git
cd pyvbscript
uv sync
```

### Running Tests

```bash
uv run pytest
```

## License

MIT License
