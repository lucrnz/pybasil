# pybasil

A VBScript parser and interpreter written in Python, aiming for full compatibility with VBScript 6.0 and beyond.

## Current Support

See [Language Support Status](docs/language_support_status.md) for a detailed list of supported features.

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

[MIT License](./LICENSE)
