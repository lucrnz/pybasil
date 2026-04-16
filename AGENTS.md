# AGENTS.md
This repository is a Python implementation of a VBScript parser and tree-walking interpreter, that aims to be fully compatible with VBScript 6.0

## Commands
Install dependencies: `uv sync`
Run all tests: `uv run pytest`
Run parser tests: `uv run pytest tests/test_parser.py -q`
Run interpreter tests: `uv run pytest tests/test_interpreter.py -q`
Run targeted tests: `uv run pytest tests/test_interpreter.py -k "<pattern>" -q`
Run CLI locally: `uv run pybasil -c 'WScript.Echo 2 + 2'`
Run linter: `uv run ruff check`
Run linter (auto-fix): `uv run ruff check --fix`

## Always
Keep behavior aligned with existing VBScript semantics covered by tests
Add or update tests for any behavior change.
Keep changes scoped and consistent with current style (typed Python, dataclasses for AST, clear method naming).
Update `docs/language_support_status.md` when language support changes, and `README.md` when public API changes.
Follow conventional commits for git
