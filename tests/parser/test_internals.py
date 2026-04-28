"""Tests for the VBScript parser."""

from pybasil import (
    parse,
)


class TestPrecompiledRemRegex:
    """Test that the REM-comment regex is pre-compiled at module level."""

    def test_rem_regex_compiled(self):
        from pybasil.parser import _REM_RE
        import re
        assert isinstance(_REM_RE, re.Pattern)

    def test_rem_comment_still_stripped(self):
        result = parse('x = 1 Rem this is a comment\n')
        assert len(result.statements) == 1

class TestCachedParser:
    """Test that the module-level parse() function caches its parser."""

    def test_parse_returns_same_parser(self):
        from pybasil.parser import parse as parse_fn
        parse_fn('Dim x\n')
        from pybasil import parser as parser_mod
        p1 = parser_mod._cached_parser
        parse_fn('Dim y\n')
        p2 = parser_mod._cached_parser
        assert p1 is p2

    def test_cached_parse_produces_valid_ast(self):
        result = parse('x = 1\ny = 2\n')
        assert len(result.statements) == 2
