"""Tests for the VBScript interpreter."""

import pytest
from pybasil import (
    Interpreter,
    parse,
    VBScriptError,
)


class TestInterpreterStringOperations:
    """Test string operations."""

    def test_concatenation(self):
        program = parse('x = "Hello" & " " & "World"')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'Hello World'

    def test_string_number_concatenation(self):
        program = parse('x = "Value: " & 42')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'Value: 42'

    def test_number_string_addition(self):
        # In VBScript, + with a numeric string and number does arithmetic addition
        # (the string is converted to a number)
        program = parse('x = "5" + 3')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 8  # Numeric addition

    def test_escaped_double_quotes(self):
        program = parse('x = "He said ""hello"""')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == 'He said "hello"'

    def test_empty_escaped_quotes(self):
        program = parse('x = """"""')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('x') == '""'

class TestStringFunctionsChr:
    """Test Chr and ChrW functions."""

    def test_chr_letter(self):
        program = parse('x = Chr(65)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'A'

    def test_chr_zero(self):
        program = parse('x = Chr(0)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == '\x00'

    def test_chr_newline(self):
        program = parse('x = Chr(10)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == '\n'

    def test_chr_out_of_range(self):
        program = parse('x = Chr(256)')
        interp = Interpreter()
        with pytest.raises(VBScriptError):
            interp.interpret(program)

    def test_chr_negative(self):
        program = parse('x = Chr(-1)')
        interp = Interpreter()
        with pytest.raises(VBScriptError):
            interp.interpret(program)

    def test_chrw_unicode(self):
        program = parse('x = ChrW(8364)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == '\u20ac'  # Euro sign

    def test_chrw_basic(self):
        program = parse('x = ChrW(65)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'A'

class TestStringFunctionsAsc:
    """Test Asc and AscW functions."""

    def test_asc_letter(self):
        program = parse('x = Asc("A")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 65

    def test_asc_lowercase(self):
        program = parse('x = Asc("a")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 97

    def test_asc_first_char(self):
        program = parse('x = Asc("Hello")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 72  # 'H'

    def test_asc_empty_string_error(self):
        program = parse('x = Asc("")')
        interp = Interpreter()
        with pytest.raises(VBScriptError):
            interp.interpret(program)

    def test_ascw_unicode(self):
        program = parse('x = AscW(ChrW(8364))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 8364

    def test_asc_space(self):
        program = parse('x = Asc(" ")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 32

class TestStringFunctionsInStrRev:
    """Test InStrRev function."""

    def test_instrrev_basic(self):
        program = parse('x = InStrRev("Hello World", "o")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 8  # Last 'o' in "World"

    def test_instrrev_not_found(self):
        program = parse('x = InStrRev("Hello", "z")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 0

    def test_instrrev_with_start(self):
        program = parse('x = InStrRev("Hello World", "o", 6)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 5  # 'o' in "Hello"

    def test_instrrev_case_insensitive(self):
        program = parse('x = InStrRev("Hello World", "WORLD", -1, 1)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 7

    def test_instrrev_substring(self):
        program = parse('x = InStrRev("abcabc", "abc")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 4

class TestStringFunctionsStrComp:
    """Test StrComp function."""

    def test_strcomp_equal(self):
        program = parse('x = StrComp("abc", "abc")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 0

    def test_strcomp_less(self):
        program = parse('x = StrComp("abc", "def")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == -1

    def test_strcomp_greater(self):
        program = parse('x = StrComp("def", "abc")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 1

    def test_strcomp_case_sensitive(self):
        program = parse('x = StrComp("ABC", "abc", 0)')
        interp = Interpreter()
        interp.interpret(program)
        # Binary compare: 'A' (65) < 'a' (97)
        assert interp._environment.get('x') == -1

    def test_strcomp_case_insensitive(self):
        program = parse('x = StrComp("ABC", "abc", 1)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 0

class TestStringFunctionsStringSpace:
    """Test String and Space functions."""

    def test_string_char(self):
        program = parse('x = String(5, "x")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'xxxxx'

    def test_string_charcode(self):
        program = parse('x = String(3, 65)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'AAA'

    def test_string_zero(self):
        program = parse('x = String(0, "x")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == ''

    def test_string_negative_error(self):
        program = parse('x = String(-1, "x")')
        interp = Interpreter()
        with pytest.raises(VBScriptError):
            interp.interpret(program)

    def test_string_first_char_only(self):
        program = parse('x = String(3, "Hello")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'HHH'

    def test_space_basic(self):
        program = parse('x = Space(5)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == '     '

    def test_space_zero(self):
        program = parse('x = Space(0)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == ''

    def test_space_negative_error(self):
        program = parse('x = Space(-1)')
        interp = Interpreter()
        with pytest.raises(VBScriptError):
            interp.interpret(program)

class TestStringFunctionsStrReverse:
    """Test StrReverse function."""

    def test_strreverse_basic(self):
        program = parse('x = StrReverse("Hello")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'olleH'

    def test_strreverse_empty(self):
        program = parse('x = StrReverse("")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == ''

    def test_strreverse_single(self):
        program = parse('x = StrReverse("A")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'A'

    def test_strreverse_palindrome(self):
        program = parse('x = StrReverse("racecar")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'racecar'

class TestStringFunctionsHexOct:
    """Test Hex and Oct functions."""

    def test_hex_positive(self):
        program = parse('x = Hex(255)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'FF'

    def test_hex_zero(self):
        program = parse('x = Hex(0)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == '0'

    def test_hex_small(self):
        program = parse('x = Hex(10)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'A'

    def test_hex_large(self):
        program = parse('x = Hex(65535)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'FFFF'

    def test_oct_basic(self):
        program = parse('x = Oct(8)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == '10'

    def test_oct_zero(self):
        program = parse('x = Oct(0)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == '0'

    def test_oct_255(self):
        program = parse('x = Oct(255)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == '377'

    def test_chr_and_asc_roundtrip(self):
        """Verify Chr(Asc(x)) == x for a character."""
        program = parse('x = Chr(Asc("Z"))')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'Z'

    def test_hex_negative(self):
        program = parse('x = Hex(-1)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'FFFF'

    def test_oct_negative(self):
        program = parse('x = Oct(-1)')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == '177777'
