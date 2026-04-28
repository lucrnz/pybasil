"""Tests for the VBScript interpreter."""

import io
from pybasil import (
    run,
)


class TestDictionaryKeyCasePreservation:
    """Tests for preserving original key case in Dictionary.Keys() with CompareMode=1."""

    def test_keys_preserve_case_with_text_compare(self):
        from pybasil.interpreter import VBScriptDictionary

        d = VBScriptDictionary()
        d.CompareMode = 1
        d.Add("Hello", 1)
        d.Add("World", 2)
        keys_arr = d.Keys()
        assert keys_arr.get_element([0]) == "Hello"
        assert keys_arr.get_element([1]) == "World"

    def test_for_each_preserves_case_with_text_compare(self):
        from pybasil.interpreter import VBScriptDictionary

        d = VBScriptDictionary()
        d.CompareMode = 1
        d.Add("FooBar", 1)
        keys = list(d)
        assert keys == ["FooBar"]

    def test_keys_binary_mode_unchanged(self):
        output = io.StringIO()
        run(
            '''Set d = CreateObject("Scripting.Dictionary")
d.Add "Hello", 1
Dim k
Set k = d.Keys()
WScript.Echo k(0)
''',
            output_stream=output,
        )
        assert output.getvalue().strip() == 'Hello'

    def test_get_key_preserves_case(self):
        from pybasil.interpreter import VBScriptDictionary

        d = VBScriptDictionary()
        d.CompareMode = 1
        d.Add("MyKey", 42)
        assert d.get_key("mykey") == "MyKey"
