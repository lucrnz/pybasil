"""Tests for the VBScript interpreter."""

import pytest
import io
from pybasil import (
    Interpreter,
    parse,
    run,
    VBScriptArray,
    VBScriptError,
    EMPTY,
)


class TestInterpreterIfStatement:
    """Test If/Then/Else/ElseIf statements."""

    def test_if_then_true(self):
        output = io.StringIO()
        run(
            """
            If True Then
                WScript.Echo "yes"
            End If
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'yes'

    def test_if_then_false(self):
        output = io.StringIO()
        run(
            """
            If False Then
                WScript.Echo "yes"
            End If
            WScript.Echo "done"
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'done'

    def test_if_then_else_true(self):
        output = io.StringIO()
        run(
            """
            If True Then
                WScript.Echo "then"
            Else
                WScript.Echo "else"
            End If
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'then'

    def test_if_then_else_false(self):
        output = io.StringIO()
        run(
            """
            If False Then
                WScript.Echo "then"
            Else
                WScript.Echo "else"
            End If
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'else'

    def test_if_elseif_else_first(self):
        output = io.StringIO()
        run(
            """
            x = 1
            If x = 1 Then
                WScript.Echo "one"
            ElseIf x = 2 Then
                WScript.Echo "two"
            Else
                WScript.Echo "other"
            End If
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'one'

    def test_if_elseif_else_second(self):
        output = io.StringIO()
        run(
            """
            x = 2
            If x = 1 Then
                WScript.Echo "one"
            ElseIf x = 2 Then
                WScript.Echo "two"
            Else
                WScript.Echo "other"
            End If
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'two'

    def test_if_elseif_else_fallback(self):
        output = io.StringIO()
        run(
            """
            x = 3
            If x = 1 Then
                WScript.Echo "one"
            ElseIf x = 2 Then
                WScript.Echo "two"
            Else
                WScript.Echo "other"
            End If
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'other'

    def test_if_multiple_elseif(self):
        output = io.StringIO()
        run(
            """
            x = 3
            If x = 1 Then
                WScript.Echo "one"
            ElseIf x = 2 Then
                WScript.Echo "two"
            ElseIf x = 3 Then
                WScript.Echo "three"
            ElseIf x = 4 Then
                WScript.Echo "four"
            Else
                WScript.Echo "other"
            End If
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'three'

    def test_if_nested(self):
        output = io.StringIO()
        run(
            """
            x = 5
            If x > 0 Then
                If x > 10 Then
                    WScript.Echo "big"
                Else
                    WScript.Echo "small"
                End If
            Else
                WScript.Echo "negative"
            End If
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'small'

    def test_if_case_insensitive(self):
        output = io.StringIO()
        run(
            """
            IF TRUE THEN
                wscript.echo "yes"
            END IF
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'yes'

class TestInterpreterForStatement:
    """Test For...Next statements."""

    def test_for_basic(self):
        output = io.StringIO()
        run(
            """
            For i = 1 To 3
                WScript.Echo i
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['1', '2', '3']

    def test_for_with_step(self):
        output = io.StringIO()
        run(
            """
            For i = 0 To 10 Step 2
                WScript.Echo i
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0', '2', '4', '6', '8', '10']

    def test_for_negative_step(self):
        output = io.StringIO()
        run(
            """
            For i = 5 To 1 Step -1
                WScript.Echo i
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['5', '4', '3', '2', '1']

    def test_for_exit_for(self):
        output = io.StringIO()
        run(
            """
            For i = 1 To 10
                If i = 3 Then
                    Exit For
                End If
                WScript.Echo i
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['1', '2']

    def test_for_nested(self):
        output = io.StringIO()
        run(
            """
            For i = 1 To 2
                For j = 1 To 2
                    WScript.Echo i & "," & j
                Next
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['1,1', '1,2', '2,1', '2,2']

    def test_for_variable_after_loop(self):
        program = parse("""
            For i = 1 To 5
            Next
        """)
        interpreter = Interpreter()
        interpreter.interpret(program)
        # After loop, i should be 6 (last value + step)
        assert interpreter._environment.get('i') == 6

    def test_for_empty_body(self):
        program = parse("""
            For i = 1 To 5
            Next
        """)
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.exists('i')

    def test_for_default_step_countdown(self):
        output = io.StringIO()
        run(
            """
            For i = 5 To 1
                WScript.Echo i
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['5', '4', '3', '2', '1']

    def test_for_default_step_countdown_single(self):
        output = io.StringIO()
        run(
            """
            For i = 3 To 3
                WScript.Echo i
            Next
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '3'

class TestInterpreterForEachDictionary:
    """Test For Each over Scripting.Dictionary."""

    def test_for_each_dictionary_yields_keys(self):
        output = io.StringIO()
        run(
            """
            Set d = CreateObject("Scripting.Dictionary")
            d.Add "a", 100
            d.Add "b", 200
            For Each k In d
                WScript.Echo k
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['a', 'b']

    def test_for_each_dictionary_access_values_via_item(self):
        output = io.StringIO()
        run(
            """
            Set d = CreateObject("Scripting.Dictionary")
            d.Add "x", 42
            d.Add "y", 99
            For Each k In d
                WScript.Echo d.Item(k)
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['42', '99']

    def test_empty_dictionary_items_returns_empty_array(self):
        from pybasil.interpreter import VBScriptDictionary
        d = VBScriptDictionary()
        arr = d.Items()
        assert isinstance(arr, VBScriptArray)
        assert arr.ubound() == -1

    def test_empty_dictionary_keys_returns_empty_array(self):
        from pybasil.interpreter import VBScriptDictionary
        d = VBScriptDictionary()
        arr = d.Keys()
        assert isinstance(arr, VBScriptArray)
        assert arr.ubound() == -1

class TestInterpreterWhileStatement:
    """Test While...Wend statements."""

    def test_while_basic(self):
        output = io.StringIO()
        run(
            """
            x = 0
            While x < 3
                WScript.Echo x
                x = x + 1
            Wend
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0', '1', '2']

    def test_while_false_initially(self):
        output = io.StringIO()
        run(
            """
            While False
                WScript.Echo "never"
            Wend
            WScript.Echo "done"
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'done'

    def test_while_nested(self):
        output = io.StringIO()
        run(
            """
            i = 0
            While i < 2
                j = 0
                While j < 2
                    WScript.Echo i & "," & j
                    j = j + 1
                Wend
                i = i + 1
            Wend
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0,0', '0,1', '1,0', '1,1']

    def test_exit_for_in_while_raises_error(self):
        program = parse("""
            Dim count
            count = 0
            While count < 5
                count = count + 1
                Exit For
            Wend
        """)
        interpreter = Interpreter()
        with pytest.raises(VBScriptError, match='Exit For not valid in While loop'):
            interpreter.interpret(program)

class TestInterpreterDoLoop:
    """Test Do...Loop statements."""

    def test_do_while_pre_test(self):
        output = io.StringIO()
        run(
            """
            x = 0
            Do While x < 3
                WScript.Echo x
                x = x + 1
            Loop
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0', '1', '2']

    def test_do_until_pre_test(self):
        output = io.StringIO()
        run(
            """
            x = 0
            Do Until x >= 3
                WScript.Echo x
                x = x + 1
            Loop
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0', '1', '2']

    def test_do_loop_while_post_test(self):
        output = io.StringIO()
        run(
            """
            x = 0
            Do
                WScript.Echo x
                x = x + 1
            Loop While x < 3
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0', '1', '2']

    def test_do_loop_until_post_test(self):
        output = io.StringIO()
        run(
            """
            x = 0
            Do
                WScript.Echo x
                x = x + 1
            Loop Until x >= 3
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0', '1', '2']

    def test_do_while_post_executes_once(self):
        output = io.StringIO()
        run(
            """
            x = 10
            Do
                WScript.Echo "once"
            Loop While x < 5
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'once'

    def test_do_until_post_executes_once(self):
        output = io.StringIO()
        run(
            """
            x = 10
            Do
                WScript.Echo "once"
            Loop Until x > 5
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'once'

    def test_do_exit_do(self):
        output = io.StringIO()
        run(
            """
            x = 0
            Do While True
                WScript.Echo x
                x = x + 1
                If x >= 3 Then
                    Exit Do
                End If
            Loop
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0', '1', '2']

    def test_do_nested_exit(self):
        output = io.StringIO()
        run(
            """
            x = 0
            Do
                y = 0
                Do
                    WScript.Echo x & "," & y
                    y = y + 1
                    If y >= 2 Then
                        Exit Do
                    End If
                Loop
                x = x + 1
                If x >= 2 Then
                    Exit Do
                End If
            Loop
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0,0', '0,1', '1,0', '1,1']

class TestInterpreterExitStatement:
    """Test Exit statements."""

    def test_exit_for_basic(self):
        output = io.StringIO()
        run(
            """
            For i = 1 To 10
                WScript.Echo i
                Exit For
            Next
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '1'

    def test_exit_do_basic(self):
        output = io.StringIO()
        run(
            """
            Do While True
                WScript.Echo "once"
                Exit Do
            Loop
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'once'

    def test_exit_for_nested(self):
        output = io.StringIO()
        run(
            """
            For i = 1 To 3
                For j = 1 To 3
                    If j = 2 Then
                        Exit For
                    End If
                    WScript.Echo i & "," & j
                Next
            Next
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['1,1', '2,1', '3,1']

class TestSelectCase:
    """Test Select Case statement."""

    def test_single_value_match(self):
        output = io.StringIO()
        run(
            """
            Dim x
            x = 2
            Select Case x
                Case 1
                    WScript.Echo "one"
                Case 2
                    WScript.Echo "two"
                Case 3
                    WScript.Echo "three"
            End Select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'two'

    def test_no_match_no_else(self):
        output = io.StringIO()
        run(
            """
            Dim x
            x = 99
            Select Case x
                Case 1
                    WScript.Echo "one"
                Case 2
                    WScript.Echo "two"
            End Select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == ''

    def test_case_else(self):
        output = io.StringIO()
        run(
            """
            Dim x
            x = 42
            Select Case x
                Case 1
                    WScript.Echo "one"
                Case Else
                    WScript.Echo "other"
            End Select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'other'

    def test_comma_separated_values(self):
        output = io.StringIO()
        run(
            """
            Dim x
            x = 5
            Select Case x
                Case 1, 2, 3
                    WScript.Echo "small"
                Case 4, 5, 6
                    WScript.Echo "medium"
                Case Else
                    WScript.Echo "large"
            End Select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'medium'

    def test_relational_checks_with_true(self):
        output = io.StringIO()
        run(
            """
            Dim score
            score = 85
            Select Case True
                Case score >= 90
                    WScript.Echo "A"
                Case score >= 80
                    WScript.Echo "B"
                Case score >= 70
                    WScript.Echo "C"
                Case Else
                    WScript.Echo "F"
            End Select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'B'

    def test_string_matching(self):
        output = io.StringIO()
        run(
            """
            Dim fruit
            fruit = "Banana"
            Select Case fruit
                Case "Apple", "Pear"
                    WScript.Echo "pome"
                Case "Banana", "Mango"
                    WScript.Echo "tropical"
                Case Else
                    WScript.Echo "unknown"
            End Select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'tropical'

    def test_select_case_in_sub(self):
        output = io.StringIO()
        run(
            """
            Sub Classify(n)
                Select Case True
                    Case n < 0
                        WScript.Echo "negative"
                    Case n = 0
                        WScript.Echo "zero"
                    Case n > 0
                        WScript.Echo "positive"
                End Select
            End Sub

            Call Classify(-5)
            Call Classify(0)
            Call Classify(42)
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['negative', 'zero', 'positive']

    def test_select_case_first_match_wins(self):
        output = io.StringIO()
        run(
            """
            Dim x
            x = 5
            Select Case True
                Case x > 0
                    WScript.Echo "positive"
                Case x > 3
                    WScript.Echo "greater than 3"
                Case x > 1
                    WScript.Echo "greater than 1"
            End Select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'positive'

    def test_select_case_with_assignment_in_body(self):
        output = io.StringIO()
        run(
            """
            Dim x, result
            x = 2
            Select Case x
                Case 1
                    result = "a"
                Case 2
                    result = "b"
                Case 3
                    result = "c"
            End Select
            WScript.Echo result
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'b'

    def test_select_case_multiple_statements_in_body(self):
        output = io.StringIO()
        run(
            """
            Dim x
            x = 1
            Select Case x
                Case 1
                    WScript.Echo "line1"
                    WScript.Echo "line2"
                Case 2
                    WScript.Echo "other"
            End Select
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['line1', 'line2']

    def test_select_case_string_comma_list_with_grade(self):
        output = io.StringIO()
        run(
            """
            Dim score, grade
            score = 85
            Select Case score
                Case 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100
                    grade = "A"
                Case 80, 81, 82, 83, 84, 85, 86, 87, 88, 89
                    grade = "B"
                Case 70, 71, 72, 73, 74, 75, 76, 77, 78, 79
                    grade = "C"
                Case Else
                    grade = "F"
            End Select
            WScript.Echo "Score " & score & " = Grade " & grade
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'Score 85 = Grade B'

    def test_select_case_relational_with_else(self):
        output = io.StringIO()
        run(
            """
            Dim score, grade
            score = 95
            Select Case True
                Case score = 100
                    grade = "Perfect"
                Case score >= 90
                    grade = "Excellent"
                Case score >= 80
                    grade = "Good"
                Case Else
                    grade = "Keep trying"
            End Select
            WScript.Echo "Score " & score & " = " & grade
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'Score 95 = Excellent'

    def test_case_insensitive_select_case(self):
        output = io.StringIO()
        run(
            """
            Dim x
            x = 1
            select case x
                case 1
                    wscript.echo "matched"
            end select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'matched'

    def test_nested_select_case(self):
        output = io.StringIO()
        run(
            """
            Dim x, y
            x = 1
            y = 2
            Select Case x
                Case 1
                    Select Case y
                        Case 1
                            WScript.Echo "1-1"
                        Case 2
                            WScript.Echo "1-2"
                    End Select
                Case 2
                    WScript.Echo "2-x"
            End Select
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '1-2'

class TestSelectCaseRangeAndIs:
    """Tests for Case x To y and Case Is <op> expr in Select Case."""

    def test_case_range_match(self):
        output = io.StringIO()
        run(
            '''Dim x
x = 5
Select Case x
    Case 1 To 3
        WScript.Echo "low"
    Case 4 To 7
        WScript.Echo "mid"
    Case Else
        WScript.Echo "other"
End Select
''',
            output_stream=output,
        )
        assert output.getvalue().strip() == 'mid'

    def test_case_range_no_match(self):
        output = io.StringIO()
        run(
            '''Dim x
x = 10
Select Case x
    Case 1 To 3
        WScript.Echo "low"
    Case 4 To 7
        WScript.Echo "mid"
    Case Else
        WScript.Echo "other"
End Select
''',
            output_stream=output,
        )
        assert output.getvalue().strip() == 'other'

    def test_case_is_greater_than(self):
        output = io.StringIO()
        run(
            '''Dim x
x = 10
Select Case x
    Case Is > 5
        WScript.Echo "big"
    Case Else
        WScript.Echo "small"
End Select
''',
            output_stream=output,
        )
        assert output.getvalue().strip() == 'big'

    def test_case_is_less_than(self):
        output = io.StringIO()
        run(
            '''Dim x
x = 2
Select Case x
    Case Is < 5
        WScript.Echo "small"
    Case Else
        WScript.Echo "big"
End Select
''',
            output_stream=output,
        )
        assert output.getvalue().strip() == 'small'

    def test_case_mixed_range_is_value(self):
        output = io.StringIO()
        run(
            '''Dim x
x = 15
Select Case x
    Case 1 To 5
        WScript.Echo "1-5"
    Case 6, 7, 8
        WScript.Echo "6-8"
    Case Is >= 10
        WScript.Echo ">=10"
    Case Else
        WScript.Echo "other"
End Select
''',
            output_stream=output,
        )
        assert output.getvalue().strip() == '>=10'

class TestSingleLineIf:
    """Test single-line If...Then [Else] statements."""

    def test_inline_if_true(self):
        program = parse('If True Then x = 1')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 1

    def test_inline_if_false(self):
        program = parse('If False Then x = 1')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == EMPTY

    def test_inline_if_else_true(self):
        program = parse('If True Then x = 1 Else x = 2')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 1

    def test_inline_if_else_false(self):
        program = parse('If False Then x = 1 Else x = 2')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 2

    def test_inline_if_expression_condition(self):
        program = parse('''
        Dim n
        n = 10
        If n > 5 Then result = "big" Else result = "small"
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'big'

    def test_inline_if_expression_condition_false(self):
        program = parse('''
        Dim n
        n = 3
        If n > 5 Then result = "big" Else result = "small"
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'small'

    def test_inline_if_multiple_statements_colon(self):
        program = parse('If True Then x = 1 : y = 2')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 1
        assert interp._environment.get('y') == 2

    def test_inline_if_else_multiple_statements(self):
        program = parse('If False Then x = 1 : y = 2 Else x = 3 : y = 4')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 3
        assert interp._environment.get('y') == 4

    def test_inline_if_with_string_containing_then(self):
        """Ensure 'Then' inside a string doesn't confuse the parser."""
        program = parse('If True Then x = "Then what"')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'Then what'

    def test_inline_if_with_string_containing_else(self):
        """Ensure 'Else' inside a string doesn't confuse the parser."""
        program = parse('If True Then x = "Else nothing"')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'Else nothing'

    def test_inline_if_method_call(self):
        output = io.StringIO()
        program = parse('If True Then WScript.Echo "yes"')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        assert output.getvalue().strip() == 'yes'

    def test_inline_if_method_call_else(self):
        output = io.StringIO()
        program = parse('If False Then WScript.Echo "no" Else WScript.Echo "yes"')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        assert output.getvalue().strip() == 'yes'

    def test_inline_if_does_not_break_block_if(self):
        """Block If...Then...End If should still work."""
        program = parse('''
        If True Then
            x = 42
        End If
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 42

    def test_inline_if_does_not_break_block_if_else(self):
        """Block If...Else...End If should still work."""
        program = parse('''
        If False Then
            x = 1
        Else
            x = 2
        End If
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 2

    def test_inline_if_does_not_break_block_if_elseif(self):
        """Block If...ElseIf...End If should still work."""
        program = parse('''
        Dim n
        n = 5
        If n > 10 Then
            x = "big"
        ElseIf n > 3 Then
            x = "medium"
        Else
            x = "small"
        End If
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 'medium'

    def test_inline_if_preserves_colon_block_if(self):
        """Colon-separated block If (with End If) should not be rewritten."""
        program = parse('If True Then : x = 42 : End If')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 42

    def test_inline_if_in_loop(self):
        output = io.StringIO()
        program = parse('''
        Dim i, total
        total = 0
        For i = 1 To 10
            If i Mod 2 = 0 Then total = total + i
        Next
        WScript.Echo total
        ''')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        assert output.getvalue().strip() == '30'

    def test_inline_if_with_function_call(self):
        program = parse('''
        Function Double(n)
            Double = n * 2
        End Function
        If True Then x = Double(5)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 10

    def test_inline_if_with_comment(self):
        program = parse("If True Then x = 42 ' set x")
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 42

    def test_inline_if_case_insensitive(self):
        program = parse('IF TRUE THEN x = 1 ELSE x = 2')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('x') == 1

    def test_inline_if_nested_in_sub(self):
        program = parse('''
        Dim result
        Sub Check(val)
            If val > 0 Then result = "positive" Else result = "non-positive"
        End Sub
        Call Check(5)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'positive'
