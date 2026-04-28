"""Tests for the VBScript interpreter."""

import io
from pybasil import (
    run,
)


class TestInterpreterSub:
    """Test Sub procedures."""

    def test_sub_no_params(self):
        output = io.StringIO()
        run(
            """
            Sub SayHello
                WScript.Echo "Hello"
            End Sub
            
            SayHello
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'Hello'

    def test_sub_with_params(self):
        output = io.StringIO()
        run(
            """
            Sub Greet(name)
                WScript.Echo "Hello, " & name
            End Sub
            
            Greet "World"
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'Hello, World'

    def test_sub_call_with_call_keyword(self):
        output = io.StringIO()
        run(
            """
            Sub SayHello
                WScript.Echo "Hello"
            End Sub
            
            Call SayHello
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == 'Hello'

    def test_sub_exit_sub(self):
        output = io.StringIO()
        run(
            """
            Sub TestExit
                WScript.Echo "Before"
                Exit Sub
                WScript.Echo "After"
            End Sub
            
            TestExit
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['Before']

    def test_sub_local_scope(self):
        output = io.StringIO()
        # Note: Using Call keyword for procedure call without arguments
        # to avoid ambiguity with member access chains
        run(
            """
            x = 10
            
            Sub TestScope
                Dim x
                x = 20
                WScript.Echo "Inside: " & x
            End Sub
            
            Call TestScope
            WScript.Echo "Outside: " & x
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['Inside: 20', 'Outside: 10']

    def test_sub_called_in_expression_returns_empty(self):
        output = io.StringIO()
        run(
            """
            Sub MySub(x)
                WScript.Echo x
            End Sub

            result = MySub(42)
            WScript.Echo TypeName(result)
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['42', 'Empty']

class TestInterpreterFunction:
    """Test Function procedures."""

    def test_function_no_params(self):
        output = io.StringIO()
        run(
            """
            Function GetAnswer
                GetAnswer = 42
            End Function
            
            result = GetAnswer
            WScript.Echo result
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '42'

    def test_function_with_params(self):
        output = io.StringIO()
        run(
            """
            Function Add(a, b)
                Add = a + b
            End Function
            
            result = Add(3, 4)
            WScript.Echo result
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '7'

    def test_function_in_expression(self):
        output = io.StringIO()
        run(
            """
            Function DoubleVal(x)
                DoubleVal = x * 2
            End Function
            
            result = DoubleVal(5) + 1
            WScript.Echo result
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '11'

    def test_function_exit_function(self):
        output = io.StringIO()
        run(
            """
            Function EarlyReturn
                EarlyReturn = 1
                Exit Function
                EarlyReturn = 2
            End Function
            
            WScript.Echo EarlyReturn
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '1'

    def test_function_nested_call(self):
        output = io.StringIO()
        run(
            """
            Function Square(x)
                Square = x * x
            End Function
            
            Function SumOfSquares(a, b)
                SumOfSquares = Square(a) + Square(b)
            End Function
            
            WScript.Echo SumOfSquares(3, 4)
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '25'

class TestInterpreterByRefByVal:
    """Test ByRef and ByVal parameter passing."""

    def test_byref_modifies_original(self):
        output = io.StringIO()
        run(
            """
            Sub Increment(ByRef x)
                x = x + 1
            End Sub
            
            value = 5
            Increment value
            WScript.Echo value
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '6'

    def test_byval_does_not_modify_original(self):
        output = io.StringIO()
        run(
            """
            Sub TryToModify(ByVal x)
                x = x + 1
            End Sub
            
            value = 5
            TryToModify value
            WScript.Echo value
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '5'

    def test_default_is_byref(self):
        output = io.StringIO()
        # Note: Using Call keyword to avoid ambiguity with member access chains
        run(
            """
            Sub DoubleVal(x)
                x = x * 2
            End Sub
            
            value = 10
            Call DoubleVal(value)
            WScript.Echo value
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '20'

    def test_byref_with_expression(self):
        output = io.StringIO()
        run(
            """
            Sub Increment(ByRef x)
                x = x + 1
            End Sub
            
            value = 5
            Increment value + 0
            WScript.Echo value
        """,
            output_stream=output,
        )
        # When passing an expression to ByRef, it should not modify the original
        assert output.getvalue().strip() == '5'

    def test_mixed_byref_byval(self):
        output = io.StringIO()
        run(
            """
            Sub Process(ByRef refVar, ByVal valVar)
                refVar = refVar + 1
                valVar = valVar + 1
                WScript.Echo "Inside: " & refVar & ", " & valVar
            End Sub
            
            a = 10
            b = 20
            Process a, b
            WScript.Echo "Outside: " & a & ", " & b
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['Inside: 11, 21', 'Outside: 11, 20']

class TestInterpreterProcedureScoping:
    """Test procedure scoping rules."""

    def test_access_outer_variable(self):
        output = io.StringIO()
        run(
            """
            x = 10
            
            Sub ShowX
                WScript.Echo x
            End Sub
            
            ShowX
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '10'

    def test_local_shadows_outer(self):
        output = io.StringIO()
        # Note: Using Call keyword to avoid ambiguity
        run(
            """
            x = 10
            
            Sub TestShadow
                Dim x
                x = 20
                WScript.Echo x
            End Sub
            
            Call TestShadow
            WScript.Echo x
        """,
            output_stream=output,
        )
        lines = output.getvalue().strip().split('\n')
        assert lines == ['20', '10']

    def test_modify_outer_without_dim(self):
        output = io.StringIO()
        run(
            """
            x = 10
            
            Sub ModifyX
                x = 20
            End Sub
            
            Call ModifyX
            WScript.Echo x
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '20'

    def test_procedure_recursion(self):
        output = io.StringIO()
        run(
            """
            Function Factorial(n)
                If n <= 1 Then
                    Factorial = 1
                Else
                    Factorial = n * Factorial(n - 1)
                End If
            End Function
            
            WScript.Echo Factorial(5)
        """,
            output_stream=output,
        )
        assert output.getvalue().strip() == '120'
