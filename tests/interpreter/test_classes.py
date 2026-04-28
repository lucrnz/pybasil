"""Tests for the VBScript interpreter."""

import pytest
import io
from pybasil import (
    Interpreter,
    parse,
    run,
    VBScriptError,
)


class TestInterpreterClass:
    """Test VBScript Class support."""

    def test_class_basic_instantiation(self):
        output = io.StringIO()
        run('''
            Class Dog
                Public Name
            End Class
            Dim d
            Set d = New Dog
            d.Name = "Rex"
            WScript.Echo d.Name
        ''', output_stream=output)
        assert output.getvalue().strip() == 'Rex'

    def test_class_private_field_with_methods(self):
        output = io.StringIO()
        run('''
            Class Dog
                Private m_name
                Public Sub Init(name)
                    m_name = name
                End Sub
                Public Function GetName()
                    GetName = m_name
                End Function
            End Class
            Dim d
            Set d = New Dog
            d.Init "Rex"
            WScript.Echo d.GetName()
        ''', output_stream=output)
        assert output.getvalue().strip() == 'Rex'

    def test_class_property_get_let(self):
        output = io.StringIO()
        run('''
            Class Person
                Private m_name
                Property Get Name()
                    Name = m_name
                End Property
                Property Let Name(val)
                    m_name = val
                End Property
            End Class
            Dim p
            Set p = New Person
            p.Name = "Alice"
            WScript.Echo p.Name
        ''', output_stream=output)
        assert output.getvalue().strip() == 'Alice'

    def test_class_multiple_properties(self):
        output = io.StringIO()
        run('''
            Class Person
                Private m_name
                Private m_age
                Property Get Name()
                    Name = m_name
                End Property
                Property Let Name(val)
                    m_name = val
                End Property
                Property Get Age()
                    Age = m_age
                End Property
                Property Let Age(val)
                    m_age = val
                End Property
            End Class
            Dim p
            Set p = New Person
            p.Name = "Bob"
            p.Age = 25
            WScript.Echo p.Name & " is " & p.Age
        ''', output_stream=output)
        assert output.getvalue().strip() == 'Bob is 25'

    def test_class_me_keyword(self):
        output = io.StringIO()
        run('''
            Class Greeter
                Private m_name
                Property Let Name(val)
                    m_name = val
                End Property
                Property Get Name()
                    Name = m_name
                End Property
                Public Function SayHello()
                    SayHello = "Hello, " & Me.Name
                End Function
            End Class
            Dim g
            Set g = New Greeter
            g.Name = "World"
            WScript.Echo g.SayHello()
        ''', output_stream=output)
        assert output.getvalue().strip() == 'Hello, World'

    def test_class_initialize(self):
        output = io.StringIO()
        run('''
            Class Counter
                Private m_count
                Private Sub Class_Initialize
                    m_count = 0
                End Sub
                Public Sub Increment()
                    m_count = m_count + 1
                End Sub
                Public Property Get Value()
                    Value = m_count
                End Property
            End Class
            Dim c
            Set c = New Counter
            WScript.Echo c.Value
            c.Increment
            c.Increment
            c.Increment
            WScript.Echo c.Value
        ''', output_stream=output)
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0', '3']

    def test_class_function_with_args(self):
        output = io.StringIO()
        run('''
            Class Calculator
                Public Function Add(a, b)
                    Add = a + b
                End Function
                Public Function Multiply(a, b)
                    Multiply = a * b
                End Function
            End Class
            Dim calc
            Set calc = New Calculator
            WScript.Echo calc.Add(3, 4)
            WScript.Echo calc.Multiply(5, 6)
        ''', output_stream=output)
        lines = output.getvalue().strip().split('\n')
        assert lines == ['7', '30']

    def test_class_multiple_instances(self):
        output = io.StringIO()
        run('''
            Class Dog
                Public Name
            End Class
            Dim d1, d2
            Set d1 = New Dog
            Set d2 = New Dog
            d1.Name = "Rex"
            d2.Name = "Buddy"
            WScript.Echo d1.Name
            WScript.Echo d2.Name
        ''', output_stream=output)
        lines = output.getvalue().strip().split('\n')
        assert lines == ['Rex', 'Buddy']

    def test_class_typename(self):
        output = io.StringIO()
        run('''
            Class MyClass
                Public Value
            End Class
            Dim obj
            Set obj = New MyClass
            WScript.Echo TypeName(obj)
        ''', output_stream=output)
        assert output.getvalue().strip() == 'MyClass'

    def test_class_isobject(self):
        program = parse('''
            Class Foo
                Public X
            End Class
            Dim f
            Set f = New Foo
            result = IsObject(f)
        ''')
        interpreter = Interpreter()
        interpreter.interpret(program)
        assert interpreter._environment.get('result') is True

    def test_class_exit_property(self):
        output = io.StringIO()
        run('''
            Class Clamped
                Private m_val
                Property Let Value(v)
                    If v < 0 Then
                        m_val = 0
                        Exit Property
                    End If
                    m_val = v
                End Property
                Property Get Value()
                    Value = m_val
                End Property
            End Class
            Dim c
            Set c = New Clamped
            c.Value = -5
            WScript.Echo c.Value
            c.Value = 42
            WScript.Echo c.Value
        ''', output_stream=output)
        lines = output.getvalue().strip().split('\n')
        assert lines == ['0', '42']

    def test_class_exit_function(self):
        output = io.StringIO()
        run('''
            Class Validator
                Public Function IsPositive(n)
                    If n <= 0 Then
                        IsPositive = False
                        Exit Function
                    End If
                    IsPositive = True
                End Function
            End Class
            Dim v
            Set v = New Validator
            WScript.Echo v.IsPositive(5)
            WScript.Echo v.IsPositive(-3)
        ''', output_stream=output)
        lines = output.getvalue().strip().split('\n')
        assert lines == ['True', 'False']

    def test_class_with_array_field(self):
        output = io.StringIO()
        run('''
            Class Container
                Private m_items(4)
                Public Sub SetItem(i, val)
                    m_items(i) = val
                End Sub
                Public Function GetItem(i)
                    GetItem = m_items(i)
                End Function
            End Class
            Dim c
            Set c = New Container
            c.SetItem 0, "alpha"
            c.SetItem 1, "beta"
            WScript.Echo c.GetItem(0)
            WScript.Echo c.GetItem(1)
        ''', output_stream=output)
        lines = output.getvalue().strip().split('\n')
        assert lines == ['alpha', 'beta']

    def test_class_method_modifies_state(self):
        output = io.StringIO()
        run('''
            Class Stack
                Private m_data(99)
                Private m_top
                Private Sub Class_Initialize
                    m_top = -1
                End Sub
                Public Sub Push(val)
                    m_top = m_top + 1
                    m_data(m_top) = val
                End Sub
                Public Function Pop()
                    Pop = m_data(m_top)
                    m_top = m_top - 1
                End Function
                Public Property Get Count()
                    Count = m_top + 1
                End Property
            End Class
            Dim s
            Set s = New Stack
            s.Push 10
            s.Push 20
            s.Push 30
            WScript.Echo s.Count
            WScript.Echo s.Pop()
            WScript.Echo s.Pop()
            WScript.Echo s.Count
        ''', output_stream=output)
        lines = output.getvalue().strip().split('\n')
        assert lines == ['3', '30', '20', '1']

    def test_class_sub_with_args_no_parens(self):
        """Test calling Sub method without parentheses."""
        output = io.StringIO()
        run('''
            Class Logger
                Public Sub Log(msg)
                    WScript.Echo "LOG: " & msg
                End Sub
            End Class
            Dim l
            Set l = New Logger
            l.Log "hello"
        ''', output_stream=output)
        assert output.getvalue().strip() == 'LOG: hello'

    def test_class_dim_field_default(self):
        """Test that Dim inside class defaults to public."""
        output = io.StringIO()
        run('''
            Class Thing
                Dim X
            End Class
            Dim t
            Set t = New Thing
            t.X = 42
            WScript.Echo t.X
        ''', output_stream=output)
        assert output.getvalue().strip() == '42'

    def test_class_case_insensitive_members(self):
        output = io.StringIO()
        run('''
            Class Foo
                Public bar
            End Class
            Dim f
            Set f = New Foo
            f.BAR = "test"
            WScript.Echo f.bar
        ''', output_stream=output)
        assert output.getvalue().strip() == 'test'

    def test_class_case_insensitive_methods(self):
        output = io.StringIO()
        run('''
            Class Foo
                Public Function GetValue()
                    GetValue = 42
                End Function
            End Class
            Dim f
            Set f = New Foo
            WScript.Echo f.GETVALUE()
        ''', output_stream=output)
        assert output.getvalue().strip() == '42'

    def test_class_method_with_loop(self):
        output = io.StringIO()
        run('''
            Class Summation
                Public Function SumTo(n)
                    Dim total, i
                    total = 0
                    For i = 1 To n
                        total = total + i
                    Next
                    SumTo = total
                End Function
            End Class
            Dim s
            Set s = New Summation
            WScript.Echo s.SumTo(10)
        ''', output_stream=output)
        assert output.getvalue().strip() == '55'

    def test_class_new_expression(self):
        """Test that New creates a class instance."""
        program = parse('''
            Class Widget
                Public Label
            End Class
            Dim w
            Set w = New Widget
        ''')
        interpreter = Interpreter()
        interpreter.interpret(program)
        from pybasil.runtime import VBScriptClassInstance
        assert isinstance(interpreter._environment.get('w'), VBScriptClassInstance)

    def test_class_not_defined_error(self):
        with pytest.raises(VBScriptError, match='Class not defined'):
            run('Set x = New NonExistent')

    def test_class_implicit_property_get(self):
        """Bare property name inside method resolves without Me."""
        output = io.StringIO()
        run('''
            Class Person
                Private m_name
                Property Let Name(v)
                    m_name = v
                End Property
                Property Get Name
                    Name = m_name
                End Property
                Public Function Greet()
                    Greet = "Hi " & Name
                End Function
            End Class
            Dim p
            Set p = New Person
            p.Name = "Alice"
            WScript.Echo p.Greet()
        ''', output_stream=output)
        assert output.getvalue().strip() == 'Hi Alice'

    def test_class_implicit_method_call(self):
        """Bare function call inside method resolves to instance method."""
        output = io.StringIO()
        run('''
            Class Calc
                Private m_val
                Property Let Value(v)
                    m_val = v
                End Property
                Public Function Double()
                    Double = m_val * 2
                End Function
                Public Function DoubledPlusTen()
                    DoubledPlusTen = Double() + 10
                End Function
            End Class
            Dim c
            Set c = New Calc
            c.Value = 7
            WScript.Echo c.DoubledPlusTen()
        ''', output_stream=output)
        assert output.getvalue().strip() == '24'

    def test_class_implicit_property_chain(self):
        """Property calling another property without Me."""
        output = io.StringIO()
        run('''
            Class Person
                Private m_first
                Private m_last
                Property Let FirstName(v)
                    m_first = v
                End Property
                Property Get FirstName
                    FirstName = m_first
                End Property
                Property Let LastName(v)
                    m_last = v
                End Property
                Property Get LastName
                    LastName = m_last
                End Property
                Property Get FullName
                    FullName = FirstName & " " & LastName
                End Property
            End Class
            Dim p
            Set p = New Person
            p.FirstName = "Bob"
            p.LastName = "Smith"
            WScript.Echo p.FullName
        ''', output_stream=output)
        assert output.getvalue().strip() == 'Bob Smith'

    def test_class_array_of_objects(self):
        """Store class instances in an array and call methods."""
        output = io.StringIO()
        run('''
            Class Item
                Public Label
                Public Function Describe()
                    Describe = "Item: " & Label
                End Function
            End Class
            Dim a, b
            Set a = New Item
            Set b = New Item
            a.Label = "A"
            b.Label = "B"
            Dim items(1)
            Set items(0) = a
            Set items(1) = b
            Dim i
            For i = 0 To UBound(items)
                WScript.Echo items(i).Describe()
            Next
        ''', output_stream=output)
        lines = output.getvalue().strip().split('\n')
        assert lines == ['Item: A', 'Item: B']

    def test_class_set_nothing(self):
        """Set instance to Nothing and verify."""
        output = io.StringIO()
        run('''
            Class Obj
                Public X
            End Class
            Dim o
            Set o = New Obj
            o.X = 1
            WScript.Echo o.X
            Set o = Nothing
            WScript.Echo TypeName(o)
        ''', output_stream=output)
        lines = output.getvalue().strip().split('\n')
        assert lines == ['1', 'Nothing']

    def test_class_property_let_validation(self):
        """Property Let with validation logic using builtins."""
        output = io.StringIO()
        run('''
            Class Bounded
                Private m_val
                Private Sub Class_Initialize()
                    m_val = 0
                End Sub
                Property Let Value(v)
                    If IsNumeric(v) And v >= 0 Then
                        m_val = CInt(v)
                    Else
                        m_val = 0
                    End If
                End Property
                Property Get Value
                    Value = m_val
                End Property
            End Class
            Dim b
            Set b = New Bounded
            b.Value = 42
            WScript.Echo b.Value
            b.Value = -5
            WScript.Echo b.Value
        ''', output_stream=output)
        lines = output.getvalue().strip().split('\n')
        assert lines == ['42', '0']

    def test_class_sub_calls_property_implicitly(self):
        """Sub method accesses a property without Me."""
        output = io.StringIO()
        run('''
            Class Greeter
                Private m_name
                Property Let Name(v)
                    m_name = v
                End Property
                Property Get Name
                    Name = m_name
                End Property
                Public Sub SayHi()
                    WScript.Echo "Hi, " & Name & "!"
                End Sub
            End Class
            Dim g
            Set g = New Greeter
            g.Name = "World"
            g.SayHi
        ''', output_stream=output)
        assert output.getvalue().strip() == 'Hi, World!'
