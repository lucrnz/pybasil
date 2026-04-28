"""Tests for the VBScript interpreter."""

import io
from pybasil import (
    Interpreter,
    parse,
)


class TestWithStatement:
    """Test With...End With blocks."""

    def test_with_class_property_get(self):
        program = parse('''
        Class Foo
            Private m_x
            Property Get X
                X = m_x
            End Property
            Property Let X(v)
                m_x = v
            End Property
        End Class
        Dim obj
        Set obj = New Foo
        obj.X = 10
        With obj
            result = .X
        End With
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 10

    def test_with_class_property_let(self):
        program = parse('''
        Class Foo
            Private m_x
            Property Get X
                X = m_x
            End Property
            Property Let X(v)
                m_x = v
            End Property
        End Class
        Dim obj
        Set obj = New Foo
        With obj
            .X = 42
        End With
        result = obj.X
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 42

    def test_with_class_method_call(self):
        output = io.StringIO()
        program = parse('''
        Class Foo
            Public Function Add(a, b)
                Add = a + b
            End Function
        End Class
        Dim obj
        Set obj = New Foo
        With obj
            result = .Add(3, 4)
        End With
        ''')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        assert interp._environment.get('result') == 7

    def test_with_dictionary(self):
        program = parse('''
        Dim d
        Set d = CreateObject("Scripting.Dictionary")
        With d
            .Add "a", 1
            .Add "b", 2
            result = .Count
        End With
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 2

    def test_with_nested(self):
        program = parse('''
        Class Inner
            Public Value
        End Class
        Class Outer
            Public Name
        End Class
        Dim o, i
        Set o = New Outer
        Set i = New Inner
        o.Name = "outer"
        i.Value = 99
        With o
            .Name = "modified"
            With i
                .Value = 100
            End With
        End With
        r1 = o.Name
        r2 = i.Value
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('r1') == 'modified'
        assert interp._environment.get('r2') == 100

    def test_with_field_access(self):
        program = parse('''
        Class Point
            Public X
            Public Y
        End Class
        Dim p
        Set p = New Point
        With p
            .X = 10
            .Y = 20
        End With
        result = p.X + p.Y
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 30

    def test_with_method_no_args(self):
        program = parse('''
        Class Counter
            Private m_val
            Private Sub Class_Initialize()
                m_val = 0
            End Sub
            Public Function Increment()
                m_val = m_val + 1
                Increment = m_val
            End Function
        End Class
        Dim c
        Set c = New Counter
        With c
            .Increment
            .Increment
            result = .Increment()
        End With
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 3

    def test_with_wscript(self):
        output = io.StringIO()
        program = parse('''
        With WScript
            .Echo "hello"
        End With
        ''')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        assert output.getvalue().strip() == 'hello'

    def test_with_dictionary_item(self):
        program = parse('''
        Dim d
        Set d = CreateObject("Scripting.Dictionary")
        d.Add "x", 42
        With d
            result = .Item("x")
        End With
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 42
