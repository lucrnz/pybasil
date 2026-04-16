Dim total
total = 0

' Arithmetic-heavy loop
Dim i
For i = 1 To 5000
    total = total + i * 2 - 1
    If total > 1000000 Then
        total = total Mod 1000000
    End If
Next

' String operations
Dim s
s = ""
For i = 1 To 200
    s = s & "x"
Next
Dim sLen
sLen = Len(s)

' Function calls
Function Factorial(n)
    If n <= 1 Then
        Factorial = 1
    Else
        Factorial = n * Factorial(n - 1)
    End If
End Function

Dim factResult
factResult = Factorial(12)

' Nested conditionals
Dim grade
For i = 1 To 500
    grade = i Mod 100
    If grade >= 90 Then
        total = total + 4
    ElseIf grade >= 80 Then
        total = total + 3
    ElseIf grade >= 70 Then
        total = total + 2
    Else
        total = total + 1
    End If
Next

' Select Case
Dim category
For i = 1 To 500
    category = i Mod 5
    Select Case category
        Case 0
            total = total + 10
        Case 1
            total = total + 20
        Case 2
            total = total + 30
        Case 3
            total = total + 40
        Case Else
            total = total + 50
    End Select
Next

' Array operations
Dim arr(49)
For i = 0 To 49
    arr(i) = i * i
Next
Dim arrSum
arrSum = 0
For i = 0 To 49
    arrSum = arrSum + arr(i)
Next

' Dictionary operations
Dim dict
Set dict = CreateObject("Scripting.Dictionary")
For i = 1 To 200
    dict.Add "key" & CStr(i), i * 10
Next
Dim dictSum
dictSum = 0
For i = 1 To 200
    dictSum = dictSum + dict.Item("key" & CStr(i))
Next

' Nested loops with arithmetic
Dim j
For i = 1 To 100
    For j = 1 To 50
        total = total + (i * j) Mod 97
    Next
Next

' Do While loop
Dim counter
counter = 0
Do While counter < 1000
    counter = counter + 1
    total = total + counter Mod 7
Loop

' String builtins
Dim testStr
testStr = "Hello World VBScript Performance Test"
For i = 1 To 200
    Dim u, l, trimmed, leftPart, rightPart, midPart
    u = UCase(testStr)
    l = LCase(testStr)
    trimmed = Trim("  hello  ")
    leftPart = Left(testStr, 5)
    rightPart = Right(testStr, 4)
    midPart = Mid(testStr, 7, 5)
Next

' Numeric builtins
For i = 1 To 500
    Dim absVal, intVal, cintVal, cdblVal
    absVal = Abs(-42)
    intVal = Int(3.7)
    cintVal = CInt("123")
    cdblVal = CDbl("3.14")
Next

' Sub calls
Sub AddToTotal(ByVal amount)
    total = total + amount
End Sub

For i = 1 To 500
    AddToTotal i
Next

' Boolean operations
Dim boolResult
For i = 1 To 500
    boolResult = (i > 250) And (i < 750)
    boolResult = boolResult Or (i Mod 2 = 0)
    If Not boolResult Then
        total = total + 1
    End If
Next

' InStr
For i = 1 To 200
    Dim pos
    pos = InStr(1, testStr, "World")
Next

' Type checking
For i = 1 To 200
    Dim isNum, isEmpty2
    isNum = IsNumeric(42)
    isNum = IsNumeric("hello")
    isEmpty2 = IsEmpty(total)
Next

' Class operations - Property Get/Let, methods, field access
Class Counter
    Private m_val
    Private Sub Class_Initialize()
        m_val = 0
    End Sub
    Public Property Get Value
        Value = m_val
    End Property
    Public Property Let Value(v)
        m_val = v
    End Property
    Public Function Increment()
        m_val = m_val + 1
        Increment = m_val
    End Function
    Public Function Add(n)
        m_val = m_val + n
        Add = m_val
    End Function
End Class

Class Point
    Public X
    Public Y
    Public Function DistSq()
        DistSq = X * X + Y * Y
    End Function
End Class

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

Dim c
Set c = New Counter
For i = 1 To 500
    c.Value = i
    total = total + c.Value
    c.Increment
    c.Add 2
Next

Dim p
Set p = New Point
For i = 1 To 500
    p.X = i
    p.Y = i * 2
    total = total + p.DistSq()
Next

Dim person
Set person = New Person
For i = 1 To 200
    person.FirstName = "First" & CStr(i)
    person.LastName = "Last" & CStr(i)
    Dim fullName
    fullName = person.FullName
Next

WScript.Echo total
