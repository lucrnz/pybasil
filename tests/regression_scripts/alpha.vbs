' pybasil Regression Test Suite
' Creation date: 2026-04-10
' Purpose: Comprehensive regression test of all supported interpreter features

Dim passCount, failCount, totalCount
passCount = 0
failCount = 0
totalCount = 0

Sub AssertEqual(testName, actual, expected)
    totalCount = totalCount + 1
    If actual = expected Then
        passCount = passCount + 1
    Else
        failCount = failCount + 1
        WScript.Echo "FAIL: " & testName & " - Expected: " & CStr(expected) & " Got: " & CStr(actual)
    End If
End Sub

Sub AssertTrue(testName, condition)
    totalCount = totalCount + 1
    If condition Then
        passCount = passCount + 1
    Else
        failCount = failCount + 1
        WScript.Echo "FAIL: " & testName & " - Expected True Got False"
    End If
End Sub

Sub AssertFalse(testName, condition)
    totalCount = totalCount + 1
    If Not condition Then
        passCount = passCount + 1
    Else
        failCount = failCount + 1
        WScript.Echo "FAIL: " & testName & " - Expected False Got True"
    End If
End Sub

' 1. LITERALS

WScript.Echo "--- Literals ---"

Dim litInt
litInt = 42
Call AssertEqual("Integer literal", litInt, 42)

Dim litFloat
litFloat = 3.14
Call AssertEqual("Float literal", litFloat, 3.14)

Dim litSci
litSci = 1.5e10
Call AssertTrue("Scientific notation", litSci = 15000000000)

Dim litStr
litStr = "Hello"
Call AssertEqual("String literal", litStr, "Hello")

Dim litBoolTrue
litBoolTrue = True
Call AssertTrue("Boolean True literal", litBoolTrue = True)

Dim litBoolFalse
litBoolFalse = False
Call AssertTrue("Boolean False literal", litBoolFalse = False)

Dim litNothing
Set litNothing = Nothing
Call AssertTrue("Nothing literal", litNothing Is Nothing)

Dim litEmpty
litEmpty = Empty
Call AssertTrue("Empty literal", IsEmpty(litEmpty))

Dim litNull
litNull = Null
Call AssertTrue("Null literal", IsNull(litNull))

' 2. VARIABLES

WScript.Echo "--- Variables ---"

Dim varA, varB, varC
varA = 10
varB = varA
Call AssertEqual("Variable assignment and lookup", varB, 10)

Dim implicitVar
implicitVar = undeclaredVar
Call AssertTrue("Implicit variable creation (Empty)", IsEmpty(undeclaredVar))

varC = 100
Call AssertEqual("Let keyword assignment", varC, 100)

Dim UPPER, lower
UPPER = 42
lower = UPPER
Call AssertEqual("Case-insensitive variables", lower, 42)

Set obj = CreateObject("Scripting.FileSystemObject")
Call AssertTrue("Set/CreateObject", Not (obj Is Nothing))

' 3. ARITHMETIC OPERATORS

WScript.Echo "--- Arithmetic Operators ---"

Call AssertEqual("Addition", 5 + 3, 8)
Call AssertEqual("Subtraction", 10 - 4, 6)
Call AssertEqual("Multiplication", 6 * 7, 42)
Call AssertEqual("Division", 15 / 3, 5.0)
Call AssertEqual("Integer division", 17 \ 5, 3)
Call AssertEqual("Modulo", 17 Mod 5, 2)
Call AssertEqual("Exponentiation", 2 ^ 10, 1024)
Call AssertEqual("Negation", -5, -5)
Call AssertEqual("Unary plus", +5, 5)
Call AssertEqual("Complex expression (2+3*4-1)", 2 + 3 * 4 - 1, 13)
Call AssertEqual("Parentheses ((2+3)*4)", (2 + 3) * 4, 20)

' 4. STRING OPERATIONS

WScript.Echo "--- String Operations ---"

Call AssertEqual("String concatenation (&)", "Hello" & " " & "World", "Hello World")
Call AssertEqual("String-number concatenation", "Value: " & 42, "Value: 42")
Call AssertEqual("Numeric string + number", "5" + 3, 8)

' 5. COMPARISON OPERATORS

WScript.Echo "--- Comparison Operators ---"

Call AssertTrue("Equals (5=5)", 5 = 5)
Call AssertFalse("Equals (5=6)", 5 = 6)
Call AssertTrue("Not equals (5<>6)", 5 <> 6)
Call AssertTrue("Less than (3<5)", 3 < 5)
Call AssertTrue("Greater than (7>5)", 7 > 5)
Call AssertTrue("Less equal (5<=5)", 5 <= 5)
Call AssertTrue("Greater equal (5>=5)", 5 >= 5)
Call AssertTrue("String comparison", "abc" < "def")
Call AssertTrue("Is operator (Nothing Is Nothing)", Nothing Is Nothing)

' 6. LOGICAL OPERATORS

WScript.Echo "--- Logical Operators ---"

Call AssertTrue("And (True And True)", True And True)
Call AssertFalse("And (True And False)", True And False)
Call AssertTrue("Or (False Or True)", False Or True)
Call AssertFalse("Or (False Or False)", False Or False)
Call AssertFalse("Not (Not True)", Not True)
Call AssertTrue("Not (Not False)", Not False)
Call AssertFalse("Xor (True Xor True)", True Xor True)
Call AssertTrue("Xor (True Xor False)", True Xor False)
Call AssertTrue("Eqv (True Eqv True)", True Eqv True)
Call AssertFalse("Eqv (True Eqv False)", True Eqv False)
Call AssertTrue("Imp (True Imp True)", True Imp True)
Call AssertFalse("Imp (True Imp False)", True Imp False)
Call AssertTrue("Imp (False Imp True)", False Imp True)
Call AssertTrue("Imp (False Imp False)", False Imp False)

' 7. TYPE COERCION & EDGE CASES

WScript.Echo "--- Type Coercion & Edge Cases ---"

Call AssertEqual("Empty + 5", Empty + 5, 5)
Call AssertEqual("Empty & 'text'", Empty & "text", "text")
Call AssertTrue("Null propagation", IsNull(Null + 5))
Call AssertTrue("Null comparison propagates Null", IsNull(Null = Null))
Call AssertEqual("String to number addition", "10" + 5, 15)
Call AssertEqual("Boolean in arithmetic (True+1)", True + 1, 0)

' 8. IF...THEN...ELSE

WScript.Echo "--- If...Then...Else ---"

Dim ifResult

If True Then
    ifResult = "yes"
End If
Call AssertEqual("If True executes body", ifResult, "yes")

ifResult = "none"
If False Then
    ifResult = "wrong"
End If
Call AssertEqual("If False skips body", ifResult, "none")

If True Then
    ifResult = "then"
Else
    ifResult = "else"
End If
Call AssertEqual("If True goes Then", ifResult, "then")

If False Then
    ifResult = "then"
Else
    ifResult = "else"
End If
Call AssertEqual("If False goes Else", ifResult, "else")

Dim elseifResult
elseifResult = 0
If elseifResult = 1 Then
    elseifResult = 100
ElseIf elseifResult = 0 Then
    elseifResult = 200
Else
    elseifResult = 300
End If
Call AssertEqual("ElseIf matches second branch", elseifResult, 200)

elseifResult = 0
If elseifResult = 1 Then
    elseifResult = 100
ElseIf elseifResult = 2 Then
    elseifResult = 200
Else
    elseifResult = 300
End If
Call AssertEqual("ElseIf falls to Else", elseifResult, 300)

Dim nestedIfResult
nestedIfResult = 5
If nestedIfResult > 0 Then
    If nestedIfResult > 10 Then
        nestedIfResult = "big"
    Else
        nestedIfResult = "small"
    End If
End If
Call AssertEqual("Nested If", nestedIfResult, "small")

IF TRUE THEN
    ifResult = "uppercase"
END IF
Call AssertEqual("Case-insensitive IF", ifResult, "uppercase")

' 9. SELECT CASE

WScript.Echo "--- Select Case ---"

Dim scResult
scResult = 2
Select Case scResult
    Case 1
        scResult = "one"
    Case 2
        scResult = "two"
    Case 3
        scResult = "three"
End Select
Call AssertEqual("Select Case basic match", scResult, "two")

scResult = 99
Select Case scResult
    Case 1
        scResult = "one"
    Case Else
        scResult = "other"
End Select
Call AssertEqual("Select Case Else", scResult, "other")

scResult = 5
Select Case scResult
    Case 1, 2, 3
        scResult = "small"
    Case 4, 5, 6
        scResult = "medium"
    Case Else
        scResult = "large"
End Select
Call AssertEqual("Select Case comma values", scResult, "medium")

Dim score, grade
score = 85
Select Case True
    Case score >= 90
        grade = "A"
    Case score >= 80
        grade = "B"
    Case score >= 70
        grade = "C"
    Case Else
        grade = "F"
End Select
Call AssertEqual("Select Case True with relational", grade, "B")

Dim fruit
fruit = "Banana"
Select Case fruit
    Case "Apple", "Pear"
        fruit = "pome"
    Case "Banana", "Mango"
        fruit = "tropical"
    Case Else
        fruit = "unknown"
End Select
Call AssertEqual("Select Case string match", fruit, "tropical")

' 10. FOR...NEXT

WScript.Echo "--- For...Next ---"

Dim forSum
forSum = 0
For i = 1 To 5
    forSum = forSum + i
Next
Call AssertEqual("For basic sum (1..5)", forSum, 15)

Dim forStepSum
forStepSum = 0
For i = 0 To 10 Step 2
    forStepSum = forStepSum + i
Next
Call AssertEqual("For with Step 2", forStepSum, 30)

Dim forNegSum
forNegSum = 0
For i = 5 To 1 Step -1
    forNegSum = forNegSum + i
Next
Call AssertEqual("For negative step", forNegSum, 15)

Dim forExitResult
forExitResult = 0
For i = 1 To 100
    If i = 5 Then
        Exit For
    End If
    forExitResult = forExitResult + 1
Next
Call AssertEqual("Exit For", forExitResult, 4)

Dim nestedForSum
nestedForSum = 0
For i = 1 To 3
    For j = 1 To 3
        nestedForSum = nestedForSum + 1
    Next
Next
Call AssertEqual("Nested For loops", nestedForSum, 9)

Dim afterLoop
For afterLoop = 1 To 5
Next
Call AssertEqual("Variable after For loop", afterLoop, 6)

' 11. FOR EACH

WScript.Echo "--- For Each ---"

Dim forEachSum, forEachItem
forEachSum = 0
Dim forEachArr(4)
forEachArr(0) = 10
forEachArr(1) = 20
forEachArr(2) = 30
forEachArr(3) = 40
forEachArr(4) = 50
For Each forEachItem In forEachArr
    forEachSum = forEachSum + forEachItem
Next
Call AssertEqual("For Each array iteration", forEachSum, 150)

' 12. WHILE...WEND

WScript.Echo "--- While...Wend ---"

Dim whileCount
whileCount = 0
Dim whileI
whileI = 0
While whileI < 5
    whileCount = whileCount + 1
    whileI = whileI + 1
Wend
Call AssertEqual("While basic", whileCount, 5)

whileCount = 0
While False
    whileCount = 1
Wend
Call AssertEqual("While False never executes", whileCount, 0)

Dim nestedWhileCount
nestedWhileCount = 0
whileI = 0
While whileI < 3
    Dim whileJ
    whileJ = 0
    While whileJ < 2
        nestedWhileCount = nestedWhileCount + 1
        whileJ = whileJ + 1
    Wend
    whileI = whileI + 1
Wend
Call AssertEqual("Nested While", nestedWhileCount, 6)

' 13. DO...LOOP

WScript.Echo "--- Do...Loop ---"

Dim doCount
doCount = 0
Dim doI
doI = 0
Do While doI < 5
    doCount = doCount + 1
    doI = doI + 1
Loop
Call AssertEqual("Do While pre-test", doCount, 5)

doCount = 0
doI = 0
Do Until doI >= 5
    doCount = doCount + 1
    doI = doI + 1
Loop
Call AssertEqual("Do Until pre-test", doCount, 5)

doCount = 0
doI = 0
Do
    doCount = doCount + 1
    doI = doI + 1
Loop While doI < 5
Call AssertEqual("Do...Loop While post-test", doCount, 5)

doCount = 0
doI = 0
Do
    doCount = doCount + 1
    doI = doI + 1
Loop Until doI >= 5
Call AssertEqual("Do...Loop Until post-test", doCount, 5)

doCount = 0
doI = 10
Do
    doCount = doCount + 1
Loop While doI < 5
Call AssertEqual("Do...Loop While executes once", doCount, 1)

doCount = 0
doI = 10
Do
    doCount = doCount + 1
Loop Until doI > 5
Call AssertEqual("Do...Loop Until executes once", doCount, 1)

Dim exitDoCount
exitDoCount = 0
Do While True
    exitDoCount = exitDoCount + 1
    If exitDoCount >= 3 Then
        Exit Do
    End If
Loop
Call AssertEqual("Exit Do", exitDoCount, 3)

' 14. ARRAYS

WScript.Echo "--- Arrays ---"

Dim fixedArr(4)
fixedArr(0) = 10
fixedArr(1) = 20
fixedArr(2) = 30
fixedArr(3) = 40
fixedArr(4) = 50
Call AssertEqual("Fixed array element access", fixedArr(2), 30)
Call AssertEqual("UBound", UBound(fixedArr), 4)
Call AssertEqual("LBound", LBound(fixedArr), 0)

Dim dynamicArr()
ReDim dynamicArr(2)
dynamicArr(0) = "a"
dynamicArr(1) = "b"
dynamicArr(2) = "c"
Call AssertEqual("Dynamic array (ReDim)", dynamicArr(1), "b")

ReDim Preserve dynamicArr(4)
dynamicArr(3) = "d"
dynamicArr(4) = "e"
Call AssertEqual("ReDim Preserve keeps data", dynamicArr(2), "c")
Call AssertEqual("ReDim Preserve new element", dynamicArr(4), "e")

Dim matrix(2, 2)
matrix(0, 0) = 1
matrix(1, 1) = 5
matrix(2, 2) = 9
Call AssertEqual("Multi-dimensional array", matrix(1, 1), 5)

Dim arrLiteral
arrLiteral = Array(10, 20, 30)
Call AssertEqual("Array() function", arrLiteral(1), 20)

' 15. PROCEDURES (Sub & Function)

WScript.Echo "--- Procedures (Sub & Function) ---"

Sub Greet(name)
    WScript.Echo "Hello, " & name
End Sub

Greet "Regression"

Function Add(a, b)
    Add = a + b
End Function

Call AssertEqual("Function with params", Add(3, 4), 7)

Function DoubleVal(x)
    DoubleVal = x * 2
End Function

Call AssertEqual("Function in expression", DoubleVal(5) + 1, 11)

Function EarlyReturn()
    EarlyReturn = 1
    Exit Function
    EarlyReturn = 2
End Function

Dim earlyResult
earlyResult = EarlyReturn()
Call AssertEqual("Exit Function", earlyResult, 1)

Sub ExitSubTest()
    WScript.Echo "Before Exit Sub"
    Exit Sub
    WScript.Echo "Should not appear"
End Sub

Call ExitSubTest()

Function Square(x)
    Square = x * x
End Function

Function SumOfSquares(a, b)
    SumOfSquares = Square(a) + Square(b)
End Function

Call AssertEqual("Nested function calls", SumOfSquares(3, 4), 25)

Function Factorial(n)
    If n <= 1 Then
        Factorial = 1
    Else
        Factorial = n * Factorial(n - 1)
    End If
End Function

Call AssertEqual("Recursive function", Factorial(5), 120)

Dim scopeTestOuter
scopeTestOuter = 10

Sub ScopeTest()
    Dim scopeTestOuter
    scopeTestOuter = 20
End Sub

Call ScopeTest()
Call AssertEqual("Sub local scope (Dim shadows)", scopeTestOuter, 10)

Dim scopeModifyTest
scopeModifyTest = 10

Sub ModifyOuter()
    scopeModifyTest = 20
End Sub

Call ModifyOuter()
Call AssertEqual("Sub modifies outer (no Dim)", scopeModifyTest, 20)

' 16. BYREF / BYVAL

WScript.Echo "--- ByRef / ByVal ---"

Sub IncrementByRef(ByRef x)
    x = x + 1
End Sub

Dim byRefVal
byRefVal = 5
IncrementByRef byRefVal
Call AssertEqual("ByRef modifies original", byRefVal, 6)

Sub TryModifyByVal(ByVal x)
    x = x + 1
End Sub

Dim byValVal
byValVal = 5
TryModifyByVal byValVal
Call AssertEqual("ByVal does not modify original", byValVal, 5)

Sub DefaultByRef(x)
    x = x * 2
End Sub

Dim defaultRefVal
defaultRefVal = 10
Call DefaultByRef(defaultRefVal)
Call AssertEqual("Default is ByRef", defaultRefVal, 20)

Sub MixedParams(ByRef refVar, ByVal valVar)
    refVar = refVar + 1
    valVar = valVar + 1
End Sub

Dim mixA, mixB
mixA = 10
mixB = 20
MixedParams mixA, mixB
Call AssertEqual("Mixed ByRef/ByVal - ByRef modified", mixA, 11)
Call AssertEqual("Mixed ByRef/ByVal - ByVal not modified", mixB, 20)

Dim byRefExprVal
byRefExprVal = 5
IncrementByRef byRefExprVal + 0
Call AssertEqual("ByRef with expression (no modify)", byRefExprVal, 5)

' 17. ERROR HANDLING

WScript.Echo "--- Error Handling ---"

Dim errNum1, errDesc1, errSrc1
On Error Resume Next
Dim divByZero
divByZero = 1 / 0
errNum1 = Err.Number
errDesc1 = Err.Description
errSrc1 = Err.Source
On Error GoTo 0
Call AssertEqual("Err.Number after division by zero", errNum1, 11)
Call AssertTrue("Err.Description has text", Len(errDesc1) > 0)
Call AssertTrue("Err.Source has VBScript", InStr(errSrc1, "VBScript") > 0)

Dim errAfterClear
On Error Resume Next
divByZero = 1 / 0
Err.Clear
errAfterClear = Err.Number
On Error GoTo 0
Call AssertEqual("Err.Clear resets number", errAfterClear, 0)

Dim errRaisedNum, errRaisedSrc, errRaisedDesc
On Error Resume Next
Err.Raise 100, "TestSource", "Test Description"
errRaisedNum = Err.Number
errRaisedSrc = Err.Source
errRaisedDesc = Err.Description
On Error GoTo 0
Call AssertEqual("Err.Raise sets Number", errRaisedNum, 100)
Call AssertEqual("Err.Raise sets Source", errRaisedSrc, "TestSource")
Call AssertEqual("Err.Raise sets Description", errRaisedDesc, "Test Description")

Dim errInitNum
errInitNum = Err.Number
Call AssertEqual("Err.Number initially 0", errInitNum, 0)

Dim errTypeMismatch
On Error Resume Next
Dim typeMismatch
typeMismatch = CInt("abc")
errTypeMismatch = Err.Number
On Error GoTo 0
Call AssertEqual("Type mismatch error number", errTypeMismatch, 13)

Dim firstErr, secondErr
On Error Resume Next
divByZero = 1 / 0
firstErr = Err.Number
On Error GoTo 0

On Error Resume Next
notnumErr = CInt("notnum")
secondErr = Err.Number
On Error GoTo 0
Call AssertEqual("First error tracked", firstErr, 11)
Call AssertEqual("Second error tracked", secondErr, 13)

' 18. BUILT-IN STRING FUNCTIONS

WScript.Echo "--- String Functions ---"

Call AssertEqual("Len", Len("Hello"), 5)
Call AssertEqual("Left", Left("Hello", 3), "Hel")
Call AssertEqual("Right", Right("Hello", 3), "llo")
Call AssertEqual("Mid", Mid("Hello", 2, 3), "ell")
Call AssertEqual("Trim", Trim("  Hello  "), "Hello")
Call AssertEqual("LTrim", LTrim("  Hello  "), "Hello  ")
Call AssertEqual("RTrim", RTrim("  Hello  "), "  Hello")
Call AssertEqual("UCase", UCase("hello"), "HELLO")
Call AssertEqual("LCase", LCase("HELLO"), "hello")
Call AssertTrue("InStr found", InStr("Hello World", "World") > 0)

Dim replacedStr
replacedStr = Replace("Hello World", "World", "VBScript")
Call AssertEqual("Replace", replacedStr, "Hello VBScript")

Dim splitArr
splitArr = Split("a,b,c", ",")
Call AssertEqual("Split creates array", TypeName(splitArr), "Variant()")

Call AssertEqual("Join", Join(Array("x", "y", "z"), "-"), "x-y-z")

' 19. BUILT-IN CONVERSION/TYPE FUNCTIONS

WScript.Echo "--- Conversion/Type Functions ---"

Call AssertEqual("CStr", CStr(42), "42")
Call AssertEqual("CInt", CInt(3.7), 4)
Call AssertEqual("CDbl", CDbl("3.14"), 3.14)
Call AssertTrue("CBool(1)", CBool(1))
Call AssertTrue("IsNumeric string number", IsNumeric("123"))
Call AssertFalse("IsNumeric non-number", IsNumeric("abc"))
Call AssertTrue("IsEmpty", IsEmpty(Empty))
Call AssertTrue("IsNull", IsNull(Null))
Call AssertTrue("IsArray", IsArray(Array(1, 2)))
Call AssertEqual("TypeName string", TypeName("Hello"), "String")
Call AssertEqual("TypeName integer", TypeName(42), "Integer")

' 20. BUILT-IN MATH FUNCTIONS

WScript.Echo "--- Math Functions ---"

Call AssertEqual("Abs", Abs(-5), 5)
Call AssertEqual("Sqr", Sqr(16), 4.0)
Call AssertEqual("Int", Int(3.7), 3)
Call AssertEqual("Fix", Fix(3.7), 3)
Call AssertEqual("Round", Round(3.14159, 2), 3.14)

' 21. COMMENTS

WScript.Echo "--- Comments ---"

' This is a comment and should do nothing
Dim commentTest
commentTest = 5 Rem This is a Rem comment
Call AssertEqual("Comment after statement", commentTest, 5)
Rem Full line Rem comment
Call AssertEqual("After Rem comment", commentTest, 5)

' 22. WScript.Echo

WScript.Echo "--- WScript.Echo ---"

Dim echoOutput
echoOutput = "Test Output"
WScript.Echo echoOutput
WScript.Echo "Multiple", "Args", 42

Call AssertTrue("Echo executed (no crash)", True)

WScript.Echo "Regression test summary"
WScript.Echo "Total:  " & totalCount
WScript.Echo "Passed: " & passCount
WScript.Echo "Failed: " & failCount
If failCount = 0 Then
    WScript.Echo "Result:  ALL TESTS PASSED"
Else
    WScript.Echo "Result:  SOME TESTS FAILED"
End If
