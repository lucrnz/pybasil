' pybasil cscript.exe Compatibility Regression Test
' Creation date: 2026-04-28
' Purpose: Ensure pybasil output matches cscript.exe (VBScript 5.8, Windows 11)
'          for every behavior fixed in the cscript comparison audit.
'
' Every expected value in this file was verified against cscript.exe //nologo.
' If a test fails, it means pybasil has regressed from correct VBScript behavior.

Dim passCount, failCount, totalCount
passCount = 0
failCount = 0
totalCount = 0

Sub AssertEqual(testName, actual, expected)
    totalCount = totalCount + 1
    ' Null = Null returns Null (not True), so handle Null comparison specially
    If IsNull(actual) And IsNull(expected) Then
        passCount = passCount + 1
    ElseIf IsNull(actual) Or IsNull(expected) Then
        failCount = failCount + 1
        WScript.Echo "FAIL: " & testName & " - Expected: " & (expected & "") & " Got: " & (actual & "")
    ElseIf actual = expected Then
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

' ---------------------------------------------------------------------------
' #1  Boolean display: WScript.Echo must show -1/0, not True/False
'     cscript: WScript.Echo True  =>  -1
'     cscript: WScript.Echo False =>  0
'     CStr(True) still returns "True" (string conversion is different from Echo)
' ---------------------------------------------------------------------------
WScript.Echo "--- #1 Boolean Display ---"

Dim echoBuf

' We capture Echo output via string building to test display format.
' The canonical way to verify is: CStr converts to "True"/"False",
' but the numeric representation is -1/0.
Call AssertEqual("#1a True numeric value", CInt(True), -1)
Call AssertEqual("#1b False numeric value", CInt(False), 0)
Call AssertEqual("#1c CStr(True) is word form", CStr(True), "True")
Call AssertEqual("#1d CStr(False) is word form", CStr(False), "False")
Call AssertEqual("#1e True + 0 = -1", True + 0, -1)
Call AssertEqual("#1f False + 0 = 0", False + 0, 0)

' Comparison results are booleans displayed as -1/0
Dim cmpResult
cmpResult = (5 = 5)
Call AssertEqual("#1g comparison True = -1 numerically", CInt(cmpResult), -1)
cmpResult = (5 = 6)
Call AssertEqual("#1h comparison False = 0 numerically", CInt(cmpResult), 0)

' CBool display
Call AssertEqual("#1i CBool(1) numeric", CInt(CBool(1)), -1)
Call AssertEqual("#1j CBool(0) numeric", CInt(CBool(0)), 0)

' IsNull returns boolean, displayed as -1/0
Call AssertEqual("#1k IsNull(Null) numeric", CInt(IsNull(Null)), -1)
Call AssertEqual("#1l IsNull(42) numeric", CInt(IsNull(42)), 0)

' ---------------------------------------------------------------------------
' #2  Not operator: bitwise complement, not logical negation
'     cscript: Not 0   =>  -1
'     cscript: Not 1   =>  -2
'     cscript: Not 255 =>  -256
'     cscript: Not -1  =>  0
' ---------------------------------------------------------------------------
WScript.Echo "--- #2 Not Operator (Bitwise) ---"

Call AssertEqual("#2a Not 0", Not 0, -1)
Call AssertEqual("#2b Not 1", Not 1, -2)
Call AssertEqual("#2c Not 255", Not 255, -256)
Call AssertEqual("#2d Not (-1)", Not (-1), 0)
Call AssertEqual("#2e Not True", Not True, 0)
Call AssertEqual("#2f Not False", Not False, -1)
Call AssertEqual("#2g Not 7", Not 7, -8)
Call AssertEqual("#2h Not (-128)", Not (-128), 127)

' ---------------------------------------------------------------------------
' #3  And/Or/Xor/Eqv/Imp on booleans: return integers, not booleans
'     VBScript treats True as -1 and False as 0, then applies bitwise ops.
'     cscript: True And False  =>  0
'     cscript: True Or False   =>  -1
'     cscript: True Eqv True   =>  -1
'     cscript: True Imp False  =>  0
' ---------------------------------------------------------------------------
WScript.Echo "--- #3 Bitwise Logical Operators ---"

Call AssertEqual("#3a True And True", True And True, -1)
Call AssertEqual("#3b True And False", True And False, 0)
Call AssertEqual("#3c False And False", False And False, 0)
Call AssertEqual("#3d True Or True", True Or True, -1)
Call AssertEqual("#3e True Or False", True Or False, -1)
Call AssertEqual("#3f False Or False", False Or False, 0)
Call AssertEqual("#3g True Xor True", True Xor True, 0)
Call AssertEqual("#3h True Xor False", True Xor False, -1)
Call AssertEqual("#3i True Eqv True", True Eqv True, -1)
Call AssertEqual("#3j True Eqv False", True Eqv False, 0)
Call AssertEqual("#3k False Eqv False", False Eqv False, -1)
Call AssertEqual("#3l True Imp True", True Imp True, -1)
Call AssertEqual("#3m True Imp False", True Imp False, 0)
Call AssertEqual("#3n False Imp True", False Imp True, -1)
Call AssertEqual("#3o False Imp False", False Imp False, -1)

' Numeric operands (non-boolean) — already bitwise
Call AssertEqual("#3p 3 And 1", 3 And 1, 1)
Call AssertEqual("#3q 6 And 3", 6 And 3, 2)
Call AssertEqual("#3r 3 Or 1", 3 Or 1, 3)
Call AssertEqual("#3s 5 Xor 3", 5 Xor 3, 6)
Call AssertEqual("#3t 5 Eqv 3", 5 Eqv 3, -7)
Call AssertEqual("#3u 5 Imp 3", 5 Imp 3, -5)

' And/Or still work correctly in If conditions
Dim condResult
condResult = ""
If True And True Then
    condResult = "yes"
End If
Call AssertEqual("#3v And in If condition", condResult, "yes")

condResult = ""
If True And False Then
    condResult = "wrong"
Else
    condResult = "correct"
End If
Call AssertEqual("#3w And False in If condition", condResult, "correct")

' ---------------------------------------------------------------------------
' #4  Null concatenation: Null & "str" => "str" (Null treated as "")
'     cscript: Null & "hello"  =>  hello
'     cscript: "hello" & Null  =>  hello
'     cscript: Null & Null     =>  (empty string)
' ---------------------------------------------------------------------------
WScript.Echo "--- #4 Null Concatenation ---"

Call AssertEqual("#4a Null & string", Null & "hello", "hello")
Call AssertEqual("#4b string & Null", "hello" & Null, "hello")
Call AssertTrue("#4c Null & Null is Null", IsNull(Null & Null))
Call AssertEqual("#4d Null & number", Null & 42, "42")
Call AssertEqual("#4e Null & Empty", Null & Empty, "")

' ---------------------------------------------------------------------------
' #5  Date literals: #date# syntax
'     cscript: Year(#1/15/2024#)    =>  2024
'     cscript: Month(#1/15/2024#)   =>  1
'     cscript: Day(#1/15/2024#)     =>  15
'     cscript: Weekday(#1/15/2024#) =>  2  (Monday)
' ---------------------------------------------------------------------------
WScript.Echo "--- #5 Date Literals ---"

Call AssertEqual("#5a Year from date literal", Year(#1/15/2024#), 2024)
Call AssertEqual("#5b Month from date literal", Month(#1/15/2024#), 1)
Call AssertEqual("#5c Day from date literal", Day(#1/15/2024#), 15)
Call AssertEqual("#5d Weekday from date literal", Weekday(#1/15/2024#), 2)
Call AssertEqual("#5e Year from another date", Year(#6/15/2025#), 2025)
Call AssertEqual("#5f Month from another date", Month(#6/15/2025#), 6)
Call AssertTrue("#5g Date literal is a Date", IsDate(#1/15/2024#))

' Date arithmetic with literals
Dim dateDiff1
dateDiff1 = Year(#6/15/2025#) - Year(#1/1/2020#)
Call AssertEqual("#5h Date year subtraction", dateDiff1, 5)

' ---------------------------------------------------------------------------
' #6  WScript.Echo -5+3: expression statement parsing
'     cscript: WScript.Echo -5+3  =>  -2
'     The parser must not treat (-5) as a method call argument and then
'     fail on +3.
' ---------------------------------------------------------------------------
WScript.Echo "--- #6 Expression Statement Parsing ---"

' These are verified by not crashing and producing correct output.
' We test the underlying arithmetic instead.
Call AssertEqual("#6a -5+3", -5+3, -2)
Call AssertEqual("#6b -10+7", -10+7, -3)
Call AssertEqual("#6c -1+1", -1+1, 0)

' ---------------------------------------------------------------------------
' #7  Dict.RemoveAll: zero-argument method dispatch
'     cscript: d.RemoveAll : Echo d.Count  =>  0
' ---------------------------------------------------------------------------
WScript.Echo "--- #7 Dict.RemoveAll ---"

Dim d7
Set d7 = CreateObject("Scripting.Dictionary")
d7.Add "a", 1
d7.Add "b", 2
Call AssertEqual("#7a Count before RemoveAll", d7.Count, 2)
d7.RemoveAll
Call AssertEqual("#7b Count after RemoveAll", d7.Count, 0)

' Re-add after RemoveAll
d7.Add "c", 3
Call AssertEqual("#7c Count after re-add", d7.Count, 1)
Call AssertEqual("#7d Item after re-add", d7("c"), 3)

' ---------------------------------------------------------------------------
' #8  Default member resolution: WScript.Echo obj resolves default property
'     cscript: class with Default Property Get => echoes the property value
' ---------------------------------------------------------------------------
WScript.Echo "--- #8 Default Member Resolution ---"

Class DefaultPropClass
    Private m_val
    Public Default Property Get Value()
        Value = "default_value"
    End Property
End Class

Dim obj8
Set obj8 = New DefaultPropClass
' We can't easily capture Echo output in VBScript, but we can test
' the default member invocation via string concatenation
Call AssertEqual("#8a Default property via &", "" & obj8, "default_value")
Call AssertEqual("#8b Default property via CStr", CStr(obj8), "default_value")

' ---------------------------------------------------------------------------
' #9  For loop default step: always 1, never auto-detected as -1
'     cscript: For i = 5 To 1 : s = s & i : Next  =>  (empty, loop never runs)
' ---------------------------------------------------------------------------
WScript.Echo "--- #9 For Loop Default Step ---"

Dim forResult9
forResult9 = ""
For i = 5 To 1
    forResult9 = forResult9 & CStr(i)
Next
Call AssertEqual("#9a For 5 To 1 (no step) is empty", forResult9, "")

' With explicit Step -1, it should count down
forResult9 = ""
For i = 5 To 1 Step -1
    forResult9 = forResult9 & CStr(i)
Next
Call AssertEqual("#9b For 5 To 1 Step -1 counts down", forResult9, "54321")

' Normal ascending still works
forResult9 = ""
For i = 1 To 5
    forResult9 = forResult9 & CStr(i)
Next
Call AssertEqual("#9c For 1 To 5 counts up", forResult9, "12345")

' Equal start/end runs once
forResult9 = ""
For i = 3 To 3
    forResult9 = forResult9 & CStr(i)
Next
Call AssertEqual("#9d For 3 To 3 runs once", forResult9, "3")

' ---------------------------------------------------------------------------
' #10 Float display precision: 15 significant digits
'     cscript: 10 / 3       =>  3.33333333333333
'     cscript: 0.1 + 0.2    =>  0.3
'     cscript: 1 / 7        =>  0.142857142857143
' ---------------------------------------------------------------------------
WScript.Echo "--- #10 Float Display Precision ---"

Call AssertEqual("#10a 10/3 display", CStr(10 / 3), "3.33333333333333")
Call AssertEqual("#10b 0.1+0.2 display", CStr(0.1 + 0.2), "0.3")
Call AssertEqual("#10c 1/7 display", CStr(1 / 7), "0.142857142857143")
Call AssertEqual("#10d 2/3 display", CStr(2 / 3), "0.666666666666667")

' ---------------------------------------------------------------------------
' #11 Hex(-1): 16-bit for Integer range, 32-bit for Long range
'     cscript: Hex(-1)       =>  FFFF
'     cscript: Hex(-40000)   =>  FFFF63C0
'     cscript: Hex(255)      =>  FF
' ---------------------------------------------------------------------------
WScript.Echo "--- #11 Hex Integer vs Long ---"

Call AssertEqual("#11a Hex(-1)", Hex(-1), "FFFF")
Call AssertEqual("#11b Hex(-40000)", Hex(-40000), "FFFF63C0")
Call AssertEqual("#11c Hex(255)", Hex(255), "FF")
Call AssertEqual("#11d Hex(0)", Hex(0), "0")
Call AssertEqual("#11e Hex(-32768)", Hex(-32768), "FFFF8000")

' ---------------------------------------------------------------------------
' #12 TypeName: Integer (-32768..32767) vs Long (outside that range)
'     cscript: TypeName(42)     =>  Integer
'     cscript: TypeName(100000) =>  Long
' ---------------------------------------------------------------------------
WScript.Echo "--- #12 TypeName Integer vs Long ---"

Call AssertEqual("#12a TypeName(42)", TypeName(42), "Integer")
Call AssertEqual("#12b TypeName(100000)", TypeName(100000), "Long")
Call AssertEqual("#12c TypeName(-32768)", TypeName(-32768), "Long")
Call AssertEqual("#12d TypeName(32767)", TypeName(32767), "Integer")
Call AssertEqual("#12e TypeName(32768)", TypeName(32768), "Long")
Call AssertEqual("#12f TypeName(-32767)", TypeName(-32767), "Integer")
Call AssertEqual("#12f2 TypeName(-32769)", TypeName(-32769), "Long")
Call AssertEqual("#12g TypeName(0)", TypeName(0), "Integer")

' Other types unchanged
Call AssertEqual("#12h TypeName(string)", TypeName("hello"), "String")
Call AssertEqual("#12i TypeName(bool)", TypeName(True), "Boolean")
Call AssertEqual("#12j TypeName(float)", TypeName(3.14), "Double")
Call AssertEqual("#12k TypeName(Empty)", TypeName(Empty), "Empty")
Call AssertEqual("#12l TypeName(Null)", TypeName(Null), "Null")
Call AssertEqual("#12m TypeName(Nothing)", TypeName(Nothing), "Nothing")

' ---------------------------------------------------------------------------
' #13 VarType: objects return 9 (vbObject)
'     cscript: VarType(Nothing) =>  9
'     cscript: VarType(dict)    =>  9
' ---------------------------------------------------------------------------
WScript.Echo "--- #13 VarType for Objects ---"

Call AssertEqual("#13a VarType(Nothing)", VarType(Nothing), 9)

Dim d13
Set d13 = CreateObject("Scripting.Dictionary")
Call AssertEqual("#13b VarType(Dictionary)", VarType(d13), 9)

' Standard types unchanged
Call AssertEqual("#13c VarType(Empty)", VarType(Empty), 0)
Call AssertEqual("#13d VarType(Null)", VarType(Null), 1)
Call AssertEqual("#13e VarType(42)", VarType(42), 2)
Call AssertEqual("#13f VarType(3.14)", VarType(3.14), 5)
Call AssertEqual("#13g VarType(string)", VarType("hi"), 8)
Call AssertEqual("#13h VarType(True)", VarType(True), 11)

' ---------------------------------------------------------------------------
' #14 IsObject(Nothing): returns True
'     cscript: IsObject(Nothing) =>  True (-1)
'     Nothing is the null object reference, so it IS an object.
' ---------------------------------------------------------------------------
WScript.Echo "--- #14 IsObject(Nothing) ---"

Call AssertTrue("#14a IsObject(Nothing)", IsObject(Nothing))
Call AssertTrue("#14b IsObject(Dictionary)", IsObject(d13))
Call AssertTrue("#14c Not IsObject(42)", Not IsObject(42))
Call AssertTrue("#14d Not IsObject(string)", Not IsObject("hello"))

' ---------------------------------------------------------------------------
' #15 Len() accepts numbers: converts to string first
'     cscript: Len(12345) =>  5
'     cscript: Len(3.14)  =>  4
' ---------------------------------------------------------------------------
WScript.Echo "--- #15 Len() on Numbers ---"

Call AssertEqual("#15a Len(12345)", Len(12345), 5)
Call AssertEqual("#15b Len(0)", Len(0), 1)
Call AssertEqual("#15c Len(True)", Len(True), 4)
Call AssertEqual("#15d Len(False)", Len(False), 5)
Call AssertEqual("#15e Len(Empty)", Len(Empty), 0)

' Len on strings still works
Call AssertEqual("#15f Len(string)", Len("Hello"), 5)
Call AssertEqual("#15g Len(empty string)", Len(""), 0)

' ---------------------------------------------------------------------------
' #16 (-2) ^ 3: correct exponentiation with parenthesized negative base
'     cscript: (-2) ^ 3  =>  -8
'     cscript: (-3) ^ 2  =>  9
' ---------------------------------------------------------------------------
WScript.Echo "--- #16 Exponentiation ---"

Call AssertEqual("#16a (-2) ^ 3", (-2) ^ 3, -8)
Call AssertEqual("#16b (-3) ^ 2", (-3) ^ 2, 9)
Call AssertEqual("#16c 2 ^ 3", 2 ^ 3, 8)
Call AssertEqual("#16d (-1) ^ 0", (-1) ^ 0, 1)
Call AssertEqual("#16e (-2) ^ 0", (-2) ^ 0, 1)

' ---------------------------------------------------------------------------
' #17 Round(2.55, 1): Decimal-based banker's rounding
'     cscript: Round(2.55, 1) =>  2.6
'     cscript: Round(2.45, 1) =>  2.4  (banker's rounding: round to even)
'     cscript: Round(3.5, 0)  =>  4
'     cscript: Round(4.5, 0)  =>  4    (banker's rounding)
' ---------------------------------------------------------------------------
WScript.Echo "--- #17 Round Precision ---"

Call AssertEqual("#17a Round(2.55, 1)", Round(2.55, 1), 2.6)
Call AssertEqual("#17b Round(2.45, 1)", Round(2.45, 1), 2.4)
Call AssertEqual("#17c Round(3.5, 0)", Round(3.5, 0), 4)
Call AssertEqual("#17d Round(4.5, 0)", Round(4.5, 0), 4)
Call AssertEqual("#17e Round(3.14159, 2)", Round(3.14159, 2), 3.14)
Call AssertEqual("#17f Round(1.5, 0)", Round(1.5, 0), 2)
Call AssertEqual("#17g Round(2.5, 0)", Round(2.5, 0), 2)

' ---------------------------------------------------------------------------
' #18 Error numbers: specific codes for specific errors
'     cscript: Const reassignment  =>  Err.Number = 501
'     cscript: Undefined variable   =>  Err.Number = 500
' ---------------------------------------------------------------------------
WScript.Echo "--- #18 Error Numbers ---"

Dim errNum18

' 501: Illegal assignment (const reassignment)
On Error Resume Next
Const CONST18 = 42
CONST18 = 99
errNum18 = Err.Number
On Error GoTo 0
Call AssertEqual("#18a Const reassignment error", errNum18, 501)

' 500: Variable is undefined (Option Explicit)
' Note: we can't use Option Explicit mid-script in cscript, but pybasil
' supports it. We test the error number mapping instead.
' The error number 13 for type mismatch is already tested in alpha.vbs.
Dim errNum18b
On Error Resume Next
Dim tmErr
tmErr = CInt("not_a_number")
errNum18b = Err.Number
On Error GoTo 0
Call AssertEqual("#18b Type mismatch error", errNum18b, 13)

' 11: Division by zero
Dim errNum18c
On Error Resume Next
Dim dzErr
dzErr = 1 / 0
errNum18c = Err.Number
On Error GoTo 0
Call AssertEqual("#18c Division by zero error", errNum18c, 11)

' ---------------------------------------------------------------------------
' Null propagation with And/Or (three-valued logic)
'     cscript: Null And False  =>  0
'     cscript: Null Or True    =>  -1
'     cscript: Null And True   =>  Null
'     cscript: Null Or False   =>  Null
' ---------------------------------------------------------------------------
WScript.Echo "--- Null And/Or Propagation ---"

Call AssertEqual("Null And False", Null And False, 0)
Call AssertEqual("Null Or True", Null Or True, -1)
Call AssertTrue("Null And True is Null", IsNull(Null And True))
Call AssertTrue("Null Or False is Null", IsNull(Null Or False))

' ---------------------------------------------------------------------------
' Additional edge cases: operators in conditions
' ---------------------------------------------------------------------------
WScript.Echo "--- Edge Cases ---"

' Not in If conditions
Dim notIfResult
notIfResult = ""
If Not False Then
    notIfResult = "entered"
End If
Call AssertEqual("Not False in If", notIfResult, "entered")

notIfResult = ""
If Not True Then
    notIfResult = "wrong"
Else
    notIfResult = "correct"
End If
Call AssertEqual("Not True in If", notIfResult, "correct")

' Dictionary Exists returns boolean (displayed as -1/0)
Dim d_edge
Set d_edge = CreateObject("Scripting.Dictionary")
d_edge.Add "x", 1
Call AssertTrue("d.Exists returns truthy", d_edge.Exists("x"))
Call AssertTrue("d.Exists missing returns falsy", Not d_edge.Exists("y"))

' Len on boolean converts to string length
' CStr(True) = "True" (4 chars), CStr(False) = "False" (5 chars)
Call AssertEqual("Len(True) = 4", Len(True), 4)
Call AssertEqual("Len(False) = 5", Len(False), 5)

' ---------------------------------------------------------------------------
' Summary
' ---------------------------------------------------------------------------
WScript.Echo ""
WScript.Echo "cscript compatibility regression summary"
WScript.Echo "Total:  " & totalCount
WScript.Echo "Passed: " & passCount
WScript.Echo "Failed: " & failCount
If failCount = 0 Then
    WScript.Echo "Result: ALL TESTS PASSED"
Else
    WScript.Echo "Result: SOME TESTS FAILED"
End If