' pybasil Regression Test Suite - Execute, ExecuteGlobal & Eval
' Creation date: 2026-04-16
' Purpose: Comprehensive regression test of dynamic code execution

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

' =========================================================================
'  EVAL - Basic Arithmetic
' =========================================================================
WScript.Echo "--- Eval: Basic Arithmetic ---"

Call AssertEqual("Eval addition", Eval("2 + 3"), 5)
Call AssertEqual("Eval subtraction", Eval("10 - 4"), 6)
Call AssertEqual("Eval multiplication", Eval("6 * 7"), 42)
Call AssertEqual("Eval division", Eval("15 / 3"), 5.0)
Call AssertEqual("Eval integer division", Eval("17 \ 5"), 3)
Call AssertEqual("Eval modulo", Eval("17 Mod 5"), 2)
Call AssertEqual("Eval exponentiation", Eval("2 ^ 10"), 1024)
Call AssertEqual("Eval complex expression", Eval("(2 + 3) * 4 - 1"), 19)
Call AssertEqual("Eval negative number", Eval("-5"), -5)
Call AssertEqual("Eval nested parens", Eval("((2 + 3) * (4 - 1))"), 15)

' =========================================================================
'  EVAL - String Expressions
' =========================================================================
WScript.Echo "--- Eval: String Expressions ---"

Call AssertEqual("Eval string literal", Eval("""Hello"""), "Hello")
Call AssertEqual("Eval string concat", Eval("""Hello"" & "" "" & ""World"""), "Hello World")
Call AssertEqual("Eval string-number concat", Eval("""Value: "" & 42"), "Value: 42")

' =========================================================================
'  EVAL - Boolean Expressions
' =========================================================================
WScript.Echo "--- Eval: Boolean Expressions ---"

Call AssertTrue("Eval True", Eval("True"))
Call AssertFalse("Eval False", Eval("False"))
Call AssertFalse("Eval And", Eval("True And False"))
Call AssertTrue("Eval Or", Eval("True Or False"))
Call AssertFalse("Eval Not True", Eval("Not True"))
Call AssertTrue("Eval comparison >", Eval("5 > 3"))
Call AssertFalse("Eval comparison > false", Eval("3 > 5"))

' =========================================================================
'  EVAL - Variable References from Current Scope
' =========================================================================
WScript.Echo "--- Eval: Variable References ---"

Dim evalX
evalX = 42
Call AssertEqual("Eval variable ref", Eval("evalX"), 42)
Call AssertEqual("Eval variable in expr", Eval("evalX + 8"), 50)
Call AssertEqual("Eval variable multiply", Eval("evalX * 2"), 84)

Dim evalA, evalB
evalA = 10
evalB = 20
Call AssertEqual("Eval two variables", Eval("evalA + evalB"), 30)
Call AssertEqual("Eval complex var expr", Eval("evalA * evalB + 5"), 205)

' =========================================================================
'  EVAL - With Built-in Functions
' =========================================================================
WScript.Echo "--- Eval: Built-in Functions ---"

Call AssertEqual("Eval Len()", Eval("Len(""Hello"")"), 5)
Call AssertEqual("Eval UCase()", Eval("UCase(""test"")"), "TEST")
Call AssertEqual("Eval LCase()", Eval("LCase(""HELLO"")"), "hello")
Call AssertEqual("Eval Abs()", Eval("Abs(-7)"), 7)
Call AssertEqual("Eval CStr()", Eval("CStr(42)"), "42")
Call AssertEqual("Eval CInt()", Eval("CInt(3.7)"), 4)
Call AssertTrue("Eval IsNumeric()", Eval("IsNumeric(""123"")"))
Call AssertTrue("Eval IsEmpty()", Eval("IsEmpty(Empty)"))

' =========================================================================
'  EVAL - With User-defined Functions
' =========================================================================
WScript.Echo "--- Eval: User-defined Functions ---"

Function Triple(n)
    Triple = n * 3
End Function

Call AssertEqual("Eval user function", Eval("Triple(5)"), 15)
Call AssertEqual("Eval user function in expr", Eval("Triple(3) + 1"), 10)

Function AddTwo(a, b)
    AddTwo = a + b
End Function

Call AssertEqual("Eval function two args", Eval("AddTwo(10, 20)"), 30)

' =========================================================================
'  EVAL - Dynamic Expression Building
' =========================================================================
WScript.Echo "--- Eval: Dynamic Expressions ---"

Dim dynOp
dynOp = "+"
Call AssertEqual("Eval dynamic op +", Eval("10 " & dynOp & " 5"), 15)
dynOp = "-"
Call AssertEqual("Eval dynamic op -", Eval("10 " & dynOp & " 5"), 5)
dynOp = "*"
Call AssertEqual("Eval dynamic op *", Eval("10 " & dynOp & " 5"), 50)

Dim dynVal
dynVal = 7
Call AssertEqual("Eval dynamic value", Eval(CStr(dynVal) & " * 6"), 42)

' =========================================================================
'  EVAL - Special Values
' =========================================================================
WScript.Echo "--- Eval: Special Values ---"

Call AssertTrue("Eval Nothing Is Nothing", Eval("Nothing Is Nothing"))
Call AssertTrue("Eval IsNull(Null)", Eval("IsNull(Null)"))
Call AssertTrue("Eval IsEmpty(Empty)", Eval("IsEmpty(Empty)"))

' =========================================================================
'  EXECUTE - Variable Assignment
' =========================================================================
WScript.Echo "--- Execute: Variable Assignment ---"

Execute "execResult1 = 42"
Call AssertEqual("Execute simple assign", execResult1, 42)

Execute "execResult2 = ""Hello Execute"""
Call AssertEqual("Execute string assign", execResult2, "Hello Execute")

Execute "execResult3 = True"
Call AssertTrue("Execute boolean assign", execResult3)

Execute "execResult4 = 3.14"
Call AssertEqual("Execute float assign", execResult4, 3.14)

' =========================================================================
'  EXECUTE - Multiple Statements (colon-separated)
' =========================================================================
WScript.Echo "--- Execute: Multiple Statements ---"

Execute "execMultiA = 10 : execMultiB = 20"
Call AssertEqual("Execute multi stmt A", execMultiA, 10)
Call AssertEqual("Execute multi stmt B", execMultiB, 20)

Execute "execMultiC = 5 : execMultiD = execMultiC * 3"
Call AssertEqual("Execute multi stmt dependency", execMultiD, 15)

' =========================================================================
'  EXECUTE - Dim Statements
' =========================================================================
WScript.Echo "--- Execute: Dim Statements ---"

Execute "Dim execDimVar : execDimVar = 99"
Call AssertEqual("Execute Dim and assign", execDimVar, 99)

Execute "Dim execDimArr(2) : execDimArr(0) = 10 : execDimArr(1) = 20 : execDimArr(2) = 30"
Call AssertEqual("Execute Dim array", execDimArr(1), 20)

' =========================================================================
'  EXECUTE - Define Sub
' =========================================================================
WScript.Echo "--- Execute: Define Sub ---"

Dim execSubCalled
execSubCalled = False
Execute "Sub ExecDefinedSub() : execSubCalled = True : End Sub"
Call ExecDefinedSub()
Call AssertTrue("Execute-defined Sub callable", execSubCalled)

Dim execSubResult
execSubResult = ""
Execute "Sub ExecSubWithParam(msg) : execSubResult = msg : End Sub"
Call ExecSubWithParam("hello from exec")
Call AssertEqual("Execute-defined Sub with param", execSubResult, "hello from exec")

' =========================================================================
'  EXECUTE - Define Function
' =========================================================================
WScript.Echo "--- Execute: Define Function ---"

Execute "Function ExecAddTen(n) : ExecAddTen = n + 10 : End Function"
Call AssertEqual("Execute-defined Function", ExecAddTen(5), 15)

Execute "Function ExecConcat(a, b) : ExecConcat = a & b : End Function"
Call AssertEqual("Execute-defined Function concat", ExecConcat("foo", "bar"), "foobar")

' =========================================================================
'  EXECUTE - Control Flow
' =========================================================================
WScript.Echo "--- Execute: Control Flow ---"

Execute "If True Then : execIfResult = ""yes"" : End If"
Call AssertEqual("Execute If True", execIfResult, "yes")

Execute "If False Then : execIfElse = ""wrong"" : Else : execIfElse = ""correct"" : End If"
Call AssertEqual("Execute If Else", execIfElse, "correct")

Execute "execForSum = 0 : For execI = 1 To 5 : execForSum = execForSum + execI : Next"
Call AssertEqual("Execute For loop", execForSum, 15)

Execute "execDoCount = 0 : execDoI = 0 : Do While execDoI < 3 : execDoCount = execDoCount + 1 : execDoI = execDoI + 1 : Loop"
Call AssertEqual("Execute Do While", execDoCount, 3)

' =========================================================================
'  EXECUTE - Scope Visibility
' =========================================================================
WScript.Echo "--- Execute: Scope ---"

' Execute can modify existing variables in the current scope
Dim execScopeVar
execScopeVar = 100
Execute "execScopeVar = execScopeVar + 50"
Call AssertEqual("Execute modifies existing var", execScopeVar, 150)

' Execute can read variables from the current scope
Dim execReadVar
execReadVar = 25
Execute "execReadResult = execReadVar * 4"
Call AssertEqual("Execute reads current scope", execReadResult, 100)

' =========================================================================
'  EXECUTE - Inside a Procedure
' =========================================================================
WScript.Echo "--- Execute: Inside Procedure ---"

Sub TestExecuteInSub()
    Dim localVar
    localVar = 10
    Execute "localVar = localVar + 5"
    execInSubResult = localVar
End Sub

Dim execInSubResult
execInSubResult = 0
Call TestExecuteInSub()
Call AssertEqual("Execute inside Sub modifies local", execInSubResult, 15)

' =========================================================================
'  EXECUTE - Nested Execute
' =========================================================================
WScript.Echo "--- Execute: Nested ---"

Execute "Execute ""nestedExecVar = 42"""
Call AssertEqual("Nested Execute", nestedExecVar, 42)

' =========================================================================
'  EXECUTE - Edge Cases
' =========================================================================
WScript.Echo "--- Execute: Edge Cases ---"

' Empty string should be a no-op
On Error Resume Next
Execute ""
Dim execEmptyErr
execEmptyErr = Err.Number
On Error GoTo 0
Call AssertEqual("Execute empty string (no error)", execEmptyErr, 0)

' =========================================================================
'  EXECUTE - Error Handling
' =========================================================================
WScript.Echo "--- Execute: Error Handling ---"

' Error inside Execute propagates to caller's On Error Resume Next
Dim execErrDiv
On Error Resume Next
Execute "Dim execBadDiv : execBadDiv = 1 / 0"
execErrDiv = Err.Number
On Error GoTo 0
Call AssertEqual("Execute div-by-zero propagates", execErrDiv, 11)

' Type mismatch inside Execute
Dim execErrType
On Error Resume Next
Execute "execTypeFail = CInt(""abc"")"
execErrType = Err.Number
On Error GoTo 0
Call AssertEqual("Execute type mismatch propagates", execErrType, 13)

' On Error Resume Next declared inside Execute'd code
Dim execInternalErr
Execute "On Error Resume Next : Dim execDivZ : execDivZ = 1 / 0 : execInternalErr = Err.Number : On Error GoTo 0"
Call AssertEqual("Execute internal error handling", execInternalErr, 11)

' =========================================================================
'  EXECUTEGLOBAL - Define Function at Global Scope
' =========================================================================
WScript.Echo "--- ExecuteGlobal: Define Function ---"

ExecuteGlobal "Function GlobalSquare(n) : GlobalSquare = n * n : End Function"
Call AssertEqual("ExecuteGlobal Function", GlobalSquare(7), 49)

' =========================================================================
'  EXECUTEGLOBAL - Define Sub at Global Scope
' =========================================================================
WScript.Echo "--- ExecuteGlobal: Define Sub ---"

Dim globalSubRan
globalSubRan = False
ExecuteGlobal "Sub GlobalMarkDone() : globalSubRan = True : End Sub"
Call GlobalMarkDone()
Call AssertTrue("ExecuteGlobal Sub callable", globalSubRan)

' =========================================================================
'  EXECUTEGLOBAL - From Inside a Procedure (the key use-case)
' =========================================================================
WScript.Echo "--- ExecuteGlobal: From Procedure ---"

' Functions defined via ExecuteGlobal inside a Sub become globally callable
Sub SetupGlobalFunc()
    ExecuteGlobal "Function GlobalDouble(n) : GlobalDouble = n * 2 : End Function"
End Sub

Call SetupGlobalFunc()
Call AssertEqual("ExecuteGlobal func from Sub", GlobalDouble(21), 42)

' Subs defined via ExecuteGlobal inside a Sub become globally callable
Sub SetupGlobalSub()
    ExecuteGlobal "Sub GlobalSetFlag() : globalFlag = True : End Sub"
End Sub

Dim globalFlag
globalFlag = False
Call SetupGlobalSub()
Call GlobalSetFlag()
Call AssertTrue("ExecuteGlobal Sub from Sub", globalFlag)

' =========================================================================
'  EXECUTEGLOBAL - Variables at Global Scope
' =========================================================================
WScript.Echo "--- ExecuteGlobal: Variables ---"

ExecuteGlobal "globalTestVar1 = 777"
Call AssertEqual("ExecuteGlobal variable", globalTestVar1, 777)

Sub DefineGlobalVar()
    ExecuteGlobal "globalTestVar2 = 888"
End Sub

Call DefineGlobalVar()
Call AssertEqual("ExecuteGlobal var from Sub", globalTestVar2, 888)

' =========================================================================
'  EXECUTEGLOBAL - Define Class
' =========================================================================
WScript.Echo "--- ExecuteGlobal: Define Class ---"

ExecuteGlobal "Class DynPoint : Public X : Public Y : End Class"

Dim pt
Set pt = New DynPoint
pt.X = 3
pt.Y = 4
Call AssertEqual("ExecuteGlobal Class field X", pt.X, 3)
Call AssertEqual("ExecuteGlobal Class field Y", pt.Y, 4)

' Class with a method
ExecuteGlobal "Class DynCalc : Public Function Multiply(a, b) : Multiply = a * b : End Function : End Class"

Dim calc
Set calc = New DynCalc
Call AssertEqual("ExecuteGlobal Class method", calc.Multiply(6, 7), 42)

' =========================================================================
'  CROSS-FEATURE: Execute then Eval
' =========================================================================
WScript.Echo "--- Cross-feature: Execute + Eval ---"

' Execute defines a variable, Eval reads it
Execute "crossVar = 99"
Call AssertEqual("Execute var, Eval reads", Eval("crossVar"), 99)

' Execute defines a function, Eval calls it
Execute "Function CrossFunc(n) : CrossFunc = n + 1 : End Function"
Call AssertEqual("Execute func, Eval calls", Eval("CrossFunc(41)"), 42)

' Eval called inside Execute'd code
Execute "evalInsideExec = Eval(""2 + 3"")"
Call AssertEqual("Eval inside Execute", evalInsideExec, 5)

' =========================================================================
'  CROSS-FEATURE: Dynamic Dispatch via Eval
' =========================================================================
WScript.Echo "--- Cross-feature: Dynamic Dispatch ---"

Dim operations(2)
operations(0) = "+"
operations(1) = "-"
operations(2) = "*"

Dim expectedResults(2)
expectedResults(0) = 15
expectedResults(1) = 5
expectedResults(2) = 50

Dim oi
For oi = 0 To 2
    Dim dynResult
    dynResult = Eval("10 " & operations(oi) & " 5")
    Call AssertEqual("Dynamic dispatch op " & operations(oi), dynResult, expectedResults(oi))
Next

' =========================================================================
'  CROSS-FEATURE: ExecuteGlobal factory pattern
' =========================================================================
WScript.Echo "--- Cross-feature: ExecuteGlobal + Eval ---"

' Build and register functions dynamically at global scope
Sub BuildAndRegister(funcName, body)
    ExecuteGlobal "Function " & funcName & "(x) : " & funcName & " = " & body & " : End Function"
End Sub

Call BuildAndRegister("DynSquare", "x * x")
Call BuildAndRegister("DynCube", "x * x * x")
Call AssertEqual("Dynamic factory DynSquare", Eval("DynSquare(5)"), 25)
Call AssertEqual("Dynamic factory DynCube", Eval("DynCube(3)"), 27)

' =========================================================================
'  EVAL - Nested Eval
' =========================================================================
WScript.Echo "--- Eval: Nested ---"

Call AssertEqual("Nested Eval", Eval("Eval(""2 + 3"")"), 5)

' =========================================================================
'  SUMMARY
' =========================================================================

WScript.Echo ""
WScript.Echo "Execute/ExecuteGlobal/Eval regression test summary"
WScript.Echo "Total:  " & totalCount
WScript.Echo "Passed: " & passCount
WScript.Echo "Failed: " & failCount
If failCount = 0 Then
    WScript.Echo "Result:  ALL TESTS PASSED"
Else
    WScript.Echo "Result:  SOME TESTS FAILED"
End If
