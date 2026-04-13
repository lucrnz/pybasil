'# Expected output (all lines must appear, in order):
'#   test1: 0
'#   test2: 42
'#     got: hello
'#   test3: after PrintVal
'#     done
'#   test4: both ran
'#   test5: i=1
'#   test5: i=2

' Test 1: Method call without args followed by another statement
On Error Resume Next
x = 1 / 0
Err.Clear
WScript.Echo "test1: " & Err.Number

' Test 2: User-defined Sub without args followed by another statement
Sub ResetState
End Sub

y = 42
ResetState
WScript.Echo "test2: " & y

' Test 3: Sub with args (no parens) -- next line must NOT be consumed
Sub PrintVal(v)
    WScript.Echo "  got: " & v
End Sub

PrintVal "hello"
WScript.Echo "test3: after PrintVal"

' Test 4: Two consecutive implicit calls
Sub DoNothing
End Sub

Sub SayDone
    WScript.Echo "  done"
End Sub

DoNothing
SayDone
WScript.Echo "test4: both ran"

' Test 5: Implicit call inside a loop body
Sub Tick
End Sub

For i = 1 To 2
    Tick
    WScript.Echo "test5: i=" & i
Next
