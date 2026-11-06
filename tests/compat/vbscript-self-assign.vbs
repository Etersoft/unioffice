' x = x(0): assign an element of an array to the variable that holds the array.
Option Explicit
Dim a, b, c, d
a = Array(Array(1, 2, 3))
On Error Resume Next
a = a(0)
WScript.Echo "a = a(0): Err.Number=" & Err.Number & " (0x" & Hex(Err.Number) & ") Description='" & Err.Description & "' Source='" & Err.Source & "'"
Err.Clear
On Error GoTo 0
WScript.Echo "after a = a(0): VarType=" & VarType(a) & " UBound=" & UBound(a)

b = Array(Array(1, 2, 3))
c = b(0)
WScript.Echo "c = b(0): UBound=" & UBound(c) & " c(0)=" & c(0)

d = Array(Array(1, 2, 3))

