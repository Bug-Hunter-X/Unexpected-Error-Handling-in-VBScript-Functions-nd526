Function MyFunction(param1)
  On Error Resume Next
  If IsEmpty(param1) Then
    Err.Raise 13, , "Parameter cannot be empty"
  End If
  On Error GoTo 0
  ' Function logic here
  MyFunction = param1 * 2
End Function

'Example Caller:
On Error Resume Next
Dim result
result = MyFunction(10)
If Err.Number <> 0 Then
  MsgBox "Error: " & Err.Description
Else
  MsgBox "Result: " & result
End If

result = MyFunction("")
If Err.Number <> 0 Then
  MsgBox "Error: " & Err.Description
Else
  MsgBox "Result: " & result
End If
On Error GoTo 0