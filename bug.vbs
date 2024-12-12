Function MyFunction(param1)
  If IsEmpty(param1) Then
    Err.Raise 13, , "Parameter cannot be empty"
  End If
End Function