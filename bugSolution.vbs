Function f(a, b)
  ' Explicitly convert inputs to numeric type
  Dim numA, numB
  numA = CDbl(a)
  numB = CDbl(b)

  If numA > numB Then
    MsgBox "a is greater than b"
  ElseIf numA < numB Then
    MsgBox "a is less than b"
  Else
    MsgBox "a is equal to b"
  End If
end Function