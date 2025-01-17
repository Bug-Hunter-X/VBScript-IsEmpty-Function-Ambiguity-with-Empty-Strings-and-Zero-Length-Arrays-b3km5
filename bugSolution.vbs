Function MyFunction(param1)
  If IsNull(param1) Or (TypeName(param1) = "String" And Len(param1) = 0) Or (TypeName(param1) = "Variant()" And UBound(param1, 1) = -1) Then
    ' Handle empty parameter, null value, zero-length string, or zero-length array
    Exit Function 
  End If
  ' ... rest of the function
End Function