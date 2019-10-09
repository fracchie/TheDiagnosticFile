Sub NumericListComponent()

    Dim size As Integer
    Dim i As Integer
    Dim coding As String

    size = Range("CodingSize").Offset(1, 0).Value

    coding = ""

    For i = 0 To ((2 ^ size) - 1)
        coding = coding + CStr(i) + " = NotUsed;" + vbCrLf
    Next i

    Range("CodingSize").Offset(1, 1).Value = coding
    Rows(Range("CodingSize").Offset(1, 0).Row).RowHeight = 15
    Range("CodingSize").Offset(1, 1).Select
End Sub
