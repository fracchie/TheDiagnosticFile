Sub NumericListComponent()

'TODO generalise this function

    Dim size As Integer
    Dim i As Integer
    Dim coding As String

    size = Range("CodingSize").Offset(1, 0).value

    coding = Empty

    For i = 0 To ((2 ^ size) - 1)
        coding = coding + CStr(i) + " : Not Used" + vbCrLf
    Next i

    Range("CodingSize").Offset(1, 1).value = coding
    Rows(Range("CodingSize").Offset(1, 0).row).RowHeight = 15
    Range("CodingSize").Offset(1, 1).Select
End Sub
