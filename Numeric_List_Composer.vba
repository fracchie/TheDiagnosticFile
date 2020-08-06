Sub NumericListComponent()

'TODO generalise this function

    Dim size As Integer
    Dim i As Integer
    Dim Coding As String

    size = Range("CodingSize").Offset(1, 0).value

    Coding = Empty

    For i = 0 To ((2 ^ size) - 2)
        Coding = Coding + CStr(i) + " : Not Used" + vbCrLf
    Next i
    Coding = Coding + CStr(i) + " : Not Used"

    Range("CodingSize").Offset(1, 1).value = Coding
    Rows(Range("CodingSize").Offset(1, 0).Row).RowHeight = 15
    Range("CodingSize").Offset(1, 1).Select
End Sub
