Sub ToggleButton()

    Dim PressedButton As String
    PressedButton = Application.Caller
    ActiveSheet.Shapes(PressedButton).Select
    ' Change Position, Color & Text of session button
    With Selection
        If .ShapeRange.Fill.ForeColor.RGB = RGB(0, 255, 0) Then
            .ShapeRange.Fill.ForeColor.RGB = RGB(255, 0, 0)
        Else
            .ShapeRange.Fill.ForeColor.RGB = RGB(0, 255, 0)
        End If

    End With

    Cells(1, 1).Select

End Sub
