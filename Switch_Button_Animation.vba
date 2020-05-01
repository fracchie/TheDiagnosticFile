Option Explicit

Sub Switch_OFF()

    ThisWorkbook.Activate

    'Identify button
    Dim PressedButton As String
    Dim RelatedButton As String
    PressedButton = Application.Caller

    If (Left(PressedButton, 13) = "ButtonSession") Then
'            Worksheets("Parameters").Shapes("ButtonRWSession1").Select
'            With Selection
'            End With
            Worksheets("Parameters").Shapes(PressedButton).Select

            ' Change Position, Color & Text of session button
            With Selection
                .ShapeRange.IncrementLeft -30
                .ShapeRange.TextFrame2.TextRange.Characters.text = "OFF"
                .ShapeRange.Fill.ForeColor.RGB = RGB(255, 0, 0)
            End With

            Worksheets("Parameters").Shapes(PressedButton).OnAction = "Switch_ON"
            RelatedButton = "ButtonRWSession" + Right(PressedButton, 1)

            If Worksheets("Parameters").Shapes(RelatedButton).TextFrame2.TextRange.Characters.text = "ON" Then
                    Worksheets("Parameters").Shapes(RelatedButton).IncrementLeft -30
                    Worksheets("Parameters").Shapes(RelatedButton).TextFrame2.TextRange.Characters.text = "OFF"
                    Worksheets("Parameters").Shapes(RelatedButton).Fill.ForeColor.RGB = RGB(255, 0, 0)
                    Worksheets("Parameters").Shapes(RelatedButton).OnAction = "Switch_ON"
            End If

    ElseIf (Left(PressedButton, 8) = "ButtonRW") Then
            RelatedButton = "Button" + Right(PressedButton, 8)
            If Worksheets("Parameters").Shapes(RelatedButton).TextFrame2.TextRange.Characters.text = "OFF" Then
            Else
                Worksheets("Parameters").Shapes(PressedButton).Select

                ' Change Position, Color & Text of session button
                With Selection
                    .ShapeRange.IncrementLeft -30
                    .ShapeRange.TextFrame2.TextRange.Characters.text = "OFF"
                    .ShapeRange.Fill.ForeColor.RGB = RGB(255, 0, 0)
                End With

                Worksheets("Parameters").Shapes(PressedButton).OnAction = "Switch_ON"

            End If

    Else 'Any others -> for now resetECU

        Worksheets("Parameters").Shapes(PressedButton).Select

        ' Change Position, Color & Text of session button
        With Selection
            .ShapeRange.IncrementLeft -30
            .ShapeRange.TextFrame2.TextRange.Characters.text = "OFF"
            .ShapeRange.Fill.ForeColor.RGB = RGB(255, 0, 0)
        End With

        Worksheets("Parameters").Shapes(PressedButton).OnAction = "Switch_ON"
    End If

    Worksheets("Parameters").Range("A1").Activate




End Sub

Sub Switch_ON()

    ThisWorkbook.Activate

    'Identify button
    Dim PressedButton As String
    Dim RelatedButton As String
    PressedButton = Application.Caller

    If (Left(PressedButton, 13) = "ButtonSession") Then
            Worksheets("Parameters").Shapes(PressedButton).Select

            ' Change Position, Color & Text of session button
            With Selection
                .ShapeRange.IncrementLeft 30
                .ShapeRange.TextFrame2.TextRange.Characters.text = "ON"
                .ShapeRange.Fill.ForeColor.RGB = RGB(0, 153, 0)
            End With

            Worksheets("Parameters").Shapes(PressedButton).OnAction = "Switch_OFF"

     ElseIf (Left(PressedButton, 8) = "ButtonRW") Then
            RelatedButton = "Button" + Right(PressedButton, 8)
            If Worksheets("Parameters").Shapes(RelatedButton).TextFrame2.TextRange.Characters.text = "OFF" Then
            Else
                Worksheets("Parameters").Shapes(PressedButton).Select

                ' Change Position, Color & Text of session button
                With Selection
                    .ShapeRange.IncrementLeft 30
                    .ShapeRange.TextFrame2.TextRange.Characters.text = "ON"
                    .ShapeRange.Fill.ForeColor.RGB = RGB(0, 153, 0)
                End With

                Worksheets("Parameters").Shapes(PressedButton).OnAction = "Switch_OFF"

            End If
    Else 'Any others -> for now resetECU

        Worksheets("Parameters").Shapes(PressedButton).Select

        ' Change Position, Color & Text of session button
        With Selection
            .ShapeRange.IncrementLeft 30
            .ShapeRange.TextFrame2.TextRange.Characters.text = "ON"
            .ShapeRange.Fill.ForeColor.RGB = RGB(0, 153, 0)
        End With

        Worksheets("Parameters").Shapes(PressedButton).OnAction = "Switch_OFF"

    End If

    Worksheets("Parameters").Range("A1").Activate

End Sub

'Sub Switch_ON()
'
'    ThisWorkbook.Activate
'
'    'Identify button
'    Dim PressedButton As String
'    PressedButton = Application.Caller
'
'
'    Worksheets("Parameters").Shapes(PressedButton).Select
'
'    ' Change Position, Color & Text
'    With Selection
'    .ShapeRange.IncrementLeft 30
'    .ShapeRange.TextFrame2.TextRange.Characters.text = "ON"
'    .ShapeRange.Fill.ForeColor.RGB = RGB(0, 153, 0)
'    .OnAction = "Switch_OFF"
'    End With
'
'    Worksheets("Parameters").Range("A1").Activate
'    ThisWorkbook.Activate
'
''    Worksheets("Parameters").Shapes(PressedButton).OnAction = "Switch_OFF"
''
'End Sub
