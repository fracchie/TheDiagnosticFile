'GFL = Generic Function Library

Public Sheet As Worksheet

Public Function CreateNewTab(TabName As String)
    For Each Sheet In ThisWorkbook.Worksheets
        If Sheet.Name Like TabName Then
            Application.DisplayAlerts = False
            Worksheets(TabName).Delete
            ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = TabName
            Exit For
        ElseIf Sheet Is Worksheets.Item(Worksheets.Count) = True Then
            ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = TabName
        End If
    Next Sheet

    Worksheets(TabName).Activate
End Function

Public Function CreateExtFile(fileName As String, Optional ByVal Path As String = "") As Boolean 'TODO find out how to pass the object of textstream out of this module

    Dim filePath As String
    Dim objShell As Object, objFolder As Object, objFolderItem As Object

    Set objShell = CreateObject("Shell.Application")

    If Path = "" Then
        Set objFolder = objShell.BrowseForFolder(&H0&, "Choose file's path. Consider using the folder 'input' of the DST script generator", &H1&)
        Set objFolderItem = objFolder.Items.Item
        filePath = objFolderItem.Path
    Else
        filePath = Path
    End If

    'la su peux faire un test pour savoir si l'utilisateur a mis un .xls ou non
    MsgBox filePath & "\" & fileName

    ' can put a switch "output txt file DST?" and use an IF here
    'use the info in new tab PVal to create the script DST

    Dim TempByteSent As String

    Dim MyFSO As New FileSystemObject
    If MyFSO.FolderExists(filePath) Then
        'MsgBox "The Folder already exists"
    Else
        MyFSO.CreateFolder (filePath) '<- Here the
    End If

    Dim FileOut As TextStream
    Set FileOut = MyFSO.CreateTextFile(filePath + "\" + fileName, True, True)

    CreateExtFile = True

End Function

Public Function BinToDec(bin As String, ConvMode As Integer) As Double

    Dim i As Integer
    BinToDec = 0

    'Conversion following this https://www.youtube.com/watch?v=2tzBCzW-4Qc

    For i = Len(bin) To 2 Step -1
        If Mid(bin, i, 1) = "1" Then
            BinToDec = BinToDec + 2 ^ (Len(bin) - i)

        End If

    Next i

    If (ConvMode = 2) Then '2 complement
        If (Mid(bin, 1, 1) = "1") Then
            BinToDec = BinToDec - (2 ^ (Len(bin) - 1))
        End If
    ElseIf (ConvMode = 1) Then
    'TODO 1 complement
    Else
        'ConvMode 0 = MSB sign carrier
        BinToDec = BinToDec + BinToDec + 2 ^ (Len(bin) - 1)
    End If

End Function

Public Function HexToBin(Hex As String) As String

    Dim i As Integer
    Dim bin As String
    Dim digit As String

    bin = ""

    Do While (Len(Hex) > 0)

        digit = Right(Hex, 1)
        Select Case digit
            Case "0"
                bin = "0000" + bin
            Case "1"
                bin = "0001" + bin
            Case "2"
                bin = "0010" + bin
            Case "3"
                bin = "0011" + bin
            Case "4"
                bin = "0100" + bin
            Case "5"
                bin = "0101" + bin
            Case "6"
                bin = "0110" + bin
            Case "7"
                bin = "0111" + bin
            Case "8"
                bin = "1000" + bin
            Case "9"
                bin = "1001" + bin
            Case "A"
                bin = "1010" + bin
            Case "B"
                bin = "1011" + bin
            Case "C"
                bin = "1100" + bin
            Case "D"
                bin = "1101" + bin
            Case "E"
                bin = "1110" + bin
            Case "F"
                bin = "1111" + bin

        End Select

        Hex = Left(Hex, Len(Hex) - 1)
    Loop

    HexToBin = bin

End Function

Public Function DecToBin(dec As Variant, NumBit As Integer, Optional ByVal res As Variant = 1, Optional ByVal off As Double = 0) As String
' converts a decimal number in a binary value in n bit
'think about the resolution
    Dim i As Integer
    If dec <> 0 Then
'        Debug.Print ("res")
'        Debug.Print (res)
        'TODO bug fixing. Temporarly:
        If res = 0 Then
            res = 1
        End If
        dec = dec / res
    End If
    dec = dec - off
    For i = NumBit - 1 To 0 Step -1 'countdown
        If Int(dec / (2 ^ i)) > 0 Then
            DecToBin = DecToBin + "1"
            dec = dec - (2 ^ i)
        Else
            DecToBin = DecToBin + "0"
        End If
    Next i
End Function

Public Function BinToHex(Binary As String) As String
    Dim value&, i&, Base#: Base = 1
    Dim l As Integer, j As Integer
    Dim text As String
    Dim substring As String
    Dim original As String
    Dim hexadecimal As String

    hexadecimal = ""
    original = Binary
    Do While Len(original) \ 8 >= 1
        substring = Left(original, 8)
        'Convert substring (byte)
        l = Len(substring)
        value = 0
        Base = 1

        For i = Len(substring) To 1 Step -1
            value = value + IIf(Mid(substring, i, 1) = "1", Base, 0) ' -> no bit limit anymore
            Base = Base * 2
        Next i
        text = Hex(value) 'Hex converts from decimal to hex

        l = l - (Len(text) * 4)

        Do While l > 0
            l = l - 4
            text = "0" + text
        Loop

        hexadecimal = hexadecimal + text
        original = Right(original, Len(original) - 8)
    Loop

    BinToHex = hexadecimal

End Function

Public Function CanoeReadDTC(Optional ByVal DTC As String = "", Optional ByVal FaultType As String = "", Optional ByVal Status As String = "") As String
    ' supposing existance of CANoe function of the type readDTC([DTC], [FaultType]) that if called without arguments reads them all, otherwise searches for the specific DTC
    Dim temp As String

    If (DTC = "") Then
        temp = "readDTC();"
    Else
        If (FaultType = "") Then
            temp = "readDTC(" + DTC
        Else
            temp = "readDTC(" + DTC + ", " + FaultType
        End If
    End If
    If Status <> "" Then
        temp = temp + "," + Status
    End If
    temp = temp + ");"
    'TODO check the status present or memorised.
    CanoeReadDTC = temp
End Function

Public Function CanoeWriteDID(Name As String, hexa As String) As String
    'Find out how to write a DID on CANoe
    'Ouput the line to be written in FileOut to write the DID
End Function

Public Function CanoeReadDID(Name As String) As String
    'Find out how to write a DID on CANoe
    'Ouput the line to be written in FileOut to read the DID
End Function

Public Function CanoeRestoreFrame(Channel As String, ECU As String, Frame As String) As String 'TODO add optional put back signal after x ms
    Dim temp As String
    temp = "@sysvar::" + Channel + "::" + ECU + "::" + Frame + "::TIMINGS::EnableCyclic=1"
    CanoeRestoreFrame = temp
End Function

Public Function CanoeCutFrame(Channel As String, ECU As String, Frame As String) As String 'TODO add optional put back signal after x ms
    Dim temp As String
    temp = "@sysvar::" + Channel + "::" + ECU + "::" + Frame + "::TIMINGS::EnableCyclic=0"
    CanoeCutFrame = temp
    'TODO the time management, but for the moment it is not designed like that. Call restoreFrame instead
End Function


Public Function CanoeReadSignalValue(signal As String, Optional ByVal ExpectedValue As String = "", Optional ByVal expResult As Boolean = True) As String
    Dim temp As String

    temp = "readSignal($" + signal
    If ExpectedValue <> "" Then
        temp = temp + "," + ExpectedValue
    End If
    temp = temp + "," + CStr(expResult)
    'difficult because the otuput of this function should be multiple lines of CAPL code.
    ' soulution would be to transmit also the reference to the TextStream to output directly from here... TODO
    temp = temp + ");"
    CanoeReadSignalValue = temp
End Function

Public Function CanoeWriteSignalValue(signal As String, value As String) As String
    Dim temp As String
    temp = "writeSignal($" + signal + "," + value + ");"
    CanoeWriteSignalValue = temp
End Function


Public Function ParameterMask(ContentBin As String, ByteStart As Integer, bitOffset As Integer, DIDLenght As Integer) As String
'want to write a parameter, without touching the other in the same DID
'i would then write all 1 in the other bits of the DID, and then write the content of the interested parm in binary
'and finally to an AND with the original content of the DID
    Dim x As Integer
    Dim y As Integer

    Dim output As String

    output = ""
    y = 0
    x = ByteStart * 8 + bitOffset 'BitOffset 0-7

    Do While y < x
        output = output + "1"
        y = y + 1
    Loop

    output = output + ContentBin
    y = y + Len(ContentBin)

    x = DIDLenght * 8
    Do While y < x
        output = output + "1"
        y = y + 1
    Loop

    ParameterMask = output

End Function

Public Function formatCell(r As Integer, c As Integer, text As String, Optional ByVal Bold As Boolean = False, Optional ByVal FontSize As Double = 10, Optional ByVal FontColor As String = "Black", Optional ByVal borders As String, Optional ByVal FillColor As String, Optional ByVal width As Double = 9, Optional ByVal height As Double = 15)
    Cells(r, c).Select
    With Selection
        .value = text
        If (Bold) Then
            .Font.Bold = True
        End If
        Select Case FontColor
            Case "DarkBlue"
                .Font.Color = RGB(48, 84, 150)
            Case "White"
                .Font.Color = RGB(255, 255, 255)
            Case Else
                .Font.Color = RGB(0, 0, 0)
        End Select

        .Font.size = FontSize
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .WrapText = True
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        Select Case FillColor
            Case "ORANGE"
                .Interior.Color = RGB(255, 192, 0)
            Case "PURPLE"
                .Interior.Color = RGB(153, 153, 255)
            Case "LightBlue"
                .Interior.Color = RGB(155, 194, 230)
            Case "DarkOrange"
                .Interior.Color = RGB(237, 125, 49)
            Case "Blue"
                .Interior.Color = RGB(0, 176, 240)
            Case "DarkBlue"
                .Interior.Color = RGB(48, 84, 150)
            Case Else
                '.Interior.Color = RGB(255, 192, 0) 'nothing, empty
        End Select

        Select Case borders
            Case "THICK"
                .borders.Weight = xlThick
                .borders(xlEdgeBottom).Color = RGB(0, 0, 0)
                .borders(xlEdgeLeft).Color = RGB(0, 0, 0)
                .borders(xlEdgeRight).Color = RGB(0, 0, 0)
                .borders(xlEdgeTop).Color = RGB(0, 0, 0)
            Case "BottomThick"
                .borders.Weight = xlMedium
                .borders(xlEdgeLeft).Color = RGB(0, 0, 0)
                .borders(xlEdgeRight).Color = RGB(0, 0, 0)
                .borders(xlEdgeTop).Color = RGB(0, 0, 0)
                .borders.Weight = xlThick
                .borders(xlEdgeBottom).Color = RGB(0, 0, 0)
            Case Else 'Case "NORMAL"
                .borders.Weight = xlMedium
                .borders(xlEdgeBottom).Color = RGB(0, 0, 0)
                .borders(xlEdgeLeft).Color = RGB(0, 0, 0)
                .borders(xlEdgeRight).Color = RGB(0, 0, 0)
                .borders(xlEdgeTop).Color = RGB(0, 0, 0)
            End Select
    End With

    Rows(r).RowHeight = height
    Columns(c).ColumnWidth = width
End Function


Public Function CollapseColumnsRight()

ActiveSheet.Outline.SummaryColumn = xlRight

End Function

Public Function CollapseColumnsLeft()

ActiveSheet.Outline.SummaryColumn = xlLeft

End Function

Public Function CollapseRowsAbove()

ActiveSheet.Outline.SummaryRow = xlAbove

End Function

Public Function CollapseRowsBelow()

ActiveSheet.Outline.SummaryRow = xlBelow

End Function

Public Function Expand_All()
    ActiveSheet.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
End Function

Public Function Collapse_All()
    ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
End Function

Public Function CanoeDelay(time As String) As String
    Dim temp As String
    temp = "Delay(" + time + ");"
    CanoeDelay = temp
End Function

Public Function replaceInString(original As String, replacement As String, startPos0based As Integer) As String
'TODO if replacing would make the original word bigger, manage it. for the moment i don't care
    Dim i As Integer
    Dim out As String
    out = ""
    Dim digit As String
    i = 0
    Do While i < startPos0based
        digit = Left(original, i + 1)
        digit = Right(digit, 1)
        out = out + digit
        i = i + 1
    Loop

    out = out + replacement
    i = i + Len(replacement)

    Do While i < Len(original)
        digit = Left(original, i + 1)
        digit = Right(digit, 1)
        out = out + digit
        i = i + 1
    Loop

    replaceInString = out

End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Integer
  Dim i As Long
  Dim found As Boolean
  ' default return value if value not found in array
  IsInArray = -1
  found = False

    Do While (found = False)
        Debug.Print ("found!")
        found = True
    Loop

  For i = LBound(arr) To UBound(arr)
    If InStr(stringToBeFound, arr(i), 1) = 0 Then
      IsInArray = i
      Exit For
    End If
  Next i
End Function
