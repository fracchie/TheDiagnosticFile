Option Explicit

Sub ToData()
    '========= Departure tab ========
    Worksheets("Parameters").Activate
    Dim HeadersRangeD As Range: Set HeadersRangeD = Range("Name", Range("Name").End(xlToRight).Address)
    Dim NameRangeD As Range: Set NameRangeD = Range(HeadersRangeD.Find("DID_Name").Address, HeadersRangeD.Find("DID_Name").End(xlDown))
    Dim MnemoRangeD As Range: Set MnemoRangeD = Range(HeadersRangeD.Find("DID").Address, HeadersRangeD.Find("DID").End(xlDown))
    Dim SizeRangeD As Range: Set SizeRangeD = Range(HeadersRangeD.Find("Size (bit)").Address, HeadersRangeD.Find("Size (bit)").End(xlDown))
    Dim NumericRangeD As Range: Set NumericRangeD = Range(HeadersRangeD.Find("Numeric").Address, HeadersRangeD.Find("Numeric").End(xlDown))
    Dim SignRangeD As Range: Set SignRangeD = Range(HeadersRangeD.Find("sign").Address, HeadersRangeD.Find("sign").End(xlDown))
    Dim UnitRangeD As Range: Set UnitRangeD = Range(HeadersRangeD.Find("unit").Address, HeadersRangeD.Find("unit").End(xlDown))
    Dim ResRangeD As Range: Set ResRangeD = Range(HeadersRangeD.Find("resolution").Address, HeadersRangeD.Find("resolution").End(xlDown))
    Dim OffsetRangeD As Range: Set OffsetRangeD = Range(HeadersRangeD.Find("Value offset").Address, HeadersRangeD.Find("Value offset").End(xlDown))
    Dim DescRangeD As Range: Set DescRangeD = Range(HeadersRangeD.Find("Description").Address, HeadersRangeD.Find("Description").End(xlDown))
    Dim ListRangeD As Range: Set ListRangeD = Range(HeadersRangeD.Find("List").Address, HeadersRangeD.Find("List").End(xlDown))
    Dim CodingRangeD As Range: Set CodingRangeD = Range(HeadersRangeD.Find("Coding", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=False).Address, HeadersRangeD.Find("Coding").End(xlDown))
    Dim ReadRangeD As Range: Set ReadRangeD = Range(HeadersRangeD.Find("Read by DID").Address, HeadersRangeD.Find("Read by DID").End(xlDown))
    Dim WriteRangeD As Range: Set WriteRangeD = Range(HeadersRangeD.Find("Write by DID").Address, HeadersRangeD.Find("Write by DID").End(xlDown))
    Dim SnapshotRangeD As Range: Set SnapshotRangeD = Range(HeadersRangeD.Find("Snapshots").Address, HeadersRangeD.Find("Snapshots").End(xlDown))
    Dim AsciiHexaRangeD As Range: Set AsciiHexaRangeD = Range(HeadersRangeD.Find("ASCII|HEXA").Address, HeadersRangeD.Find("ASCII|HEXA").End(xlDown))
    Dim DID As String
    Dim Reso As Long, CoefC As Long, off As Long
    Dim DecReso As Integer
    Dim DecOff As Integer
    '============================== Arrival sheet :toDATA
    Dim Sheet As Worksheet
    For Each Sheet In ThisWorkbook.Worksheets
        If Sheet.Name Like "ToData" Then
            Application.DisplayAlerts = False
            Worksheets("ToData").Delete
            ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = "ToData"
            Exit For
        ElseIf Sheet Is Worksheets.Item(Worksheets.Count) = True Then
            ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = "ToData"
        End If
    Next Sheet
    Worksheets("ToData").Activate
    Dim A As Integer
    A = 1
    Dim NameColA As Integer: NameColA = 1: Call formatCell(A, NameColA, "Parameter_name", True, , , "THICK", "ORANGE", 40, 20)
    Dim MnemoColA As Integer: MnemoColA = 2: Call formatCell(A, MnemoColA, "Mnemo", True, , , "THICK", "ORANGE", 10, 20)
    Dim SizeColA As Integer: SizeColA = 3: Call formatCell(A, SizeColA, "Size (bit)", True, , , "THICK", "ORANGE", 10, 20)
    Dim SignColA As Integer: SignColA = 4: Call formatCell(A, SignColA, "Sign", True, , , "THICK", "ORANGE", 10, 20)
    Dim UnitColA As Integer: UnitColA = 5: Call formatCell(A, UnitColA, "Unit", True, , , "THICK", "ORANGE", 10, 20)
    Dim CoefAColA As Integer: CoefAColA = 6: Call formatCell(A, CoefAColA, "Coef A", True, , , "THICK", "ORANGE", 10, 20)
    Dim CoefBColA As Integer: CoefBColA = 7: Call formatCell(A, CoefBColA, "Coef B", True, , , "THICK", "ORANGE", 10, 20)
    Dim CoefCColA As Integer: CoefCColA = 8: Call formatCell(A, CoefCColA, "Coef C", True, , , "THICK", "ORANGE", 10, 20)
    Dim DescColA As Integer: DescColA = 9: Call formatCell(A, DescColA, "Description", True, , , "THICK", "ORANGE", 40, 20)
    Dim NumericColA As Integer
    Dim ListColA As Integer: ListColA = 10: Call formatCell(A, ListColA, "List", True, , , "THICK", "ORANGE", 10, 20)
    Dim HeadersRangeA As Range
    Dim list() As String
    Dim value As String
    Dim Label As String
    Dim l As Integer
    Dim D As Integer
    Dim Cell As Range

    '----------------------------------------------------------------------------------------
    'Headers -> can be replaced by new function  GFL.formatCell
    '----------------------------------------------------------------------------------------
    A = 1
    Cells(A, NameColA).value = "Parameter_name"
    Cells(A, MnemoColA).value = "Mnemo"
    Cells(A, SizeColA).value = "Size (bit)"
    Cells(A, SignColA).value = "Sign"
    Cells(A, UnitColA).value = "Unit"
    Cells(A, CoefAColA).value = "Coef A"
    Cells(A, CoefBColA).value = "Coef B"
    Cells(A, CoefCColA).value = "Coef C"
    Cells(A, DescColA).value = "Description"
    Cells(A, ListColA).value = "List"
    Columns(MnemoColA).NumberFormat = "@"
    Columns.NumberFormat = "@"
    'Limit the height of rows
'    Range(Columns(StartColA), Columns(RefColA)).ColumnWidth = 14
'    'Format:interior color
    Set HeadersRangeA = Range(Cells(A, NameColA), Cells(A, ListColA))
    'for changing all the header format
    'HeadersRangeA.Interior.Color = RGB(255, 192, 0)
    'HeadersRangeA.RowHeight = 30
    'HeadersRangeA.Font.Bold = 1
    'HeadersRangeA.HorizontalAlignment = xlCenter
    'HeadersRangeA.VerticalAlignment = xlCenter
    'Columns("A:J").HorizontalAlignment = xlCenter
    ''Format:Borders
    'HeadersRangeA.borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    'HeadersRangeA.borders(xlEdgeLeft).Color = RGB(0, 0, 0)
    'HeadersRangeA.borders(xlEdgeRight).Color = RGB(0, 0, 0)
    'HeadersRangeA.borders(xlEdgeTop).Color = RGB(0, 0, 0)
    'HeadersRangeA.borders(xlInsideVertical).Color = RGB(0, 0, 0)                                        
'----------------------------------------------------------------------------------------------------
'"Arrival" sheet : BetaToDID
'----------------------------------------------------------------------------------------------------
    A = 2
    D = 2
     For Each Cell In NameRangeD.Cells 'name format DID_name.param -> structure data, or just data_name if one-param DID
        Rows(A).RowHeight = 17
        If NameRangeD.Cells(D, 1) = 0 Then 'Temporary. Find A way to iterate one time less. It counts another cell after. Was trying with for D = 2 to NameRangeD.Count (-1 in case) but still some bug. keep trying
        Else
            Cells(A, NameColA).value = NameRangeD.Cells(D, 1).value
            Debug.Print (NameRangeD.Cells(D, 1).value)
            Cells(A, MnemoColA).value = MnemoRangeD.Cells(D, 1).value
            Cells(A, SizeColA).value = SizeRangeD.Cells(D, 1).value
            Cells(A, DescColA).value = DescRangeD.Cells(D, 1).value
            If NumericRangeD.Cells(D, 1).value <> 0 Then
                If SignRangeD.Cells(D, 1).value = "s" Then
                    Cells(A, SignColA).value = 1
                Else
                    Cells(A, SignColA).value = 0
                End If
                Cells(A, UnitColA).value = UnitRangeD.Cells(D, 1).value
                Debug.Print (ResRangeD.Cells(D, 1).value)
                Debug.Print (OffsetRangeD.Cells(D, 1).value)
                Cells(A, CoefAColA).value = ResRangeD.Cells(D, 1).value
                Cells(A, CoefBColA).value = OffsetRangeD.Cells(D, 1).value
                Cells(A, CoefCColA).value = 1 'always 1 for the moment. Still have to analyse its effect

            ElseIf ListRangeD.Cells(D, 1).value <> 0 Then
                Cells(A, ListColA).value = "List"
                Cells(A, SignColA).value = 0 'There 4 assignations because otherwise DDT will not recognise it as List
                Cells(A, CoefAColA).value = 1
                Cells(A, CoefBColA).value = 0
                Cells(A, CoefCColA).value = 1
                A = A + 1
                Cells(A, MnemoColA).value = "Value"
                Cells(A, SizeColA).value = "label"
                list = Split(CodingRangeD.Cells(D, 1), vbLf)
                For l = 0 To UBound(list)
                    If (InStr(list(l), "Not Used") = 0) Then 'avoid creating values for Not Used
                        A = A + 1
                        Cells(A, MnemoColA).value = Left(list(l), InStr(list(l), ":") - 1)
                        Cells(A, SizeColA).value = Right(list(l), Len(list(l)) - InStr(list(l), ":"))
    '                    Label = Right(List(L), Len(List(L)) - InStr(List(L), "="))
    '                    Cells(A, SizeColA).Value = Label
    '                       Cells(A, MnemoColA).Value = Split(List(L), "=")
    '                       Cells(R, 2).Value = Value
    '                       Label = Right(List(L), Len(List(L)) - InStr(List(L), "="))
    '                       Cells(R, 3).Value = Label
                    End If
                Next l
            ElseIf AsciiHexaRangeD.Cells(D, 1).value <> 0 Then
                Cells(A, ListColA).value = AsciiHexaRangeD.Cells(D, 1).value
            End If
        End If
        D = D + 1
        A = A + 1
    Next Cell
    Range("A2", Cells(A - 2, ListColA)).Select
End Sub
