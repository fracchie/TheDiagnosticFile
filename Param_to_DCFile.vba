Option Explicit

' Data need to be sorted by DID first, then by Start byte and then for Bit offset. all smaller first
' The whole parameter tab needs to be sorted before launching the Macro

'============================================================================================================================================================================================================================================================================================
'      Global Variables Declaration
'============================================================================================================================================================================================================================================================================================

Public A As Integer
Public D As Integer
Public HeadersRangeD As Range
Public NameRangeD As Range
Public DIDRangeD As Range
Public LengthRangeD As Range
Public DescriptionRangeD As Range
Public SizeRangeD As Range
Public WriteRangeD As Range
Public ReadRangeD As Range
Public SnapRangeD As Range
Public StartRangeD As Range
Public BitOffsetRangeD As Range
Public DefaultRangeD As Range
Public NumericRangeD As Range
Public ListRangeD As Range
Public MinRangeD As Range
Public MaxRangeD As Range
Public ResRangeD As Range
Public CodingRangeD As Range
Public SignRangeD As Range
Public OffsetRangeD As Range
Public ConfigRangeD As Range
Public DID As String

Public DataNameColA As Integer
Public SizeColA As Integer
Public DescriptionColA As Integer
Public MTCColA As Integer
Public ValueColA As Integer
Public CommentsColA As Integer

Public Color
Public Sheet As Worksheet
Public Cell As Range
Public i As Integer
Public Bit As Integer


Sub DCgen()

  'Set HeadersRangeD = Range("HeadersRangeD", Range("HeadersRangeD").End(xlToRight).Address)
    'Set HeadersRangeD = Range("Headers", Range("Headers").End(xlToRight).Address)
    Worksheets("Parameters").Activate

    'Set HeadersRangeD = Range("HeadersRangeD", Range("HeadersRangeD").End(xlToRight).Address)
    Set HeadersRangeD = Range("Name", Range("Name").End(xlToRight).Address)
    HeadersRangeD.Select
    'would like to format the whole thing as a tab, and maybe formatting the headers as text
    'Find the needed columns in the header list. By default is NOT CASE SENSITIVE
    Set NameRangeD = Range(HeadersRangeD.Find("Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set DescriptionRangeD = Range(HeadersRangeD.Find("Description", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Description", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    'Set DIDRangeD = Range(HeadersRangeD.Find("DID").Address, HeadersRangeD.Find("DID").End(xlDown))
    'Set LengthRangeD = Range(HeadersRangeD.Find("Length").Address, HeadersRangeD.Find("Length").End(xlDown))
    'Set WriteRangeD = Range(HeadersRangeD.Find("Write").Address, HeadersRangeD.Find("Write").End(xlDown))
    'Set ReadRangeD = Range(HeadersRangeD.Find("Read").Address, HeadersRangeD.Find("Read").End(xlDown))
    Set SizeRangeD = Range(HeadersRangeD.Find("Size (bit)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Size (bit)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set DefaultRangeD = Range(HeadersRangeD.Find("Default Value", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Default Value", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set NumericRangeD = Range(HeadersRangeD.Find("Numeric", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Numeric", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set MinRangeD = Range(HeadersRangeD.Find("min", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("min", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set MaxRangeD = Range(HeadersRangeD.Find("max", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("max", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set ResRangeD = Range(HeadersRangeD.Find("resolution", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("resolution", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    'Set SignRangeD = Range(HeadersRangeD.Find("sign").Address, HeadersRangeD.Find("sign").End(xlDown))
    Set OffsetRangeD = Range(HeadersRangeD.Find("Value offset", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Value offset", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set ListRangeD = Range(HeadersRangeD.Find("List", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("List", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    'Set StartRangeD = Range(HeadersRangeD.Find("Start Byte").Address, HeadersRangeD.Find("Start Byte").End(xlDown))
    'Set BitOffsetRangeD = Range(HeadersRangeD.Find("Bit offset").Address, HeadersRangeD.Find("Bit offset").End(xlDown))
    Set CodingRangeD = Range(HeadersRangeD.Find("Coding", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Coding", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set ConfigRangeD = Range(HeadersRangeD.Find("Config", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Config", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))

    Dim A As Integer
    Dim list() As String, value As String, Label As String, l As Integer

    For Each Sheet In ThisWorkbook.Worksheets
        If Sheet.Name Like "DCfile" Then
            Application.DisplayAlerts = False
            Worksheets("DCfile").Delete
            ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = "DCfile"
            Exit For
        ElseIf Sheet Is Worksheets.Item(Worksheets.Count) = True Then
            ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = "DCfile"
        End If
    Next Sheet

    'Work on the "toDATA" sheet
    Worksheets("DCfile").Activate

    'Define  the columns containing the data, by default
    DataNameColA = 1
    SizeColA = 2
    DescriptionColA = 3
    MTCColA = 4
    ValueColA = 5
    CommentsColA = 6

    Columns("A").ColumnWidth = 60
    Columns("B").ColumnWidth = 7
    Columns("C").ColumnWidth = 40
    Columns("D").ColumnWidth = 25
    Columns("F").ColumnWidth = 50

    Range("A1", "L1000").HorizontalAlignment = xlLeft
    Range("A1", "L1000").VerticalAlignment = xlCenter
    Range("A1", "L1000").WrapText = True
    Range("A1", "L1000").RowHeight = 15

    Worksheets("DC Format").Activate
    Range("A1", "I45").Copy
    Worksheets("DCfile").Activate
    Range("A1").PasteSpecial
    Rows(1).RowHeight = 25
    Rows(26).RowHeight = 25
    Rows(28).RowHeight = 25
    Rows(36).RowHeight = 25


    A = 46 'starting Arrival sheet line should be 55 if considering the header of DC, TO ADD

    For D = 2 To NameRangeD.Cells.Count
        If ConfigRangeD.Cells(D, 1) <> 0 Then 'If this data is meant to be in config
            Debug.Print (NameRangeD.Cells(D, 1).value)
            Cells(A, DataNameColA).value = NameRangeD.Cells(D, 1).value
            Cells(A, SizeColA).value = SizeRangeD.Cells(D, 1).value
            'Cells(A, DescriptionColA).Value = DescriptionRangeD.Cells(D, 1).Value 'TODO removed because tradconf bugs with description if things like " are used
            If NumericRangeD.Cells(D, 1) <> 0 Then
                Cells(A, ValueColA).value = Left(DefaultRangeD.Cells(D, 1), Len(DefaultRangeD.Cells(D, 1)) - InStr(DefaultRangeD.Cells(D, 1), " ") - 1)
            ElseIf ListRangeD.Cells(D, 1) <> 0 Then

                list = Split(CodingRangeD.Cells(D, 1), vbLf)
                l = 0
                Cells(A, ValueColA).value = Left(list(l), InStr(list(l), "=") - 1)
                Cells(A, CommentsColA).value = Right(list(l), Len(list(l)) - InStr(list(l), "="))
                For l = 1 To UBound(list)
                    A = A + 1
                    Cells(A, ValueColA).value = Left(list(l), InStr(list(l), "=") - 1)
                    Cells(A, CommentsColA).value = Right(list(l), Len(list(l)) - InStr(list(l), "="))
                Next l

            End If

            A = A + 1
        End If

    Next D

    Range("A46").Activate
End Sub
