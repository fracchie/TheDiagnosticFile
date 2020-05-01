Option Explicit

Sub ToDID()
'
' toDID Macro
' Create tab in DDT DID format from the generic parameter list
' Data in parameter list must be written accordingto the format DID_name.Data_name
' Numeric list must be written like 0 = item 1
'                                   1 = item 2 ...
' One can copy paste its whole parameters list into the tab called (for the moment) "Parameters".
'One can copy all the columns, no matter the number and differences, but the name of the used columns must be changed as specified here below:
'Name
'DID
'Size (Bit)
'Write
'Read
'Snapshot
'Numeric
'Signed unsigned
'unit
'res
'Min
'Max
'Offset
'Start Byte
'Bit Offset
'Length (Byte)
'List
'Coding


'----------------------------------------------------------------------------------------------------
'Variables declaration
'----------------------------------------------------------------------------------------------------
    '"Departure" sheet
    Dim HeadersRangeD As Range
    Dim NameRangeD As Range
    Dim DIDRangeD As Range
    Dim LengthRangeD As Range
    Dim ReadRangeD As Range
    Dim WriteRangeD As Range
    Dim SnapshotRangeD As Range
    Dim StartRangeD As Range
    Dim BitOffsetRangeD As Range
    Dim DID As String
    '"Arrival" sheet :toDATA
    Dim NameColA As Integer, MnemoColA As Integer, ParamColA As Integer, SignColA As Integer, StartColA As Integer, OffColA As Integer, EndianColA As Integer, RefColA As Integer
    Dim HeadersRangeA As Range
    Dim list As String, value As String, Label As String
    Dim A As Integer, D As Integer
    Dim Color
    Dim Sheet As Worksheet
    Dim Cell As Range


    Worksheets("Parameters").Activate
'----------------------------------------------------------------------------------------------------
''look for the specific headers defining a DDT data, stored in the ListHeaders.
' find where headers are written, and define the headers row
    'Range("HeaderRowD", Range("HeaderRowD").End(xlToRight).Address).Select
    'Range("HeadersRangeD").EntireRow.

    'Set HeadersRangeD = Range("HeadersRangeD", Range("HeadersRangeD").End(xlToRight).Address)
    Set HeadersRangeD = Range("Name", Range("Name").End(xlToRight).Address)
    HeadersRangeD.Select
    'would like to format the whole thing as a tab, and maybe formatting the headers as text
    'Find the needed columns in the header list. By default is NOT CASE SENSITIVE
    'HeadersRangeD.Find("Name").Select
    Set NameRangeD = Range(HeadersRangeD.Find("Name").Address, HeadersRangeD.Find("Name").End(xlDown))

    Set DIDRangeD = Range(HeadersRangeD.Find("DID").Address, HeadersRangeD.Find("DID").End(xlDown))
    Set LengthRangeD = Range(HeadersRangeD.Find("Length (Byte)").Address, HeadersRangeD.Find("Length (Byte)").End(xlDown))
    Set StartRangeD = Range(HeadersRangeD.Find("Start Byte").Address, HeadersRangeD.Find("Start Byte").End(xlDown))
    Set BitOffsetRangeD = Range(HeadersRangeD.Find("Bit Offset").Address, HeadersRangeD.Find("Bit Offset").End(xlDown))
    Set ReadRangeD = Range(HeadersRangeD.Find("Read by DID").Address, HeadersRangeD.Find("Read by DID").End(xlDown))
    Set WriteRangeD = Range(HeadersRangeD.Find("Write by DID").Address, HeadersRangeD.Find("Write by DID").End(xlDown))
    Set SnapshotRangeD = Range(HeadersRangeD.Find("Snapshots").Address, HeadersRangeD.Find("Snapshots").End(xlDown))


    WriteRangeD.Select



'----------------------------------------------------------------------------------------------------
'"Arrival" sheet : BetaToDID
'----------------------------------------------------------------------------------------------------

    For Each Sheet In ThisWorkbook.Worksheets
        If Sheet.Name Like "ToDID" Then
            Application.DisplayAlerts = False
            Worksheets("ToDID").Delete
            ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = "ToDID"
            Exit For
        ElseIf Sheet Is Worksheets.Item(Worksheets.Count) = True Then
            ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = "ToDID"
        End If
    Next Sheet
    'Work on the "toDATA" sheet
    Worksheets("ToDID").Activate

    'Define  the columns containing the data, by default
    NameColA = 1
    MnemoColA = 2
    ParamColA = 3
    StartColA = 4
    OffColA = 5
    EndianColA = 6
    RefColA = 7

    '----------------------------------------------------------------------------------------
    'Headers
    '----------------------------------------------------------------------------------------
    A = 1
    Cells(A, NameColA).value = "DID_name"
    Cells(A, MnemoColA).value = "Mnemo"
    Cells(A, ParamColA).value = "Data_name"
    Cells(A, StartColA).value = "Size / Start Byte"
    Cells(A, OffColA).value = "Bit Offset"
    Cells(A, EndianColA).value = "Little/Big Endian"
    Cells(A, RefColA).value = "Ref"

    'Format
    'Format:Columns width
    Columns(NameColA).ColumnWidth = 40
    Columns(MnemoColA).ColumnWidth = 11
    Columns(MnemoColA).NumberFormat = "@"
    Columns(ParamColA).ColumnWidth = 60
    Range(Columns(StartColA), Columns(RefColA)).ColumnWidth = 14
    'Format:interior color
    Set HeadersRangeA = Range(Cells(A, NameColA), Cells(A, RefColA))
    HeadersRangeA.Interior.Color = RGB(255, 192, 0)
    HeadersRangeA.RowHeight = 30
    HeadersRangeA.Font.Bold = 1
    HeadersRangeA.HorizontalAlignment = xlCenter
    HeadersRangeA.VerticalAlignment = xlCenter


    'Format:Borders
    HeadersRangeA.borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    HeadersRangeA.borders(xlEdgeLeft).Color = RGB(0, 0, 0)
    HeadersRangeA.borders(xlEdgeRight).Color = RGB(0, 0, 0)
    HeadersRangeA.borders(xlEdgeTop).Color = RGB(0, 0, 0)
    HeadersRangeA.borders(xlInsideVertical).Color = RGB(0, 0, 0)


    A = 2
    D = 2

    For Each Cell In NameRangeD.Cells 'name format DID_name.param -> structure data, or just data_name if one-param DID

        If DIDRangeD.Cells(D, 1) = 0 Then   'Temporary. Find a way to iterate one time less
            'for avoiding creating
        'if new DID
        ElseIf DIDRangeD.Cells(D, 1) <> DIDRangeD.Cells(D - 1, 1) Then 'if first voice of new DID


            If Not (TypeName(DIDRangeD.Cells(D, 1).value) Like "String") Then
                    DID = Str(DIDRangeD.Cells(D, 1).value)
            Else
                DID = DIDRangeD.Cells(D, 1).value
            End If

'            DID = Replace(DID, "0113", " ")
'            DID = Replace(DID, " ", "")
'            Debug.Print (DID)
'            Debug.Print (InStr(DID, "$"))
            DID = Replace(DID, "$", "&H")
'            Debug.Print (CDec(DID))

            Cells(A, MnemoColA).value = CDec(DID)

            If SnapshotRangeD.Cells(D, 1).value = "X" Then
                Cells(A, ParamColA).value = 2
            ElseIf WriteRangeD.Cells(D, 1).value <> 0 Then
                Cells(A, ParamColA).value = 4
            Else
                Cells(A, ParamColA).value = 3
            End If

            Cells(A, StartColA).value = LengthRangeD.Cells(D, 1).value
            Cells(A, EndianColA).value = 0
            Cells(A, RefColA).value = 0
            Cells(A, OffColA).value = 0

           'if is a DID containing several data -> create first line DID in Arrival sheet
            If InStr(NameRangeD.Cells(D, 1).value, ".") <> 0 Then

'                 Cells(A, NameColA).Value = Left(NameRangeD.Cells(D, 1).Value, InStr(NameRangeD.Cells(D, 1).Value, ".") - 1)
                    DID = Left(NameRangeD.Cells(D, 1).value, InStr(NameRangeD.Cells(D, 1).value, ".") - 1)
            'otherwise it is a DID with just one data
            Else
                Cells(A, NameColA).value = NameRangeD.Cells(D, 1).value
                DID = NameRangeD.Cells(D, 1).value

            End If

            Debug.Print ("------ " + DID + " --------")
            Cells(A, NameColA).value = DID

            A = A + 1
            Cells(A, MnemoColA).value = "record"
'            Cells(A, ParamColA).Value = Right(NameRangeD.Cells(D, 1).Value, Len(NameRangeD.Cells(D, 1).Value) - InStr(NameRangeD.Cells(D, 1).Value, "."))
            Cells(A, ParamColA).value = NameRangeD.Cells(D, 1).value
            Cells(A, StartColA).value = StartRangeD.Cells(D, 1).value + 3
            Cells(A, OffColA).value = BitOffsetRangeD.Cells(D, 1).value
            Cells(A, EndianColA).value = 0
            Cells(A, RefColA).value = 1


        'or if nth record of same DID
        Else
            Cells(A, MnemoColA).value = "record"
'            Cells(A, ParamColA).Value = Right(NameRangeD.Cells(D, 1).Value, Len(NameRangeD.Cells(D, 1).Value) - InStr(NameRangeD.Cells(D, 1).Value, "."))
            Cells(A, ParamColA).value = NameRangeD.Cells(D, 1).value
            Cells(A, StartColA).value = StartRangeD.Cells(D, 1).value + 3 'there is an offset of -3 in
            Cells(A, OffColA).value = BitOffsetRangeD.Cells(D, 1).value
            Cells(A, EndianColA).value = 0
            Cells(A, RefColA).value = 1

        End If

        D = D + 1
        A = A + 1
    Next Cell

    Range("A2", Cells(A - 2, RefColA)).Select

End Sub
