Option Explicit

Sub ToDID()
'
' toDID Macro
' Create tab in DDT DID format from the generic parameter list
' Data in parameter list must be written accordingto the format DID_name.Data_name
' Numeric list must be written like 0 = item 1
'                                   1 = item 2 ...
' One can copy paste its whole parameters list into the tab called (for the moment) "Parameters".
'One can copy all the columns, no matter the number and differences, but the name of the used columns must be changed as specified in the parameters tab of the dfiagnostic file


    Worksheets("Parameters").Activate

    '"Departure" sheet: Parameters
    Dim HeadersRangeD As Range: Set HeadersRangeD = Range("Name", Range("Name").End(xlToRight).Address) 'the search of the header of the table is based on the top left cell which is named "Name"
    Dim NameRangeD As Range: Set NameRangeD = Range(HeadersRangeD.Find("Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim DIDRangeD As Range: Set DIDRangeD = Range(HeadersRangeD.Find("DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim LengthRangeD As Range: Set LengthRangeD = Range(HeadersRangeD.Find("Length (Byte)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Length (Byte)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim ReadRangeD As Range: Set ReadRangeD = Range(HeadersRangeD.Find("Read by DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Read by DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim WriteRangeD As Range: Set WriteRangeD = Range(HeadersRangeD.Find("Write by DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Write by DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim SnapshotRangeD As Range: Set SnapshotRangeD = Range(HeadersRangeD.Find("Snapshots", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Snapshots", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim StartRangeD As Range: Set StartRangeD = Range(HeadersRangeD.Find("Start Byte", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Start Byte", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim BitOffsetRangeD As Range: Set BitOffsetRangeD = Range(HeadersRangeD.Find("Bit Offset", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Bit Offset", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))

    Dim DID As String

    '"Arrival" sheet :toDID
    Call CreateNewTab("ToDID")

    Dim NameColA As Integer: NameColA = 1
    Dim MnemoColA As Integer: MnemoColA = 2
    Dim ParamColA As Integer: ParamColA = 3
    Dim SignColA As Integer
    Dim StartColA As Integer: StartColA = 4
    Dim OffColA As Integer: OffColA = 5
    Dim EndianColA As Integer: EndianColA = 6
    Dim RefColA As Integer: RefColA = 7
    Dim HeadersRangeA As Range:  Set HeadersRangeA = Range(Cells(1, NameColA), Cells(1, RefColA))

    Dim list As String
    Dim value As String
    Dim Label As String
    Dim A As Integer
    Dim D As Integer
    Dim Color
    Dim Sheet As Worksheet
    Dim Cell As Range

    '----------------------------------------------------------------------------------------
    'Headers
    '----------------------------------------------------------------------------------------
    A = 1

    Call formatCell(A, NameColA, "DID_name", True, 10, "BLACK", "NORMAL", "ORANGE", 40, 15)
    Call formatCell(A, MnemoColA, "Mnemo", True, 10, "BLACK", "NORMAL", "ORANGE", 11, 15)
    Columns(MnemoColA).NumberFormat = "@"
    Call formatCell(A, ParamColA, "Data_name", True, 10, "BLACK", "NORMAL", "ORANGE", 60, 15)
    Call formatCell(A, StartColA, "Size / Start Byte", True, 10, "BLACK", "NORMAL", "ORANGE", 15, 15)
    Call formatCell(A, OffColA, "Bit Offset", True, 10, "BLACK", "NORMAL", "ORANGE", 15, 15)
    Call formatCell(A, EndianColA, "Little/Big Endian", True, 10, "BLACK", "NORMAL", "ORANGE", 15, 15)
    Call formatCell(A, RefColA, "Ref", True, 10, "BLACK", "NORMAL", "ORANGE", 15, 15)


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

            DID = Replace(DID, "$", "&H")

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
