Option Explicit
Sub Sort()
'
' Sort Macro
'

'
    ActiveWorkbook.Worksheets("Parameters").Sort.SortFields.Clear

        'Set HeadersRangeD = Range("HeadersRangeD", Range("HeadersRangeD").End(xlToRight).Address)

    Dim HeadersRangeD As Range
    Set HeadersRangeD = Range("Name", Range("Name").End(xlToRight).Address)
    HeadersRangeD.Select
    'would like to format the whole thing as a tab, and maybe formatting the headers as text
    'Find the needed columns in the header list. By default is NOT CASE SENSITIVE

    Dim NameRangeD As Range
    Dim DIDRangeD As Range
    Dim StartRangeD As Range
    Dim BitOffsetRangeD As Range
    Dim ParametersTableD As Range

    Set NameRangeD = Range(HeadersRangeD.Find("Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set DIDRangeD = Range(HeadersRangeD.Find("DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set StartRangeD = Range(HeadersRangeD.Find("Start Byte", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Start Byte", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set BitOffsetRangeD = Range(HeadersRangeD.Find("Bit offset", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Bit offset", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set ParametersTableD = Range(NameRangeD.Cells(1, 1), "BB1000") 'TODO select exactly the table

    ActiveWorkbook.Worksheets("Parameters").Sort.SortFields.Add Key:=DIDRangeD, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Parameters").Sort.SortFields.Add Key:=StartRangeD, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Parameters").Sort.SortFields.Add Key:=BitOffsetRangeD, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Parameters").Sort
        .SetRange Range("B12", "Y24") 'TODO need to select the whole table range
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
