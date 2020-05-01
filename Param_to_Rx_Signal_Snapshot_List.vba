Public Sheet As Worksheet

Option Explicit

'One can capy the interesting signals from the messageSet classic format and paste them in this file
'select what needs to be checked - unavailable value for the moment, but can be others will be selectable next
'and generate a txt file in the canoe format
'so that it can be used as a caoe script to be run and check in automatic the value of all the selected signals

'apparently for the moment i just need to take SignalName and ExpectedValue
' the CANoe format is like that:
    'if ($SpotHalosFrozenWindshield== 0) {
    '       TestStepPass("", "SpotHalosFrozenWindshield = 0");
    '    } else {
    '        TestStepFail("", "SpotHalosFrozenWindshield = %f EXPECTED: 0", $SpotHalosFrozenWindshield);
    '    }



'Clean immediate windows
'Application.SendKeys "^g ^a {DEL}"

'============================================================================================================================================================================================================================================================================================
'      Global Variables Declaration
'============================================================================================================================================================================================================================================================================================


'Public CodingRangeD As Range 'note that the numeric value and the meaning are in two different columns in the MSRS format...

Sub SignalSnapshotListCreation()

'============================================================================================================================================================================================================================================================================================
'           Setup
'============================================================================================================================================================================================================================================================================================

    'Workbooks("TestUnavailable").Activate 'use that when debugging with several workbooks open
    Worksheets("Signals").Activate
    '----------------------------------------------------------------------------------------------------
    'Variables declaration and init
    '----------------------------------------------------------------------------------------------------

    Dim A As Integer
    Dim D As Integer
    Dim HeadersRangeD As Range
    Dim SignalNameRangeD As Range
    Dim FrameNameRangeD As Range
    Dim UnavailableValueRangeD As Range
    Dim MinValueRangeD As Range
    Dim MaxValueRangeD As Range
    Dim ResolutionRangeD As Range
    Dim SizeRangeD As Range
    Dim SignRangeD As Range
    Dim OffsetRangeD As Range
    Dim CodingRangeD As Range
    Dim temp As String
    Dim ExpectedValueRangeD As Range

    Dim ExpectedValueDec As String
    Dim ExpectedValueHex As String

    'Find the needed columns in the header list. By default is NOT CASE SENSITIVE
    'Using fixed header A1 cell called SignalName. In case of modification, this needs to be modify accordingly
    Set HeadersRangeD = Range("SignalName", Range("SignalName").End(xlToRight).Address)
    Set SignalNameRangeD = Range(HeadersRangeD.Find("Signal Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Signal Name").End(xlDown))
    Set FrameNameRangeD = Range(HeadersRangeD.Find("Frame Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Frame Name").End(xlDown))
    Set UnavailableValueRangeD = Range(HeadersRangeD.Find("Unavailable Value (Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Unavailable Value (Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set MinValueRangeD = Range(HeadersRangeD.Find("Min (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Min (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set MaxValueRangeD = Range(HeadersRangeD.Find("Max (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Max (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set ResolutionRangeD = Range(HeadersRangeD.Find("Resolution (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Resolution (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set SizeRangeD = Range(HeadersRangeD.Find("Signal Size (Bits)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Signal Size (Bits)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set SignRangeD = Range(HeadersRangeD.Find("Value Type (Sign)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Value Type (Sign)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set OffsetRangeD = Range(HeadersRangeD.Find("Offset (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Offset (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set CodingRangeD = Range(HeadersRangeD.Find("Coding (Bin/Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Coding (Bin/Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set ExpectedValueRangeD = Range(HeadersRangeD.Find("Expected Value", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Expected Value", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))

    Dim TempByteSent As String

    For Each Sheet In ThisWorkbook.Worksheets
        If Sheet.Name Like "ParamListSignals" Then
            Application.DisplayAlerts = False 'used to mute the message to confirm tab deletion
            Worksheets("ParamListSignals").Delete
            ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = "ParamListSignals"
            Exit For
        ElseIf Sheet Is Worksheets.Item(Worksheets.Count) = True Then
            ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = "ParamListSignals"
        End If
    Next Sheet

    Cells(1, 1).value = "Data Name"

    A = 2

'---- Start the loop
    For D = 2 To SignalNameRangeD.Cells.Count

        Cells(A, 1).value = FrameNameRangeD.Cells(D, 1).value + "_" + SignalNameRangeD.Cells(D, 1).value
        A = A + 1

    Next D

End Sub
