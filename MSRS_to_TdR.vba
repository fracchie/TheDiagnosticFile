Option Explicit

'One can copy the interesting signals from the messageSet classic format and paste them in this file
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


Public FrameName As String
Public FrameID As String
Public SignalName As String
Public FramePeriod As Integer
Public Sheet As Worksheet

'Arrival sheet layout -> TODO it can be done kind of automatically giving an array with the name of the columns
Const FirstColumnA As Integer = 1 'Easiest way?
Const ChannelNameA As Integer = 1
Const ECUNameA As Integer = 2
Const FrameNameA As Integer = 3
Const IPDUNameA As Integer = 4
Const SignalNameA As Integer = 5
Const FrameIDA As Integer = 6
Const PeriodA As Integer = 7
Const SignA As Integer = 8
Const sizeA As Integer = 9
Const UnitA As Integer = 10
Const ResolutionA As Integer = 11
Const OffsetA As Integer = 12
Const MinA As Integer = 13
Const MaxA As Integer = 14
Const UnavailableA As Integer = 15
Const CodingA As Integer = 16
Const MeaningA As Integer = 17
Const FailureNameA As Integer = 18 'NB they should not be needed as public, see later
Const CommentA As Integer = 19
Const ReactionA As Integer = 20
Const DiagActivationA As Integer = 21
Const ConfigDIDA As Integer = 22
Const DevDIDA As Integer = 23
Const TransitionA As Integer = 24
Const ConfirmationA As Integer = 25
Const DisappearenceA As Integer = 26
Const DTCcodeA As Integer = 27
Const SignalReactionA As Integer = 28
Const SysReHeader As Integer = 29
Const SysReInfo As Integer = 30
Const SysReFunctionalReactivity As Integer = 31
Const SysReID As Integer = 32
Const SysReCustEffect As Integer = 33
Const RequirementIDA As Integer = 34
Const ValItemA As Integer = 35
Const ValStepA As Integer = 36
Const ValCommentA As Integer = 37
Const ValScriptA As Integer = 38
Const ValJiraA As Integer = 39
Const LastColumnA As Integer = 39 'easiest way? used to select and format the whole line

Const Purple As String = "PURPLE"
Const Orange As String = "ORANGE"
Const LightBlue As String = "LightBlue"
Const DarkOrange As String = "DarkOrange"
Const Blue As String = "Blue"
Const Thick As String = "Thick"
Const FrameColor As String = "DarkBlue"
Const White As String = "White"


Public A As Integer
Public D As Integer
Const Frame As String = "FRAME"
Const ECU As String = "ECU"


'Clean immediate windows
'Application.SendKeys "^g ^a {DEL}"

'============================================================================================================================================================================================================================================================================================
'      Global Variables Declaration
'============================================================================================================================================================================================================================================================================================


'Public CodingRangeD As Range 'note that the numeric value and the meaning are in two different columns in the MSRS format...

Sub CreateReactivityTable()

'============================================================================================================================================================================================================================================================================================
'           Setup
'============================================================================================================================================================================================================================================================================================

    Workbooks("TheDiagnosticFile_V11bb").Activate 'use that when debugging with several workbooks open
    Worksheets("Signals").Activate
    '----------------------------------------------------------------------------------------------------
    'Variables declaration and init
    '----------------------------------------------------------------------------------------------------

    Dim HeadersRangeD As Range: Set HeadersRangeD = Range("SignalName", Range("SignalName").End(xlToRight).Address)
    Dim SignalNameRangeD As Range: Set SignalNameRangeD = Range(HeadersRangeD.Find("Signal Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Signal Name").End(xlDown))
    Dim FrameNameRangeD As Range: Set FrameNameRangeD = Range(HeadersRangeD.Find("Frame Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Frame Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim FrameIDRangeD As Range: Set FrameIDRangeD = Range(HeadersRangeD.Find("Frame ID (Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Frame ID (Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim FrameSizeRangeD As Range: Set FrameSizeRangeD = Range(HeadersRangeD.Find("Frame Size (Bytes)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Frame Size (Bytes)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim FramePeriodRangeD As Range: Set FramePeriodRangeD = Range(HeadersRangeD.Find("Period (ms)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Period (ms)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim SignRangeD As Range: Set SignRangeD = Range(HeadersRangeD.Find("Value Type (Sign)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Value Type (Sign)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim SizeRangeD As Range: Set SizeRangeD = Range(HeadersRangeD.Find("Signal Size (Bits)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Signal Size (Bits)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim UnitRangeD As Range: Set UnitRangeD = Range(HeadersRangeD.Find("Unit", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Unit", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim ResolutionRangeD As Range: Set ResolutionRangeD = Range(HeadersRangeD.Find("Resolution (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Resolution (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim OffsetRangeD As Range: Set OffsetRangeD = Range(HeadersRangeD.Find("Offset (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Offset (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim MinValueRangeD As Range: Set MinValueRangeD = Range(HeadersRangeD.Find("Min (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Min (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim MaxValueRangeD As Range: Set MaxValueRangeD = Range(HeadersRangeD.Find("Max (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Max (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim UnavailableValueRangeD As Range: Set UnavailableValueRangeD = Range(HeadersRangeD.Find("Unavailable Value (Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Unavailable Value (Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim CodingRangeD As Range: Set CodingRangeD = Range(HeadersRangeD.Find("Coding (Bin/Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Coding (Bin/Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim MeaningRangeD As Range: Set MeaningRangeD = Range(HeadersRangeD.Find("Meaning", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Meaning", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))

    Dim ExpectedValueRangeD As Range: Set ExpectedValueRangeD = Range(HeadersRangeD.Find("Expected Value", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Expected Value", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim ConfigRangeD As Range
    Dim temp As String
    Dim ExpectedValueDec As String
    Dim ExpectedValueHex As String

    'Find the needed columns in the header list. By default is NOT CASE SENSITIVE
    'Using fixed header A1 cell called SignalName. In case of modification, this needs to be modify accordingly


    '----------------------------------------------------------------------------------------------------
    '"Arrival" sheet : TdR
    '----------------------------------------------------------------------------------------------------

    For Each Sheet In ThisWorkbook.Worksheets
        If Sheet.Name Like "ReactivityTable" Then
            Application.DisplayAlerts = False 'used to mute the message to confirm tab deletion
            Worksheets("ReactivityTable").Delete
            ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = "ReactivityTable"
            Exit For
        ElseIf Sheet Is Worksheets.Item(Worksheets.Count) = True Then
            ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = "ReactivityTable"
        End If
    Next Sheet

    Cells.Select
    With Selection
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .WrapText = True
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Rows.RowHeight = 15 'all rows same height
    Columns.NumberFormat = "@" 'to all?

    A = 1   'startingRow in Arrival sheet

    'TODO generic function createHeader in GFL, returning a range -> headerRangeA. but then selecting width of each column will be worse
    Call formatCell(A, ChannelNameA, "Channel", True, 10, , Thick, Orange, 15, 30)
    Call formatCell(A, ECUNameA, "ECU", True, 10, , Thick, Orange, 15, 30)
    Call formatCell(A, FrameNameA, "Frame Name", True, 10, , Thick, Orange, 15, 30)
    Call formatCell(A, IPDUNameA, "IPDU", False, 10, , Thick, Orange, 15, 30)
    Call formatCell(A, SignalNameA, "Signal Name", True, 10, , Thick, Purple, 30, 30)
    Call formatCell(A, FrameIDA, "Frame ID (Hex)", True, 10, , Thick, Purple, 15, 30)
    Call formatCell(A, PeriodA, "Period (ms)", True, 10, , Thick, Purple, 15, 30)
    Call formatCell(A, SignA, "Value Type (Sign)", True, 10, , Thick, Purple, 15, 30)
    Call formatCell(A, sizeA, "Size Size (Bits)", True, 10, , Thick, Purple, 15, 30)
    Call formatCell(A, UnitA, "Unit", True, 10, , Thick, Purple, 15, 30)
    Call formatCell(A, ResolutionA, "Resolution (Dec)", True, 10, , Thick, Purple, 15, 30)
    Call formatCell(A, OffsetA, "Offset (Dec)", True, 10, , Thick, Purple, 15, 30)
    Call formatCell(A, MinA, "Min (Dec)", True, 10, , Thick, Purple, 15, 30)
    Call formatCell(A, MaxA, "Max (Dec)", True, 10, , Thick, Purple, 15, 30)
    Call formatCell(A, UnavailableA, "Unavailable Value (Bin/Hex)", True, 10, , Thick, Purple, 15, 30)
    Call formatCell(A, CodingA, "Coding (Bin/Hex)", True, 10, , Thick, Purple, 15, 30)
    Call formatCell(A, MeaningA, "Meaning", True, 10, , Thick, Purple, 15, 30)
    Call CollapseColumnsLeft
    Range(Cells(A, FrameIDA), Cells(A, MeaningA)).Group

    Call formatCell(A, FailureNameA, "Failure Type", True, 10, , Thick, , 30, 30)
    Call formatCell(A, CommentA, "Comment log", False, 10, , Thick, , 15, 30)

    Call formatCell(A, ReactionA, "Reaction -> See conditions", False, 10, , Thick, LightBlue, 15, 30)
    Call formatCell(A, DiagActivationA, "Diag Activation", False, 10, , Thick, LightBlue, 15, 30)
    Call formatCell(A, ConfigDIDA, "Configuration DID", False, 10, , Thick, LightBlue, 15, 30)
    Call formatCell(A, DevDIDA, "Dev DID", False, 10, , Thick, LightBlue, 15, 30)
    Call formatCell(A, TransitionA, "Transition Value", False, 10, , Thick, LightBlue, 15, 30)
    Call formatCell(A, ConfirmationA, "Confirmation Time (ms)", False, 10, , Thick, LightBlue, 15, 30)
    Call formatCell(A, DisappearenceA, "Disappearence Time (ms)", False, 10, , Thick, LightBlue, 15, 30)
    Range(Cells(A, DiagActivationA), Cells(A, DisappearenceA)).Group

    Call formatCell(A, DTCcodeA, "DTC Code", True, 10, , Thick, Orange, 15, 30)
    Call formatCell(A, SignalReactionA, "Signal Reaction", True, 10, , Thick, DarkOrange, 15, 30)

    Call formatCell(A, SysReHeader, "Sys X", True, 10, , Thick, Blue, 15, 30)
    Call formatCell(A, SysReInfo, "Info", False, 10, , Thick, Blue, 15, 30)
    Call formatCell(A, SysReFunctionalReactivity, "Functional reactivity", False, 10, , Thick, Blue, 15, 30)
    Call formatCell(A, SysReID, "system reaction", False, 10, , Thick, Blue, 15, 30)
    Call formatCell(A, SysReCustEffect, "Customer effect", False, 10, , Thick, Blue, 15, 30)
    Range(Cells(A, SysReInfo), Cells(A, SysReCustEffect)).Group

    Call formatCell(A, RequirementIDA, "Requirement ID", False, 10, , Thick, , 15, 30)
    Call formatCell(A, ValItemA, "Val Item", False, 10, , Thick, , 15, 30)
    Call formatCell(A, ValStepA, "Val Step", False, 10, , Thick, , 15, 30)
    Call formatCell(A, ValCommentA, "Comment", False, 10, , Thick, , 15, 30)
    Call formatCell(A, ValScriptA, "Script", False, 10, , Thick, , 15, 30)
    Call formatCell(A, ValJiraA, "Jira", False, 10, , Thick, , 15, 30)
    Range(Cells(A, ValCommentA), Cells(A, ValJiraA)).Group

    A = 2
    'RequirementLine think about it
    'ECUNameA TODO


'---- Start the loop
    For D = 2 To SignalNameRangeD.Cells.Count

    'TODO recognise new ECU and new channel from message set. so far not, the modification will be done after, manually

        If (FrameNameRangeD.Cells(D, 1) <> FrameNameRangeD.Cells(D - 1, 1)) Then
            FrameName = FrameNameRangeD.Cells(D, 1).value
            FrameID = FrameIDRangeD.Cells(D, 1).value
            FramePeriod = FramePeriodRangeD.Cells(D, 1).value
            Debug.Print (FrameName)

            newHeaderTDRLine ("FRAME")

            'TODO missing frame test line
            'TODO line for separing ECUs?
            Cells(A, FrameNameA).Select
            With Selection
                .value = FrameName
                .Font.Color = RGB(48, 84, 150)
                .Interior.Color = RGB(48, 84, 150)
            End With
            Cells(A, FrameIDA).value = FrameID
            Cells(A, PeriodA).value = FramePeriod
            Cells(A, SignalNameA).value = "Frame"
            Cells(A, FailureNameA).value = "Missing frame"
            A = A + 1
        End If

        Debug.Print (SignalNameRangeD.Cells(D, 1).value)
        Cells(A, FrameNameA).Select
        With Selection
            .value = FrameName
            .Font.Color = RGB(48, 84, 150)
            .Interior.Color = RGB(48, 84, 150)
        End With
        Cells(A, FrameIDA).value = FrameID
        Cells(A, PeriodA).value = FramePeriod
        Cells(A, SignalNameA).value = SignalNameRangeD.Cells(D, 1).value
        Cells(A, UnavailableA).value = UnavailableValueRangeD.Cells(D, 1).value
        Cells(A, FailureNameA).value = "tbd"

        'TODO not used ?
        A = A + 1

    Next D

    Columns("A:T").AutoFilter
    Call Collapse_All
    'Range("A1", Cells(A - 1, HeadersRangeA.Count)).Select 'TODO selection at the end of the whole table to copy-aste it directly on other tab



End Sub

Public Function newHeaderTDRLine(HeaderType As String)

    Select Case HeaderType
        Case "FRAME"
            'Call formatCell(A, FrameNameA, FrameName, False, "", White, "", FrameColor, "", "")
            Cells(A, FrameNameA).value = FrameName
            Cells(A, FrameNameA).Font.Color = RGB(255, 255, 255)
            Range(Cells(A, FrameNameA), Cells(A, LastColumnA)).Interior.Color = RGB(48, 84, 150)
            Range(Cells(A, FrameNameA), Cells(A, LastColumnA)).Font.Color = RGB(248, 248, 248)
            Range(Cells(A, FailureNameA), Cells(A, LastColumnA)).value = "."
        Case "CHANNEL"
            Cells(A, ChannelNameA).value = ChannelName
            Range(Cells(A, ChannelNameA), Cells(A, LastColumnA)).Interior.Color = RGB(34, 43, 53)
            Range(Cells(A, ChannelNameA), Cells(A, LastColumnA)).Font.Color = RGB(34, 43, 53)
            Range(Cells(A, FailureNameA), Cells(A, LastColumnA)).value = "."
        Case "IPDU"
            'Cells(A, IPDUNameA).Value = IPDUName 'TODO
            Range(Cells(A, IPDUNameA), Cells(A, LastColumnA)).Interior.Color = RGB(142, 169, 219)
            Range(Cells(A, IPDUNameA), Cells(A, LastColumnA)).Font.Color = RGB(142, 169, 219)
            Range(Cells(A, FailureNameA), Cells(A, LastColumnA)).value = "."
        Case "ECU"
            Cells(A, ECUNameA).value = ECUName
            Range(Cells(A, ECUNameA), Cells(A, LastColumnA)).Interior.Color = RGB(32, 55, 100)
            Range(Cells(A, ECUNameA), Cells(A, LastColumnA)).Font.Color = RGB(32, 55, 100)
            Range(Cells(A, FailureNameA), Cells(A, LastColumnA)).value = "."
    End Select


    A = A + 1
End Function
