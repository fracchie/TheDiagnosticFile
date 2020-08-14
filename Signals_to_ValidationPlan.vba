Option Explicit

'TODO generalise this function

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

Public Sheet As Worksheet
Public FrameName As String
Public FramePeriod As Double
Public SignalName As String
Public MacroState As Boolean
Public StepStatusColumn As Integer
Public ItemStatusColumn As Integer
Public A As Integer
Public D As Integer

'Public CodingRangeD As Range 'note that the numeric value and the meaning are in two different columns in the MSRS format...

Sub MessageCheckValidationPlan()

'============================================================================================================================================================================================================================================================================================
'           Setup
'============================================================================================================================================================================================================================================================================================

    Workbooks("TheDiagnosticFile_V11_beta").Activate 'use that when debugging with several workbooks open
    Worksheets("Signals").Activate
    '----------------------------------------------------------------------------------------------------
    'Variables declaration and init
    '----------------------------------------------------------------------------------------------------

    Dim HeadersRangeD As Range
    Dim SignalNameRangeD As Range
    Dim FrameNameRangeD As Range
    Dim FramePeriodRangeD As Range
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
    Dim ScriptCreation As Boolean
    Dim RequestColumn As Integer
    Dim ResponseColumn As Integer
    Dim StatusItemColumn As Integer
    Dim StatusStepColumn As Integer

    Dim ExpectedValueDec As String
    Dim ExpectedValueHex As String

    'Find the needed columns in the header list. By default is NOT CASE SENSITIVE
    'Using fixed header A1 cell called SignalName. In case of modification, this needs to be modify accordingly
    Set HeadersRangeD = Range("SignalName", Range("SignalName").End(xlToRight).Address)
    Set SignalNameRangeD = Range(HeadersRangeD.Find("Signal Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Signal Name").End(xlDown))
    Set FrameNameRangeD = Range(HeadersRangeD.Find("Frame Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Frame Name").End(xlDown))
    Set FramePeriodRangeD = Range(HeadersRangeD.Find("Period (ms)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Period (ms)").End(xlDown))
    Set UnavailableValueRangeD = Range(HeadersRangeD.Find("Unavailable Value (Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Unavailable Value (Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set MinValueRangeD = Range(HeadersRangeD.Find("Min (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Min (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set MaxValueRangeD = Range(HeadersRangeD.Find("Max (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Max (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set ResolutionRangeD = Range(HeadersRangeD.Find("Resolution (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Resolution (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set SizeRangeD = Range(HeadersRangeD.Find("Signal Size (Bits)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Signal Size (Bits)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set SignRangeD = Range(HeadersRangeD.Find("Value Type (Sign)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Value Type (Sign)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set OffsetRangeD = Range(HeadersRangeD.Find("Offset (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Offset (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set CodingRangeD = Range(HeadersRangeD.Find("Coding (Bin/Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Coding (Bin/Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set ExpectedValueRangeD = Range(HeadersRangeD.Find("Expected Value", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Expected Value", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))


    ScriptCreation = False
    'TODO take scriptCreation value from button on page
    If (ScriptCreation = True) Then
        'TODO create script with function
    End If

    'Create Tab
    CreateNewTab ("Tx_Msg_Val")

    'ExpectedValueRangeD.Select

'---- Start the loop

    A = 1
    'TODO first part, A to G
    RequestColumn = 1 'A
    'TODO second part, H to N
    ResponseColumn = 8 'H
    'TODO O and P status cells
    ItemStatusColumn = 15
    StepStatusColumn = 16

    Range("A1:P1").Merge
    Range("A1:P1").Interior.Color = RGB(255, 192, 0)
    Range("A1:P1").RowHeight = 30
    Range("A1:P1").Font.Bold = 1
    Range("A1:P1").HorizontalAlignment = xlCenter
    Range("A1:P1").VerticalAlignment = xlCenter


    Cells(A, 1).value = "Signal Validation Plan"


    A = 2
    For D = 2 To SignalNameRangeD.Cells.Count

        If (FrameNameRangeD.Cells(D, 1) <> FrameNameRangeD.Cells(D - 1, 1)) Then
            FrameName = FrameNameRangeD.Cells(D, 1)
            FramePeriod = FramePeriodRangeD.Cells(D, 1)
            MacroState = newValidationline("Request", "Frame " + FrameName, True)
            MacroState = newValidationline("Response", "Check period is " + Str(FramePeriod), True)
            MacroState = newValidationline("Response", "Check DLC", True)
        End If

        SignalName = SignalNameRangeD.Cells(D, 1).value
        MacroState = newValidationline("Request", "'-> Signal " + SignalName, False)
        MacroState = newValidationline("Response", "Check value", True)

    Next D

    'MsgBox "File Created"
End Sub

Function CreateFile(fileName As String)

    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(&H0&, "Choose file's path. Consider using the folder 'input' of the DST script generator", &H1&)

    Set objFolderItem = objFolder.Items.Item
    filePath = objFolderItem.Path
    'fileName = "SignalVal.xml"
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

End Function

Function newValidationline(ReqOrRespOrHead As String, text As String, Optional ByVal Status As Boolean = False) As Boolean

    If ReqOrRespOrHead = "Request" Then
        Range(Cells(A, 1), Cells(A, 7)).Merge
        Cells(A, 1).value = text
        Range(Cells(A, 8), Cells(A, 14)).Merge
        If Status = True Then
            Range(Cells(A, ItemStatusColumn), Cells(A, StepStatusColumn)).Merge
            Cells(A, ItemStatusColumn).value = "TBV"
            Cells(A, ItemStatusColumn).HorizontalAlignment = xlCenter
            Cells(A, ItemStatusColumn).VerticalAlignment = xlCenter
        End If
    ElseIf ReqOrRespOrHead = "Response" Then
        Range(Cells(A, 1), Cells(A, 7)).Merge
        Range(Cells(A, 8), Cells(A, 14)).Merge
        Cells(A, 8).value = text
        If Status = True Then
            Cells(A, StepStatusColumn).value = "TBV"
            Cells(A, StepStatusColumn).HorizontalAlignment = xlCenter
            Cells(A, StepStatusColumn).VerticalAlignment = xlCenter
        End If
        Cells(A, StepStatusColumn).value = "TBV"
    ElseIf ReqOrRespOrHead = "Header" Then
        Range(Cells(A, 1), Cells(A, 14)).Merge
        Cells(A, 1).value = text
        If Status = True Then
            Range(Cells(A, ItemStatusColumn), Cells(A, StepStatusColumn)).Merge
            Cells(A, ItemStatusColumn).value = "TBV"
            Cells(A, ItemStatusColumn).HorizontalAlignment = xlCenter
            Cells(A, ItemStatusColumn).VerticalAlignment = xlCenter
        End If
    End If

    A = A + 1
    newValidationline = True

End Function
