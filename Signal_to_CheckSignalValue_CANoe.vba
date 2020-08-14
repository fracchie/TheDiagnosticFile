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

Sub SignalsToCheckSignalValueCANoeScriptGen()

'============================================================================================================================================================================================================================================================================================
'           Setup
'============================================================================================================================================================================================================================================================================================

    Worksheets("Signals").Activate
    '----------------------------------------------------------------------------------------------------
    'Variables declaration and init
    '----------------------------------------------------------------------------------------------------

    Dim A As Integer
    Dim D As Integer
    'Find the needed columns in the header list using fixed header A1 cell called SignalName. In case of modification, this needs to be modify accordingly. By default is NOT CASE SENSITIVE
    Dim HeadersRangeD As Range: Set HeadersRangeD = Range("SignalName", Range("SignalName").End(xlToRight).Address)
    Dim SignalNameRangeD As Range: Set SignalNameRangeD = Range(HeadersRangeD.Find("Signal Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Signal Name").End(xlDown))
    Dim tableRows As Double: tableRows = SignalNameRangeD.Rows.Count
    Dim FrameNameRangeD As Range: Set FrameNameRangeD = Range(HeadersRangeD.Find("Frame Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Frame Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim UnavailableValueRangeD As Range: Set UnavailableValueRangeD = Range(HeadersRangeD.Find("Unavailable Value (Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Unavailable Value (Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim MinValueRangeD As Range: Set MinValueRangeD = Range(HeadersRangeD.Find("Min (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Min (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim MaxValueRangeD As Range: Set MaxValueRangeD = Range(HeadersRangeD.Find("Max (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Max (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim ResolutionRangeD As Range: Set ResolutionRangeD = Range(HeadersRangeD.Find("Resolution (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Resolution (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim SizeRangeD As Range: Set SizeRangeD = Range(HeadersRangeD.Find("Signal Size (Bits)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Signal Size (Bits)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim SignRangeD As Range: Set SignRangeD = Range(HeadersRangeD.Find("Value Type (Sign)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Value Type (Sign)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim OffsetRangeD As Range: Set OffsetRangeD = Range(HeadersRangeD.Find("Offset (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Offset (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim CodingRangeD As Range: Set CodingRangeD = Range(HeadersRangeD.Find("Coding (Bin/Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Coding (Bin/Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim ExpectedValueRangeD As Range: Set ExpectedValueRangeD = Range(HeadersRangeD.Find("Expected Value", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Expected Value", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))

    ExpectedValueRangeD.Select
    Dim signalName As String
    Dim temp As String
    Dim ExpectedValueDec As String
    Dim ExpectedValueHex As String

'----- Output File Creation ------
    Dim filePath As String
    Dim fileName As String
    Dim objShell As Object, objFolder As Object, objFolderItem As Object

    fileName = Range("Signal_Read_Script_Name").value + ".xml" 'Specific cell in the file with this name. Now in Cell(B,2)

    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(&H0&, "Choose file's path. Consider using the folder 'input' of the DST script generator", &H1&)

    Set objFolderItem = objFolder.Items.Item
    filePath = objFolderItem.Path
    'MsgBox filePath & "\" & fileName

    Dim TempByteSent As String

    Dim MyFSO As New FileSystemObject
    If MyFSO.FolderExists(filePath) Then
        'MsgBox "The Folder already exists"
    Else
        MyFSO.CreateFolder (filePath)
    End If

    Dim FileOut As TextStream
    Set FileOut = MyFSO.CreateTextFile(filePath + "\" + fileName, True, True)

    FileOut.Write ("testCase ")
    FileOut.Write (Range("Signal_Read_Script_Name").value)
    FileOut.Write ("(){")
    FileOut.WriteBlankLines (1)

'---- Start the loop
    For D = 2 To SignalNameRangeD.Cells.Count

        signalName = (SignalNameRangeD.Cells(D, 1).value)

        Debug.Print (signalName)

        If (IsEmpty(ExpectedValueRangeD.Cells(D, 1).value)) Then

            FileOut.WriteLine ("testStep(" + Chr(34) + Chr(34) + "," + Chr(34) + signalName + " = %f" + Chr(34) + ",$" + signalName + ");")
            FileOut.WriteBlankLines (1)
        'testStep("","GADE = %f",getSignal(IC_T32019::BCM_A7::GADE));



        Else '(ExpectedValueRangeD.Cells(D, 1).value <> "") Then

           ExpectedValueHex = ExpectedValueRangeD.Cells(D, 1).value
           Debug.Print ("--- Hex: " + ExpectedValueHex)

           If (CodingRangeD.Cells(D, 1).value = "") Then 'it is a numeric value
               If (SignRangeD.Cells(D, 1).value = "Unsigned") Then
                   Debug.Print ("Numeric Unsigned")
                   ExpectedValueDec = CLng("&H" & ExpectedValueHex) * ResolutionRangeD.Cells(D, 1).value + OffsetRangeD.Cells(D, 1).value
               Else 'is a 2complement signed value
                   Debug.Print ("Numeric Signed 2 complement")
                   temp = HexToBin(ExpectedValueHex)
                   ExpectedValueDec = BinToDec(temp, 2) * ResolutionRangeD.Cells(D, 1).value + OffsetRangeD.Cells(D, 1).value
               End If
           Else 'it is a list
               Debug.Print ("List")
               ExpectedValueDec = CLng("&H" & ExpectedValueHex)
           End If
           Debug.Print ("--- Dec: " + ExpectedValueDec)

           FileOut.Write ("if ($")
           FileOut.Write (SignalNameRangeD.Cells(D, 1).value)
           FileOut.Write (" == ")
           FileOut.Write (ExpectedValueDec)
           FileOut.Write (") {")
           FileOut.WriteBlankLines (1)
           FileOut.Write ("     TestStepPass(")
           FileOut.Write (Chr(34) + Chr(34)) ' to output -> ""
           FileOut.Write (", " + Chr(34))
           FileOut.Write (SignalNameRangeD.Cells(D, 1).value + " = " + ExpectedValueDec)
           FileOut.Write (Chr(34) + ");")
           FileOut.WriteBlankLines (1)
           FileOut.Write ("       } else {")
           FileOut.WriteBlankLines (1)
           FileOut.Write ("             TestStepFail(" + Chr(34) + Chr(34) + ", " + Chr(34))
           FileOut.Write (SignalNameRangeD.Cells(D, 1).value + " = %f EXPECTED: " + ExpectedValueDec + Chr(34) + ",$" + SignalNameRangeD.Cells(D, 1).value + ");")
           FileOut.WriteBlankLines (1)
           FileOut.Write ("       }")
           FileOut.WriteBlankLines (2)

           FileOut.WriteBlankLines (1)

        End If

    Next D

    FileOut.WriteLine ("}")

    MsgBox "File Created in " + filePath & "\" & fileName
End Sub
