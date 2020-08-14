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
    Dim HeadersRangeD As Range
    Dim SignalNameRangeD As Range
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
    Set UnavailableValueRangeD = Range(HeadersRangeD.Find("Unavailable Value (Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Unavailable Value (Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set MinValueRangeD = Range(HeadersRangeD.Find("Min (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Min (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set MaxValueRangeD = Range(HeadersRangeD.Find("Max (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Max (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set ResolutionRangeD = Range(HeadersRangeD.Find("Resolution (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Resolution (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set SizeRangeD = Range(HeadersRangeD.Find("Signal Size (Bits)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Signal Size (Bits)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set SignRangeD = Range(HeadersRangeD.Find("Value Type (Sign)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Value Type (Sign)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set OffsetRangeD = Range(HeadersRangeD.Find("Offset (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Offset (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set CodingRangeD = Range(HeadersRangeD.Find("Coding (Bin/Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Coding (Bin/Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set ExpectedValueRangeD = Range(HeadersRangeD.Find("Expected Value", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Expected Value", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))

'----- Output File Creation ------
    Dim filePath As String
    Dim fileName As String
    Dim objShell As Object, objFolder As Object, objFolderItem As Object

    fileName = Range("Signal_Read_Script_Name").value + ".xml" 'Specific cell in the file with this name. Now in Cell(B,2)

    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(&H0&, "Choose file's path. Consider using the folder 'input' of the DST script generator", &H1&)

    Set objFolderItem = objFolder.Items.Item
    filePath = objFolderItem.Path
    MsgBox filePath & "\" & fileName

    Dim TempByteSent As String

    Dim MyFSO As New FileSystemObject
    If MyFSO.FolderExists(filePath) Then
        'MsgBox "The Folder already exists"
    Else
        MyFSO.CreateFolder (filePath)
    End If

    Dim FileOut As TextStream
    Set FileOut = MyFSO.CreateTextFile(filePath + "\" + fileName, True, True)

    ExpectedValueRangeD.Select

    FileOut.Write ("testCase ")
    FileOut.Write (Range("Signal_Read_Script_Name").value)
    FileOut.Write ("(){")
    FileOut.WriteBlankLines (1)

'---- Start the loop
    For D = 2 To SignalNameRangeD.Cells.Count

        Debug.Print (SignalNameRangeD.Cells(D, 1).value)

        If (ExpectedValueRangeD.Cells(D, 1).value <> "") Then

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
        Else
        ' What to do if expected value is nothing (i.e. "")? expected nothing i would say, just read the value. TODO Check How

        End If

    Next D

    FileOut.WriteLine ("}")

    MsgBox "File Created"
End Sub
