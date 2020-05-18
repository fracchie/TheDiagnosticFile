' ===========================
' the whole script here is generated supposing that, on CANoe side, functions have been created such as
'ReadDID(DIDNumber as String, Optional ByVal ExpectedValueDec as Double)
'ReadParameter(DIDNumber As String, ParameterName As String, Optional ByVal ExpectedValueDec as Double)
'WriteDID(DIDNumber as String, DIDContent as String)
'WriteParameter(DIDNumber as String, ParameterName As String, ValueDec As Double)


Option Explicit

'============================================================================================================================================================================================================================================================================================
'      Global Variables Declaration
'============================================================================================================================================================================================================================================================================================
Public ECUName As String
Public ChannelName As String
Public FrameName As String
Public SignalName As String
Public FrameID As String
Public FramePeriod As Integer
Public D As Integer

Public IndentIndex As Integer
Public FailureType As String
Public ConfirmationTime As Double
Public DisappearenceTime As Double
Public TotalConfig As String
Public DTC As String
Public FaultType As String
Public configurationDIDBrut As String
Public ConfigurationDIDListed() As String
Public diagActivationBrut As String
Public diagActivationListed() As String
Public SignalReactionsBrut As String
Public SignalReactionListed() As String
Public SysReIDs As String
Public RequirementID As String
Public testConfigArray() As String
Public unavailableValue As String

Public Sheet As Worksheet

Public FileOut As TextStream
Public FrameArxmlAddress As String 'TODO link between this message set and arxml file variables -> i.e. something like IL_CAN1::NODES::N_GWBridge::MESSAGES::framename

Const PRESENT As String = "PRESENT"
Const MEMORISED As String = "MEMORISED"
Const NOTPRESENT As String = "NOTPRESENT"
Const LESSTHAN As String = "LESSTHAN"
Const MORETHAN As String = "MORETHAN"

'Clean immediate windows
'Application.SendKeys "^g ^a {DEL}"

Sub Reactivity_Table_Script_Gen()

'======================================================================================================
' START
' Repeating the process for each selected session, considering also the RO RW switch
Debug.Print ("==================================================")
Debug.Print ("Application.SendKeys " + Chr(34) + "^g ^a {DEL}")
Debug.Print ("==================================================")

Debug.Print ("")
Debug.Print ("$$$$$$$$$ START $$$$$$$$$")
Debug.Print ("")
'======================================================================================================


'============================================================================================================================================================================================================================================================================================
'           Setup
'============================================================================================================================================================================================================================================================================================

    'Workbooks("TheDiagnosticFile_V11error").Activate 'use that when debugging with several workbooks open, to avoid any bullshit. comment/remove it when not debugging anymore
    Worksheets("TdR").Activate
    '----------------------------------------------------------------------------------------------------
    'Variables declaration and init
    '----------------------------------------------------------------------------------------------------

    Call Expand_All

    Dim HeadersRangeD As Range: Set HeadersRangeD = Range(Range("HereBelow").Offset(1, 0).Address, Range("HereBelow").Offset(1, 0).End(xlToRight).Address)
    Dim ChannelRangeD As Range: Set ChannelRangeD = Range(HeadersRangeD.Find("Channel", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Channel", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim ECUNameRangeD As Range: Set ECUNameRangeD = Range(HeadersRangeD.Find("ECU", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("ECU", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim FrameNameRangeD As Range: Set FrameNameRangeD = Range(HeadersRangeD.Find("Frame Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Frame Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    'TODO IPDU
    Dim SignalNameRangeD As Range: Set SignalNameRangeD = Range(HeadersRangeD.Find("Signal Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Signal Name").End(xlDown))
    Dim FrameIDRangeD As Range: Set FrameIDRangeD = Range(HeadersRangeD.Find("Frame ID (Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Frame ID (Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    'Dim FrameSizeRangeD As Range: Set FrameSizeRangeD = Range(HeadersRangeD.Find("Frame Size (Bytes)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Frame Size (Bytes)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim FramePeriodRangeD As Range: Set FramePeriodRangeD = Range(HeadersRangeD.Find("Period (ms)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Period (ms)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim SignRangeD As Range: Set SignRangeD = Range(HeadersRangeD.Find("Value Type (Sign)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Value Type (Sign)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim SizeRangeD As Range: Set SizeRangeD = Range(HeadersRangeD.Find("Signal Size (Bits)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Signal Size (Bits)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim UnitRangeD As Range: Set UnitRangeD = Range(HeadersRangeD.Find("Unit", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Unit", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim ResolutionRangeD As Range: Set ResolutionRangeD = Range(HeadersRangeD.Find("Resolution (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Resolution (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim OffsetRangeD As Range: Set OffsetRangeD = Range(HeadersRangeD.Find("Offset (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Offset (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim MinValueRangeD As Range: Set MinValueRangeD = Range(HeadersRangeD.Find("Min (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Min (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim MaxValueRangeD As Range: Set MaxValueRangeD = Range(HeadersRangeD.Find("Max (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Max (Dec)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim UnavailableValueRangeD As Range: Set UnavailableValueRangeD = Range(HeadersRangeD.Find("Unavailable Value (Bin/Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Unavailable Value (Bin/Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim CodingRangeD As Range: Set CodingRangeD = Range(HeadersRangeD.Find("Coding (Bin/Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Coding (Bin/Hex)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim MeaningRangeD As Range: Set MeaningRangeD = Range(HeadersRangeD.Find("Meaning", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Meaning", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))

    Dim FailureRangeD As Range: Set FailureRangeD = Range(HeadersRangeD.Find("Failure Type", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Failure Type").End(xlDown))
    Dim DiagActivationRangeD As Range: Set DiagActivationRangeD = Range(HeadersRangeD.Find("Diag Activation", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Diag Activation").End(xlDown))
    Dim ConfigRangeD As Range: Set ConfigRangeD = Range(HeadersRangeD.Find("Configuration DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Configuration DID").End(xlDown))
    Dim DevDIDRangeD As Range: Set DevDIDRangeD = Range(HeadersRangeD.Find("Dev DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Dev DID").End(xlDown))
    'TODO transition?
    'TODO consider if keeping this format forcing to write it in (ms) or as before max(200ms, 3T)...
    Dim ConfirmationTimeRangeD As Range: Set ConfirmationTimeRangeD = Range(HeadersRangeD.Find("Confirmation Time (ms)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Confirmation Time (ms)").End(xlDown))
    Dim DisappearenceTimeRangeD As Range: Set DisappearenceTimeRangeD = Range(HeadersRangeD.Find("Disappearence Time (ms)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Disappearence Time (ms)").End(xlDown))

    Dim DTCRangeD As Range: Set DTCRangeD = Range(HeadersRangeD.Find("DTC Code", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("DTC Code").End(xlDown))
    Dim SignalReactionRangeD As Range: Set SignalReactionRangeD = Range(HeadersRangeD.Find("Signal Reaction", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Signal Reaction").End(xlDown))

    Dim ScriptRangeD As Range: Set ScriptRangeD = Range(HeadersRangeD.Find("Script", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Script").End(xlDown))

    Dim temp As String
    Dim ScriptNameSpecific As String
    Dim i As Integer
    Dim l As Integer
    Dim NumberOfFileOut As Integer: NumberOfFileOut = 0

    '----------------------------------------------------------------------------------------------------
    '"Arrival" file
    '----------------------------------------------------------------------------------------------------

'================ .xml file declaration ==================
    Dim filePath As String
    Dim fileName As String
    Dim objShell As Object, objFolder As Object, objFolderItem As Object
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(&H0&, "Choose file's path", &H1&)
    Set objFolderItem = objFolder.Items.Item
    filePath = objFolderItem.Path
'================ .xml file declaration ==================

    Dim ServiceRangeA As Range: Set ServiceRangeA = Range("B2", Range("B2").End(xlDown).Address)
    Dim CommandSentRangeA As Range: Set CommandSentRangeA = Range("I2", Range("I2").End(xlDown).Address)
    Dim ResponseExpectedRangeA As Range: Set ResponseExpectedRangeA = Range("J2", Range("J2").End(xlDown).Address)


'---- Start the loop
    For D = 2 To FailureRangeD.Cells.Count

        FailureType = FailureRangeD.Cells(D, 1).value
        Select Case FailureType

            Case "Missing Frame", "Unavailable", "Out Of Range", "Not Used", "NotUsed/OutOfRange" 'TODO for the last two, consider
                'Call getInfoTestLine(D) 're-moved here all the lines of getInfoTestLine because otherwise all the ranged must be public, and beside waste of memory and heavier code, i cannot dim and set in the same line
                'FailureType = FailureRangeD.Cells(D, 1).value 'Frame absent, CRC, Clock, Unavailable value, OutOfRange/NotUsed
                ChannelName = ChannelRangeD.Cells(D, 1)
                ECUName = ECUNameRangeD.Cells(D, 1)
                FrameName = FrameNameRangeD.Cells(D, 1)
                'TODO? IPDU manage difference CAN CANFD
                FrameID = FrameIDRangeD.Cells(D, 1)
                FramePeriod = FramePeriodRangeD.Cells(D, 1)
                SignalName = SignalNameRangeD.Cells(D, 1).value 'if whole frame, will be "Frame"
                unavailableValue = UnavailableValueRangeD.Cells(D, 1).value
                Debug.Print ("---------------------------------------------------------------------")
                Debug.Print ("-- ECU: " + ECUName + "; Frame : " + FrameName + "; Signal: " + SignalName + "; Failure: " + FailureType)
                ScriptNameSpecific = Str(Range("TDR_V").value) + "_" + ECUName + "_" + FrameName + "_" + SignalName + "_" + FailureType 'TODO if two line with same failure type - no case so far -  consider to do something here
                '============ xml file creation =============
                'fileName = "PVal_DID.xml"
                'fileName = Range("ScriptName").value 'Cells C2 contains the name of the script. The reference of this cell has been changed in "ScriptName"
                fileName = ScriptNameSpecific
                ' can put a switch "output txt file DST?" and use an IF here
                'use the info in new tab PVal to create the script DST

                Dim TempByteSent As String

                Dim MyFSO As New FileSystemObject
                If MyFSO.FolderExists(filePath) Then
                    'MsgBox "The Folder already exists"
                Else
                    MyFSO.CreateFolder (filePath) '<- Here the
                End If

                'Dim FileOut As TextStream 'Declared as public
                Set FileOut = MyFSO.CreateTextFile(filePath + "\" + fileName + ".xml", True, True)
                Call CanoeInitTestScript(ScriptNameSpecific)
                '============ !xml file creation =============
                DTC = DTCRangeD.Cells(D, 1).value 'temp: for the moment, could be formatted as $XXXX-YY'
                Debug.Print ("> DTC : " + DTC)
                If InStr(DTC, "-") Then
                    DTC = Left(DTCRangeD.Cells(D, 1).value, InStr(DTCRangeD.Cells(D, 1).value, "-") - 1)
                    FaultType = Right(DTCRangeD.Cells(D, 1).value, Len(DTCRangeD.Cells(D, 1)) - InStr(DTCRangeD.Cells(D, 1).value, "-"))
                    ' Debug.Print (DTC + " - " + FaultType)
                Else
                    FaultType = Empty
                    ' Debug.Print (DTC)
                End If
                diagActivationBrut = DiagActivationRangeD.Cells(D, 1).value

                configurationDIDBrut = ConfigRangeD.Cells(D, 1).value
                Call processConfigurationDID

                'TODO Decide how to combine the two. For the moment, only configurationDID
                ConfirmationTime = ConfirmationTimeRangeD.Cells(D, 1).value
                DisappearenceTime = DisappearenceTimeRangeD.Cells(D, 1).value

                'TODO list the signals, expanding the system reactions if some'
                SignalReactionsBrut = SignalReactionRangeD.Cells(D, 1)
                Call ProcessSignalReaction

                'Process conditions
                Call processConfigurationDID 'Function is writin in public testConfigArray
                'for each condition, play test scenario

                'TODO here, loop for each line of testConfigArray. For the moment, just once!
                'For i = 0 To UBound(testConfigArray)
                    'write config of line element i of array config; Imagine testConfigArray to be of the kind "TRUE: DID_1 = X And DID_2 = Y" or "FALSE : DID_1 = Y DID_2 = Y". if there is nothing, it will simply do the test considering one iteration for what concerns config DID
                    'TODO Loop with function write config, or not if nothing

                    Call createPositiveTestScript
                    'read pre condition (DTC and signal), checking that DTC is not already reaised and that signal output are <> than reaction (if one; otherwise, just read)

                    'TODO if positiveConditionCase Then
                        'TODO test positiveTimingCase -> expect TRUE
                        'TODO test negativeTimingCase -> expect FALSE
                    'TODO else
                        'TODO test positiveTimingCase -> expect FALSE
                    'End If
                'Next i

                ' TODO format Close CANoe script
                Call CanoeEndScript(1) 'CHECK number of paenthesis needed to close the script
                NumberOfFileOut = NumberOfFileOut + 1
                'Close file, otherwise it will leave the reference to the file, and will not allow you to re-write this file if launching another macro with the same file name
                FileOut.Close
                Set MyFSO = Nothing
                Set FileOut = Nothing

        Case Else 'FailureType = "", NA, "_", and whatever else not listed in previous case
                ScriptNameSpecific = "."
        End Select

        ScriptRangeD.Cells(D, 1).value = ScriptNameSpecific

    Next D

    MsgBox "created " & NumberOfFileOut & " scripts in " & filePath
    'TODO open file after creation
    'MyFSO.OpenTextFile (Temp)

    Call Collapse_All
End Sub

Public Function ScriptNewFrameLine(Frame As String) As Boolean

    FileOut.Write ("//------------- frame ")
    FileOut.Write (Frame)
    FileOut.Write (" ; Period ")
    FileOut.Write (FramePeriod)
    FileOut.WriteBlankLines (2)

End Function

Public Function CanoeInitTestScript(ScriptName As String)
    Dim temp As String

    FileOut.WriteLine ("testCase(" + ScriptName + "){")

    'CanoeInitTestScript (temp)
    'FileOut.Write ("Testcase ")
    'FileOut.Write (fileName)
    'FileOut.WriteLine ("()")
    'FileOut.WriteLine ("{")
    'FileOut.WriteBlankLines (2)

    'Note: testCaseDescription useless?
    'FileOut.Write ("TestCaseDescription (")
    'FileOut.Write (Chr(34)) 'quote mark "
    'FileOut.Write (Chr(34)) 'quote mark "
    'FileOut.Write (");")

    'FileOut.WriteBlankLines (2)

    'Logging start
    'FileOut.Write ("TestStepBegin(")
    'FileOut.Write (Chr(34)) 'quote mark "
    'FileOut.Write (Chr(34)) 'quote mark "
    'FileOut.Write (",")
    'FileOut.Write (Chr(34)) 'quote mark "
    'FileOut.Write (Chr(34)) 'quote mark "
    'FileOut.Write (");") 'TODO check how to start logging -> Ask


    'FileOut.WriteLine ("<?xml version=" + Chr(34) + "1.0" + Chr(34) + " encoding=" + Chr(34) + "windows-1252" + Chr(34) + " ?> ")
    'FileOut.Write ("<?xml version=")
    'FileOut.Write (Chr(34))
    'FileOut.Write ("1.0")
    'FileOut.Write (Chr(34))
    'FileOut.Write (" encoding=")
    'FileOut.Write (Chr(34))
    'FileOut.Write ("windows-1252")
    'FileOut.Write (Chr(34))
    'FileOut.Write ("?>")
    'FileOut.WriteBlankLines (1)



    'FileOut.WriteBlankLines (2)

End Function

Sub SpaceIndent(index As Integer)
    Dim i As Integer
    For i = 0 To index
        FileOut.Write ("  ")
    Next i
End Sub

Function CheckConfigDIDs(CellContent As String) 'TODO process cell content
    'format: DID_Name.Param_Name = X & DID_Name.Param_Name = Y

End Function

Function CanoeWriteComment(text As String) As String
    Dim temp As String

    temp = "// " + text
    CanoeWriteComment = temp

End Function

Function ProcessConfiguration(CellContent As String) As String() 'CHECK if array as argument. but consider that w VlbF willhave same inpact of array, only in one line
    Dim list() As String
    Dim parameterList() As String
    Dim l As Integer
    Dim i As Integer

    'WriteCommentConfig (CellContent)

    list = Split(CellContent, vbLf)
    ' each even raw of list is a condition
    ' each odd raw is an OR
    ' list will end with an even raw
    l = 0
    For l = 0 To UBound(list)

        If l Mod (2) = 0 Then
            Debug.Print ("- Condition: " + list(l))
            FileOut.WriteLine ("// " + list(l))

            parameterList = Split(list(l), " & ")
            For i = 0 To UBound(parameterList)
                Debug.Print (parameterList(i))
                ' Go catching Parameter info in "parameters" tab
            Next i
        End If




        'Cells(A, MnemoColA).Value = Left(list(l), InStr(list(l), "=") - 1)
        'Cells(A, SizeColA).Value = Right(list(l), Len(list(l)) - InStr(list(l), "="))
'                    Label = Right(List(L), Len(List(L)) - InStr(List(L), "="))
'                    Cells(A, SizeColA).Value = Label
'                       Cells(A, MnemoColA).Value = Split(List(L), "=")
'                       Cells(R, 2).Value = Value
'                       Label = Right(List(L), Len(List(L)) - InStr(List(L), "="))
'                       Cells(R, 3).Value = Label



    Next l

End Function

Public Function getReactionFromTable(reaction As String) As String()

'NOT BEING USED ANYMORE: MERGED IN processSignalReaction()
    Dim output() As String
    Dim signal As String
    Dim value As String
    Worksheets("System_Reaction").Activate
    Dim HeaderRangeSys As Range: Set HeaderRangeSys = Range(Range("SysSignalName").Address, Range("SysSignalName").End(xlToRight).Address)
    Dim SignalNameRangeSys As Range: Set SignalNameRangeSys = Range(HeaderRangeSys.Find("Signal Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeaderRangeSys.Find("Signal Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    SignalNameRangeSys.Activate
    Dim ReactionNameRangeSys As Range: Set ReactionNameRangeSys = Range(HeaderRangeSys.Find(reaction, LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeaderRangeSys.Find(reaction, LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim j As Integer
    For j = 2 To SignalNameRangeSys.Cells.Count
        If ReactionNameRangeSys(j, 1) <> 0 Then
            signal = SignalNameRangeSys.Cells(j, 1).value
            value = ReactionNameRangeSys.Cells(j, 1).value
            Debug.Print (signal + " -> " + value)
            ReDim Preserve output(j - 2)
            output(j - 2) = signal 'TODO associative arrays vba?
            ReDim Preserve output(j - 1)
            output(j - 1) = value
        End If
    Next j

    getReactionFromTable = output
    Worksheets("TdR").Activate

End Function

Public Function getSignalFromTable(signal As String, Optional ByVal what As String) As String
    Dim output As String
    Worksheets("Signals").Activate
    Dim HeaderRangeMSRS As Range: Set HeaderRangeMSRS = Range(Range("SignalName").Address, Range("SignalName").End(xlDown))
    Dim SignalNameRangeMSRS As Range: Set SignalNameRangeMSRS = Range(HeaderRangeMSRS.Find("Signal Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeaderRangeMSRS.Find("Signal Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim j As Integer
    j = 2

    Debug.Print (SignalNameRangeMSRS.Cells(D, 1).value)

    output = SignalNameRangeMSRS.Cells(D, 1).value
    getSignalFromTable = output
    Worksheets("TdR").Activate

End Function

Public Function TestMissingFrameNegative(Channel As String, ECU As String, Frame As String, ConfirmationTime As String)
    'TODO for the moment just cutted from the main loop
    Debug.Print ("> NegativeCaseTime cutting frame " + FrameName + " for a period shorter than " + Str(ConfirmationTime) + "ms")
    If DTC <> "" Then
        Debug.Print ("  > Checking the DTC " + DTC + " did not raise")
    End If
    Debug.Print ("  > Checking the signals " + SignalReactionsBrut + " did not change")
    Debug.Print ("> PositiveCaseTime cutting frame " + FrameName + " for " + Str(ConfirmationTime) + " ms")
    Debug.Print ("  > Checking the DTC " + DTC + " raised")
    Debug.Print ("  > Checking the signals " + SignalReactionsBrut + " changed")
    Debug.Print ("?? find out how to compute all the positive cases and to iterate within them")
    Debug.Print ("> PositiveCaseConfig: cutting the frame after having configured " + FailureConfiguration)
    Debug.Print ("  > Checking the DTC " + DTC + " raised")
    Debug.Print ("  > Checking the signals " + SignalReactionsBrut + " changed")
    Debug.Print ("?? find out how to compute the negative cases and to iterate within them")
    Debug.Print ("> NegativeCaseConfig: cutting the frame after having configured in the wrong ways " + FailureConfiguration)
    Debug.Print ("  > Checking the DTC " + DTC + " did not raise")
    Debug.Print ("  > Checking the signals " + SignalReactionsBrut + " did not change")
End Function

Public Function getInfoTestLine(Line As Integer)
    'NOT BEING USED ANYMORE: WRITTEN DIRECTLY IN MAIN FOR NOT MAKING RANGES PUBLIC AND KEEP SIMPLE DIM + SET Ranges
    Debug.Print ("!! getInfoTestLine being used")
    Dim i As Integer
    Dim temp() As String

    'FailureType = FailureRangeD.Cells(D, 1).value 'Frame absent, CRC, Clock, Unavailable value, OutOfRange/NotUsed
    ChannelName = ChannelRangeD.Cells(D, 1)
    ECUName = ECUNameRangeD.Cells(D, 1)
    FrameName = FrameNameRangeD.Cells(D, 1)
    'TODO? IPDU manage difference CAN CANFD
    FrameID = FrameIDRangeD.Cells(D, 1)
    FramePeriod = FramePeriodRangeD.Cells(D, 1)
    SignalName = SignalNameRangeD.Cells(D, 1).value 'if whole frame, will be "Frame"
    Debug.Print ("-- ECU: " + ECUName + "; Frame : " + FrameName + "; Signal: " + SignalName + "; Failure: " + FailureType)
    ScriptNameSpecific = Str(Range("TDR_V").value) + "_" + ECUName + "_" + FrameName + "_" + SignalName + "_" + FailureType 'TODO if two line with same failure type - no case so far -  consider to do something here
    Debug.Print (" -> Script :" + ScriptNameSpecific)
    DTC = DTCRangeD.Cells(D, 1).value 'temp: for the moment, could be formatted as $XXXX-YY'
    Debug.Print ("> DTC : " + DTC)
    If InStr(DTC, "-") Then
        DTC = Left(DTCRangeD.Cells(D, 1).value, InStr(DTCRangeD.Cells(D, 1).value, "-") - 1)
        FaultType = Right(DTCRangeD.Cells(D, 1).value, Len(DTCRangeD.Cells(D, 1)) - InStr(DTCRangeD.Cells(D, 1).value, "-"))
        ' Debug.Print (DTC + " - " + FaultType)
    Else
        FaultType = Empty
        ' Debug.Print (DTC)
    End If
    diagActivationBrut = DiagActivationRangeD.Cells(D, 1).value

    configurationDIDBrut = ConfigRangeD.Cells(D, 1).value
    ConfigurationDIDListed = processConfigurationDID()

    'TODO Decide how to combine the two. For the moment, only configurationDID
    ConfirmationTime = ConfirmationTimeRangeD.Cells(D, 1).value
    DisappearenceTime = DisappearenceTimeRangeD.Cells(D, 1).value

    'TODO list the signals, expanding the system reactions if some'
    SignalReactionsBrut = SignalReactionRangeD.Cells(D, 1)
    SignalReactionListed = ProcessSignalReaction()

End Function

Public Function processConfigurationDID()
    Dim i As Integer
    Dim temp As String

    ConfigurationDIDListed = Split(configurationDIDBrut, vbLf)  'TODO check what it does when no vbLf'
    For i = 0 To UBound(ConfigurationDIDListed)
        Debug.Print (ConfigurationDIDListed(i))
        'TODO supposedly here, should be able to define negative conditions and positive conditions. in the form "TRUE: DID1 = X & DID2 = Y"
        'TODO for the moment, only positive
        ReDim Preserve testConfigArray(i)
        testConfigArray(i) = "TRUE: " + ConfigurationDIDListed(i)
        Debug.Print ("processed: " + testConfigArray(i))
    Next i

    FileOut.WriteLine (CanoeWriteComment("TODO write config; TODO true/false iteration"))

End Function

Public Function processDiagActivation()
  Dim i As Integer
  diagActivationListed = Split(diagActivationBrut, vbLf) 'TODO check what it does when no vbLf'
  For i = 0 To UBound(diagActivationListed)
      Debug.Print (diagActivationListed(i))
      'TODO supposedly here, should be able to define negative conditions and positive conditions. in the form "TRUE: signal1 = X And Signal2 = Y"
  Next i
End Function

Public Function createPositiveTestScript(Optional ByVal expResult As Boolean = True)
    Dim i As Integer
    Dim signal As String
    Dim value As String

    FileOut.WriteLine (CanoeWriteComment("Test failure :" + FailureType))

    Select Case DTC
        Case "", "_", "None"
            FileOut.WriteLine (CanoeReadDTC()) 'just read. warning if some DTC is present, and has not been indicated to be ingored
            FileOut.WriteLine (CanoeWriteComment("Warning if not masked DTC is already present"))
        Case Else
            FileOut.WriteLine (CanoeReadDTC(DTC, FaultType)) 'Warning if already present
            FileOut.WriteLine (CanoeWriteComment("Warning if already present"))
    End Select

    Select Case FailureType
        Case "Missing Frame"
            'Negative confirmation time case
            If expResult = True Then
                FileOut.WriteLine (CanoeCutFrame(ChannelName, ECUName, FrameName))
                FileOut.WriteLine (CanoeWriteComment("TODO test negativeConfirmationTime case"))
                FileOut.WriteLine (CanoeDelay(Str(ConfirmationTime / 2))) 'TODO check tollerancies
                FileOut.WriteLine (CanoeRestoreFrame(ChannelName, ECUName, FrameName))
                FileOut.WriteLine (CanoeReadDTC(DTC, FaultType, NOTPRESENT))
                FileOut.WriteLine (CanoeWriteComment("TODO Check signals negative?"))
            End If
            'Positive confirmation time case
            FileOut.WriteLine (CanoeCutFrame(ChannelName, ECUName, FrameName))
            FileOut.WriteLine (CanoeDelay(Str(ConfirmationTime)))
            Call scanSignalReactionListed
            FileOut.WriteLine (CanoeReadDTC(DTC, FaultType, PRESENT)) 'TODO? multiple DTCs?
            'Negative disappearence time case
            If expResult = True Then
                FileOut.WriteLine (CanoeRestoreFrame(ChannelName, ECUName, FrameName))
                FileOut.WriteLine (CanoeDelay(Str(ConfirmationTime / 2))) 'TODO check tollerancies
                FileOut.WriteLine (CanoeReadDTC(DTC, FaultType, PRESENT))
                Call scanSignalReactionListed
                FileOut.WriteLine (CanoeCutFrame(ChannelName, ECUName, FrameName))
                FileOut.WriteLine (CanoeDelay(Str(ConfirmationTime)))
                FileOut.WriteLine (CanoeRestoreFrame(ChannelName, ECUName, FrameName))
                FileOut.WriteLine (CanoeDelay(Str(DisappearenceTime)))
                Call scanSignalReactionListed
                FileOut.WriteLine (CanoeReadDTC(DTC, FaultType, MEMORISED))
            End If
        Case "CRC"
            FileOut.WriteLine (CanoeWriteComment("TODO function wrongCRC"))

        Case "Clock"
            FileOut.WriteLine (CanoeWriteComment("TODO function freeze clock"))

        Case "Unavailable"
            FileOut.WriteLine (CanoeWriteComment("TODO check situation before test"))
            FileOut.WriteLine (CanoeWriteComment("Set signal to unavailable value as defined in Message Set"))
            FileOut.WriteLine (CanoeWriteSignalValue(SignalName, unavailableValue))
            FileOut.WriteLine (CanoeWriteComment("TODO test negativeConfirmationTime case"))
            FileOut.WriteLine (CanoeDelay(Str(ConfirmationTime)))
            FileOut.WriteLine (CanoeReadDTC(DTC, FaultType, PRESENT)) 'TODO? multiple DTCs?
            Call scanSignalReactionListed
            FileOut.WriteLine (CanoeWriteComment("Set signal back to acceptable value - take value before writing unavailable?"))
            FileOut.WriteLine (CanoeWriteSignalValue(SignalName, "TBD AcceptableValue"))
            FileOut.WriteLine (CanoeWriteComment("TODO test negativeDisappearenceTime case"))
            FileOut.WriteLine (CanoeDelay(Str(DisappearenceTime)))
            Call scanSignalReactionListed
            FileOut.WriteLine (CanoeReadDTC(DTC, FaultType, MEMORISED))

        Case "Out Of Range"
            FileOut.WriteLine (CanoeWriteComment("TODO test case out of range"))

        Case "Not Used" 'TODO separing not used and out of range for the moment, will be easier to know if list or numeric..?
            FileOut.WriteLine (CanoeWriteComment("TODO test case not used"))
    End Select

End Function

Public Function createNegativeTestScript()
  'TODO merge it in the positive. for the moment trying only the positive
End Function

Public Function ProcessSignalReaction()
    Dim D As Integer
    Dim A As Integer
    Dim temp() As String
    ReDim SignalReactionListed(0)
    'SignalReactionBrur in the format
    '$Signal1 = X
    'SYS_R1'
    'SignalReactionListed: Format of each line: $signal = value

    temp = Split(SignalReactionsBrut, vbLf)
    A = 0
    For D = 0 To UBound(temp)
        Debug.Print (temp(D))
        Debug.Print (Str(A))
        If InStr(temp(D), "=") = False Then 'It means it is a reaction set
            Worksheets("System_Reaction").Activate
            Dim HeaderRangeSys As Range: Set HeaderRangeSys = Range(Range("SysSignalName").Address, Range("SysSignalName").End(xlToRight).Address)
            Dim SignalNameRangeSys As Range: Set SignalNameRangeSys = Range(HeaderRangeSys.Find("Signal Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeaderRangeSys.Find("Signal Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
            SignalNameRangeSys.Activate
            Dim ReactionNameRangeSys As Range: Set ReactionNameRangeSys = Range(HeaderRangeSys.Find(temp(D), LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeaderRangeSys.Find(temp(D), LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
            Dim jthCell As Integer
            Dim jthSignal As Integer
            jthSignal = 0
            Dim value As String
            Dim signal As String
            For jthCell = 2 To SignalNameRangeSys.Cells.Count
                If ReactionNameRangeSys(jthCell, 1) <> 0 Then
                    signal = SignalNameRangeSys.Cells(jthCell, 1).value
                    value = ReactionNameRangeSys.Cells(jthCell, 1).value
                    ReDim Preserve SignalReactionListed(A + jthSignal)
                    SignalReactionListed(A + jthSignal) = "$" + signal + " = " + value 'TODO associative arrays vba?
                    Debug.Print ("SignalReactionList(" + Str(A + jthSignal) + "): " + SignalReactionListed(A + jthSignal))
                    A = A + 1
                    jthSignal = jthSignal + 1
                End If
            Next jthCell
        Else '&Signal = value
            ReDim Preserve SignalReactionListed(A)
            SignalReactionListed(A) = temp(D)
            Debug.Print ("SignalReactionList(" + Str(A) + "): " + SignalReactionListed(A))
            A = A + 1
        End If

  Next D

  Debug.Print ("Listed items: " + Str(UBound(SignalReactionListed)))

    Worksheets("TdR").Activate
End Function

Public Function CanoeEndScript(numberOfParenthesis As Integer)
    Dim i As Integer

    For i = 1 To numberOfParenthesis
        FileOut.WriteLine ("}")
    Next i

End Function

Public Function ProcessConfigListed()
    Dim i As Integer
    FileOut.WriteLine (CanoeWriteComment("TODO process config TRUE DID1 x DID2 y"))
    'TODO testConfigArray
End Function

Public Function scanSignalReactionListed()
    Dim i As Integer
    Dim signal As String
    Dim value As String
    For i = 0 To UBound(SignalReactionListed)
        If SignalReactionListed(i) <> "" Then
            Debug.Print (SignalReactionListed(i))
            signal = Left(SignalReactionListed(i), InStr(SignalReactionListed(i), "=") - 1)
            Debug.Print (signal)
            value = Right(SignalReactionListed(i), Len(SignalReactionListed(i)) - InStr(SignalReactionListed(i), " = ") - 2)
            Debug.Print (value)
            If InStr(value, ":") <> 0 Then
                FileOut.WriteLine (CanoeReadSignalValue(signal, Left(value, InStr(value, ":") - 1))) 'OK if matching, NOK if not
            Else
                FileOut.WriteLine (CanoeReadSignalValue(signal, value)) 'OK if matching, NOK if not
            End If
        End If
    Next i
End Function
