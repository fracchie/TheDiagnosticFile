' ===========================
' the whole script here is generated supposing that, on CANoe side, functions have been created such as
'ReadDID(DIDNumber as String, Optional ByVal ExpectedValueDec as Double)
'ReadParameter(DIDNumber As String, ParameterName As String, Optional ByVal ExpectedValueDec as Double)
'WriteDID(DIDNumber as String, DIDContent as String)
'WriteParameter(DIDNumber as String, ParameterName As String, ValueDec As Double)

'The macro is generating a set of scripts that, for each line of missing frame cases, compares the actual configuration of the ECU (read by CAPL readDID function)
' with the configuration defined in the reativity table, and if the result is true, checks that the defined DTC is reaised
'otherwise checks that the DTC IS NOT RAISED'

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
Public DIDName As String
Public ParamName As String
Public DIDNumber As String

Public IndentIndex As Integer
Public FailureType As String
Public ConfirmationTime As Double
Public DisappearenceTime As Double
Public TotalConfig As String
Public DTC As String
Public FaultType As String
Public configurationDIDList() As String
Public diagActivationList() As String
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

Sub tdr_line_validation_script_gen()

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

    ' Workbooks("TheDiagnosticFile_V11error").Activate 'use that when debugging with several workbooks open, to avoid any bullshit. comment/remove it when not debugging anymore
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
    Dim ConfirmationTimeRangeD As Range: Set ConfirmationTimeRangeD = Range(HeadersRangeD.Find("Confirmation Time (ms)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Confirmation Time (ms)").End(xlDown))
    Dim DisappearenceTimeRangeD As Range: Set DisappearenceTimeRangeD = Range(HeadersRangeD.Find("Disappearence Time (ms)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Disappearence Time (ms)").End(xlDown))
    Dim DTCRangeD As Range: Set DTCRangeD = Range(HeadersRangeD.Find("DTC Code", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("DTC Code").End(xlDown))
    Dim SignalReactionRangeD As Range: Set SignalReactionRangeD = Range(HeadersRangeD.Find("Signal Reaction", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Signal Reaction").End(xlDown))
    Dim ScriptRangeD As Range: Set ScriptRangeD = Range(HeadersRangeD.Find("Script", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Script").End(xlDown))

    Dim temp As String
    Dim ScriptNameSpecific As String
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

'---- Start the loop
    For D = 2 To FailureRangeD.Cells.Count
        FailureType = FailureRangeD.Cells(D, 1).value
        If FailureType = "Missing Frame" Or FailureType = "Unavailable" Or FailureType = "Out Of Range" Then
            ChannelName = ChannelRangeD.Cells(D, 1)
            ECUName = ECUNameRangeD.Cells(D, 1)
            FrameName = FrameNameRangeD.Cells(D, 1)
            FrameID = FrameIDRangeD.Cells(D, 1)
            FramePeriod = FramePeriodRangeD.Cells(D, 1)
            SignalName = SignalNameRangeD.Cells(D, 1).value 'if whole frame, will be "Frame"
            unavailableValue = UnavailableValueRangeD.Cells(D, 1).value
            DTC = DTCRangeD.Cells(D, 1).value 'temp: for the moment, could be formatted as $XXXX-YY'
            If InStr(DTC, "-") Then 'in case DTC is written in the format DTCcode - FaultType'
                DTC = Left(DTCRangeD.Cells(D, 1).value, InStr(DTCRangeD.Cells(D, 1).value, "-") - 1)
                FaultType = Right(DTCRangeD.Cells(D, 1).value, Len(DTCRangeD.Cells(D, 1)) - InStr(DTCRangeD.Cells(D, 1).value, "-"))
            Else
                FaultType = Empty
            End If
            ConfirmationTime = ConfirmationTimeRangeD.Cells(D, 1).value
            DisappearenceTime = DisappearenceTimeRangeD.Cells(D, 1).value
            diagActivationList = Split(DiagActivationRangeD.Cells(D, 1).value, vbLf)
            configurationDIDList = Split(ConfigRangeD.Cells(D, 1).value, vbLf)  'Written as OK param1 = 1 & Param2 = 1 vbLF NOK param1 = 0 & Param2 = 1 vbLF NOK param1 = 1 & Param2 = 0 NOK param1 = 0 & Param2 = 0. in same line just &. code will loop into lines of these configuration array of strings

            Debug.Print ("---------------------------------------------------------------------")
            Debug.Print ("-- ECU: " + ECUName + "; Frame : " + FrameName + "; Signal: " + SignalName + "; Failure: " + FailureType)

            '============ xml file creation =============
            ScriptNameSpecific = Range("TDR_V").value + "_" + ECUName + "_" + FrameName + "_" + SignalName + "_" + FailureType 'TODO if two line with same failure type - no case so far -  consider to do something here
            ScriptNameSpecific = "ciao"
            fileName = ScriptNameSpecific
            Dim MyFSO As New FileSystemObject
            If MyFSO.FolderExists(filePath) Then
                'MsgBox "The Folder already exists"
            Else
                MyFSO.CreateFolder (filePath) '<- Here the
            End If
            'Dim FileOut As TextStream 'Declared as public to be used out of this scope
            Set FileOut = MyFSO.CreateTextFile(filePath + "\" + fileName + ".xml", True, True)
            Call CanoeInitTestScript(ScriptNameSpecific)
            '=========== end of xml file creation =============

            'CANoeReadDTCs ' check if there are DTC before the test
            'CANcheckSignals ' check relevant signals before the test. suppose list of signals and value defined as OK and NOK in the diagnostic file

            'loop in the configuration list, and for each line perform a OK or NOK test
            Dim i As Integer
            For i = 0 To UBound(configurationDIDList)
                Debug.Print (configurationDIDList(i))
                'TODO boolean if config OK or NOK instead of If Left(configurationDIDList(i), 1) = "O" Then 'starts with O -> OK
                Select Case FailureType
                    Case "Missing Frame"
                        FileOut.WriteLine (CanoeCutFrame(ChannelName, ECUName, FrameName))
                        FileOut.WriteLine (CanoeDelay(Str(ConfirmationTime)))
                        If Left(configurationDIDList(i), 1) = "O" Then 'starts with O -> OK
                            'FileOut.WriteLine (CanoeWriteComment("if finalBoolean == True{"))
                            FileOut.WriteLine (CanoeReadDTC(DTC, FaultType, "PRESENT"))
                            'FileOut.WriteLine (CanoeWriteComment("}"))
                        Else 'it starts with N -> NOK
                             'FileOut.WriteLine (CanoeWriteComment("else {"))
                            FileOut.WriteLine (CanoeReadDTC(DTC, FaultType, "NOTPRESENT"))
                            'FileOut.WriteLine (CanoeWriteComment("}"))
                        End If
                        FileOut.WriteLine (CanoeRestoreFrame(ChannelName, ECUName, FrameName))
                    Case Else

                End Select
            Next i
            Call CanoeEndScript(2) 'CHECK number of paenthesis needed to close the script
            NumberOfFileOut = NumberOfFileOut + 1
            'Close file, otherwise it will leave the reference to the file, and will not allow you to re-write this file if launching another macro with the same file name
            FileOut.Close
            Set MyFSO = Nothing
            Set FileOut = Nothing
        Else ' Failure Type = Empty Or . Or NA
            ScriptNameSpecific = "."
        End If
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
    FileOut.WriteLine ("includes{")
    FileOut.WriteLine ("     #include " + Chr(34) + Chr(34) + "Common_Function_Lib_Reactivity.cin" + Chr(34) + Chr(34))
    FileOut.WriteLine ("     #include " + Chr(34) + Chr(34) + "DTC_Check_With_GADE.cin" + Chr(34) + Chr(34))
    FileOut.WriteLine ("}")
    FileOut.WriteBlankLines (2)
    FileOut.WriteLine ("testCase(" + ScriptName + "){")
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
    Dim a As Integer
    Dim temp() As String
    ReDim SignalReactionListed(0)
    'SignalReactionBrur in the format
    '$Signal1 = X
    'SYS_R1'
    'SignalReactionListed: Format of each line: $signal = value
    temp = Split(SignalReactionsBrut, vbLf)
    a = 0
    For D = 0 To UBound(temp)
        Debug.Print (temp(D))
        Debug.Print (Str(a))
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
                    ReDim Preserve SignalReactionListed(a + jthSignal)
                    SignalReactionListed(a + jthSignal) = "$" + signal + " = " + value 'TODO associative arrays vba?
                    Debug.Print ("SignalReactionList(" + Str(a + jthSignal) + "): " + SignalReactionListed(a + jthSignal))
                    a = a + 1
                    jthSignal = jthSignal + 1
                End If
            Next jthCell
        Else '&Signal = value
            ReDim Preserve SignalReactionListed(a)
            SignalReactionListed(a) = temp(D)
            Debug.Print ("SignalReactionList(" + Str(a) + "): " + SignalReactionListed(a))
            a = a + 1
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

Public Function ProcessConfigListedCases() As String()
    Dim i As Integer
    Dim temp() As String
    temp = Split(configurationDIDBrut, vbLf)


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

Public Function maskDIDbin(parameterName As String, value As Integer, Optional ByVal mask As String = "X") As String
    ' only for list parameter. Parameter with onoly the name of the param, no DIDname, and value as int
    Dim temp As String
    Dim i As Integer
    Dim Dt As Integer
    Dim startByte As Integer
    Dim size As Integer
    Dim bitOff As Integer
    Dim length As Integer
    Dim Coding As String

    Worksheets("Parameters").Activate
    Dim HeaderRangeParam As Range: Set HeaderRangeParam = Range(Range("Name").Address, Range("Name").End(xlToRight).Address)
    Dim NameRangeParam As Range: Set NameRangeParam = Range(HeaderRangeParam.Find("Parameter_Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeaderRangeParam.Find("Parameter_Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim DIDNumberRangeParam As Range: Set DIDNumberRangeParam = Range(HeaderRangeParam.Find("DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeaderRangeParam.Find("DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim StartByteRangeParam As Range: Set StartByteRangeParam = Range(HeaderRangeParam.Find("Start Byte", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeaderRangeParam.Find("Start Byte", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim SizeRangeParam As Range: Set SizeRangeParam = Range(HeaderRangeParam.Find("Size (bit)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeaderRangeParam.Find("Size (bit)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim BitOffsetRangeParam As Range: Set BitOffsetRangeParam = Range(HeaderRangeParam.Find("Bit offset", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeaderRangeParam.Find("Bit offset", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim LengthRangeParam As Range: Set LengthRangeParam = Range(HeaderRangeParam.Find("Length (Byte)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeaderRangeParam.Find("Length (Byte)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dim CodingRangeParam As Range: Set CodingRangeParam = Range(HeaderRangeParam.Find("Coding", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeaderRangeParam.Find("Coding", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Dt = NameRangeParam.Find(parameterName, LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Row - NameRangeParam.Find("Parameter_Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Row + 1

    DID = DIDNumberRangeParam.Cells(Dt, 1).value

    temp = ""
    For i = 1 To (CDbl(StartByteRangeParam.Cells(Dt, 1) - 1) * 8 + CDbl(BitOffsetRangeParam.Cells(Dt, 1)))
        temp = temp + mask
    Next i
    temp = temp + DecToBin(value, SizeRangeParam.Cells(Dt, 1))
    i = i + SizeRangeParam.Cells(Dt, 1)
    Do While i < (LengthRangeParam.Cells(Dt, 1) * 8)
        temp = temp + mask
        i = i + 1
    Loop
    maskDIDbin = temp
End Function

Public Function CanoeWriteComment(text As String) As String
    Dim temp As String

    temp = "// " + text
    CanoeWriteComment = temp

End Function
