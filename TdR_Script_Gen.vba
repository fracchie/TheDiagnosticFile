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

Public IndentIndex As Integer
Public FailureType As String
Public ConfirmationTime As Double
Public DisappearenceTime As Double
Public A As Integer
Public D As Integer
Public TotalConfig As String
Public DTC As String
Public FaultType As String
Public FailureConfiguration As String
Public DiagActivation As String
Public SignalReactions As String
Public SysReIDs As String
Public RequirementID As String

Public Sheet As Worksheet

Public FileOut As TextStream
Public FrameArxmlAddress As String 'TODO link between this message set and arxml file variables -> i.e. something like IL_CAN1::NODES::N_GWBridge::MESSAGES::framename

Const PRESENT As String = "PRESENT"
Const MEMORISED As String = "MEMORISED"
Const LESSTHAN As String = "LESSTHAN"
Const MORETHAN As String = "MORETHAN"

'Clean immediate windows
'Application.SendKeys "^g ^a {DEL}"

Sub Reactivity_Table_Script_gen()

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

    Workbooks("TheDiagnosticFile_V11error").Activate 'use that when debugging with several workbooks open, to avoid any bullshit. comment/remove it when not debugging anymore
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

'================ .xml file creation ==================
    Dim filePath As String
    Dim fileName As String
    Dim objShell As Object, objFolder As Object, objFolderItem As Object
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(&H0&, "Choose file's path", &H1&)
    Set objFolderItem = objFolder.Items.Item
    filePath = objFolderItem.Path
    'TODO input file name at same time of browsing address
    'fileName = "PVal_DID.xml"
    fileName = Range("ScriptName").value 'Cells C2 contains the name of the script. The reference of this cell has been changed in "ScriptName"
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
'================ .xml file creation ==================

    Dim ServiceRangeA As Range: Set ServiceRangeA = Range("B2", Range("B2").End(xlDown).Address)
    Dim CommandSentRangeA As Range: Set CommandSentRangeA = Range("I2", Range("I2").End(xlDown).Address)
    Dim ResponseExpectedRangeA As Range: Set ResponseExpectedRangeA = Range("J2", Range("J2").End(xlDown).Address)

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
    FileOut.WriteLine ("========================= Script Canoe =============================")
    FileOut.WriteLine ("to be copied in CAPL editor and integrated with missing informations")
    FileOut.WriteBlankLines (3)

    FileOut.Write ("Testcase ")
    FileOut.Write (fileName)
    FileOut.WriteLine ("()")
    FileOut.WriteLine ("{")
    FileOut.WriteBlankLines (2)
    IndentIndex = 1

    'Init and start logging
    InitTestCase

'---- Start the loop
    For D = 2 To FailureRangeD.Cells.Count

    Call getInfoTestLine(D)

    'New TestLine. collect info

        If ChannelRangeD.Cells(D, 1) <> ChannelRangeD.Cells(D - 1, 1) Then
            ChannelName = ChannelRangeD.Cells(D, 1)
        End If

        If (ECUNameRangeD.Cells(D, 1) <> "" And ECUNameRangeD.Cells(D, 1) <> ECUNameRangeD.Cells(D - 1, 1)) Then
            ECUName = ECUNameRangeD.Cells(D, 1)
            Debug.Print ("== ECU: " + ECUName)
            'FileOut.WriteLine ("// ECU : " + ECUName)
        End If

        If (FrameNameRangeD.Cells(D, 1).value <> "" And FrameNameRangeD.Cells(D, 1) <> FrameNameRangeD.Cells(D - 1, 1)) Then
            FrameName = FrameNameRangeD.Cells(D, 1)
            Debug.Print ("-- Frame: " + FrameName)
            FrameID = FrameIDRangeD.Cells(D, 1)
            FramePeriod = FramePeriodRangeD.Cells(D, 1)
            'FileOut.WriteLine ("// Frame: " + FrameName)
        End If

        'TODO IPDU

        If FailureRangeD.Cells(D, 1) <> "" And FailureRangeD.Cells(D, 1) <> "NA" And FailureRangeD.Cells(D, 1) <> "." And FailureRangeD.Cells(D, 1) <> "tbd" Then 'Valid validation line

            FailureType = FailureRangeD.Cells(D, 1).value 'Frame absent, CRC, Clock, Unavailable value, OutOfRange/NotUsed
            SignalName = SignalNameRangeD.Cells(D, 1).value 'if whole frame, will be "Frame"

            'Script name
            ScriptNameSpecific = Str(Range("TDR_V").value) + "_" + ECUName + "_" + FrameName + "_" + SignalName + "_" + FailureType 'TODO if two line with same failure type - no case so far -  consider to do something here
            Debug.Print ("------------------------------------------")
            Debug.Print ("Script :" + ScriptNameSpecific)
            FileOut.WriteLine ("// Script: " + ScriptNameSpecific)

            'DTC pre read
            Debug.Print ("> Read DTC before, and some interesting signal")
            DTC = DTCRangeD.Cells(D, 1).value
            If InStr(DTC, "-") Then
                DTC = Left(DTCRangeD.Cells(D, 1).value, InStr(DTCRangeD.Cells(D, 1).value, "-") - 1)
                FaultType = Right(DTCRangeD.Cells(D, 1).value, Len(DTCRangeD.Cells(D, 1)) - InStr(DTCRangeD.Cells(D, 1).value, "-"))
                Debug.Print (DTC + " - " + FaultType)
            Else
                FaultType = Empty
                Debug.Print (DTC)
            End If
            FileOut.WriteLine (CanoeReadDTC(DTC, FaultType))

            'Take Config
            'Diag activation
            DiagActivation = DiagActivationRangeD.Cells(D, 1).value
            Dim DiagActivationListed() As String
            DiagActivationListed = Split(DiagActivation, vbLf)
            For i = 0 To UBound(DiagActivationListed)
                Debug.Print (DiagActivationListed(i))
            Next i
            'Failure config
            Debug.Print ("> Checking configurations")
            FailureConfiguration = ConfigRangeD.Cells(D, 1).value
            Dim ConfigurationListed() As String
            ConfigurationListed = Split(FailureConfiguration, vbLf)
            For l = 0 To UBound(ConfigurationListed)
                Debug.Print (ConfigurationListed(l))
                'TODO supposedly here, should be able to define negative conditions and positive conditions
            Next l
            'TODO Dev DID

            'Reaction
            Debug.Print ("> checking reaction:")
            SignalReactions = SignalReactionRangeD.Cells(D, 1).value
            Dim SignalReactionListed() As String
            SignalReactionListed = Split(SignalReactions, vbLf)
            For l = 0 To UBound(SignalReactionListed)
                Debug.Print (SignalReactionListed(l))
            Next l

            ConfirmationTime = ConfirmationTimeRangeD.Cells(D, 1).value
            DisappearenceTime = DisappearenceTimeRangeD.Cells(D, 1).value

            'FileOut.WriteLine ("//------ Configuration: ")
            'ProcessConfiguration (ConfigRangeD.Cells(D, 1).Value)


            '============ Test cases
            'Test line execution
            Select Case FailureType 'Consider we are here only if failureType != <Check if conditions>
                Case "Missing frame"
                    'CanoeWriteComment ("Cause failure" + FailureType)
                    'cut the frame
                    Debug.Print ("------------------- Missing frame ----------------------------------")

                    'Positive case
                    Debug.Print ("..........Positive case 1..........")
                    'Write positive Config
                    Debug.Print ("> write config")
                    'Write positive check
                    FileOut.WriteLine (CanoeCutFrame(ChannelName, ECUName, FrameName))
                    FileOut.WriteLine (CanoeDelay(Str(ConfirmationTime)))
                    'FileOut.WriteLine ("Delay(" + Str(ConfirmationTime) + ");")
                    FileOut.WriteLine (CanoeReadDTC(DTC, FaultType, PRESENT))
                    Dim value As String
                    Dim signal As String
                    For i = 0 To UBound(SignalReactionListed)
                        If InStr(SignalReactionListed(i), "=") = False Then 'It means it is a reaction set
                            getReactionFromTable (SignalReactionListed(i))
                            'TODO go catching the list the table, if not done before
                        Else
                            signal = Left(SignalReactionListed(i), InStr(SignalReactionListed(i), "=") - 1)
                            value = Right(SignalReactionListed(i), Len(SignalReactionListed(i)) - InStr(SignalReactionListed(i), " = "))
                            FileOut.WriteLine (CanoeReadSignalValue(signal, value))
                        End If
                    Next i



                    'FileOut.WriteLine (GFL.CanoeCutFrame(ChannelName, ECUName, FrameName))

                    'FileOut.Write ("cutFrame(")
                    'FileOut.Write (FrameName)
                    'FileOut.Write (",")
                    'FileOut.Write (Str(ConfirmationTime))
                    'FileOut.Write (");")
                    'CanoeWriteComment (CanoeCheckDTC(DTC)) 'TODO check in which format communicate DTC. with $, 4 digit, fault type?
                    'Debug.Print ("> restoring frame " + FrameName)
                    'FileOut.WriteLine (GFL.CanoeRestoreFrame(ChannelName, ECUName, FrameName))
                Case "CRC"
                    Debug.Print ("> block CRC of " + FrameName + " for " + Str(ConfirmationTime) + " ms")
                    'CanoeWriteComment ("Cause failure" + FailureType)
                    'write bad CRC, frozen
                    'check
                    'put CRC automatic again
                Case "Clock"
                    Debug.Print ("> block Clock of " + FrameName + " for " + Str(ConfirmationTime) + " ms")
                    'CanoeWriteComment ("Cause failure" + FailureType)
                    'Froze the clock
                    'check
                    'put clock back to automatic
                Case "Unavailable"
                    CanoeWriteComment ("Cause failure" + FailureType)
                    'compute unavailable value for signal X and write
                    Debug.Print ("> writing unavailable value in " + FrameName)
                    'check
                    Debug.Print ("> Checking reaction to " + FailureType)
                    If DTC <> "" Then
                        Debug.Print (" DTC : " + DTC)
                        CanoeWriteComment (CanoeReadDTC(DTC)) 'TODO check in which format communicate DTC. with $, 4 digit, fault type?
                    End If
                    'put signal back to good value
                Case "OutOfRange/NotUsed"
                    CanoeWriteComment ("Cause failure" + FailureType)
                    'compute outofrange value for signal x and write it
                    Debug.Print ("> writing OutOfRange value in " + FrameName)
                    'check
                    'put signal back to good value
            End Select

            CanoeWriteComment (" cause failure")
            If DTCRangeD.Cells(D, 1) <> 0 Then
                'CutFrame (FrameName)
                FileOut.WriteLine (GFL.CanoeCutFrame(ChannelName, ECUName, FrameName))
            End If

        End If

        ScriptRangeD.Cells(D, 1).value = ScriptNameSpecific

        FileOut.WriteBlankLines (1)
        FileOut.WriteLine ("}")
        'Close file, otherwise it will leave the reference to the file, and will not allow you to re-write this file if launching another macro with the same file name
        FileOut.Close
        Set MyFSO = Nothing
        Set FileOut = Nothing
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

Public Function InitTestCase()

    'Note: testCaseDescription useless?
    FileOut.Write ("TestCaseDescription (")
    FileOut.Write (Chr(34)) 'quote mark "
    FileOut.Write (Chr(34)) 'quote mark "
    FileOut.Write (");")

    FileOut.WriteBlankLines (2)

    'Logging start
    FileOut.Write ("TestStepBegin(")
    FileOut.Write (Chr(34)) 'quote mark "
    FileOut.Write (Chr(34)) 'quote mark "
    FileOut.Write (",")
    FileOut.Write (Chr(34)) 'quote mark "
    FileOut.Write (Chr(34)) 'quote mark "
    FileOut.Write (");") 'TODO check how to start logging -> Ask



    FileOut.WriteBlankLines (2)

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

Function CanoeWriteComment(CellContent As String)
    FileOut.WriteLine ("// " + CellContent)
    FileOut.WriteBlankLines (2)
End Function

Function ProcessConfiguration(CellContent As String)
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
    Debug.Print ("  > Checking the signals " + SignalReactions + " did not change")
    Debug.Print ("> PositiveCaseTime cutting frame " + FrameName + " for " + Str(ConfirmationTime) + " ms")
    Debug.Print ("  > Checking the DTC " + DTC + " raised")
    Debug.Print ("  > Checking the signals " + SignalReactions + " changed")
    Debug.Print ("?? find out how to compute all the positive cases and to iterate within them")
    Debug.Print ("> PositiveCaseConfig: cutting the frame after having configured " + FailureConfiguration)
    Debug.Print ("  > Checking the DTC " + DTC + " raised")
    Debug.Print ("  > Checking the signals " + SignalReactions + " changed")
    Debug.Print ("?? find out how to compute the negative cases and to iterate within them")
    Debug.Print ("> NegativeCaseConfig: cutting the frame after having configured in the wrong ways " + FailureConfiguration)
    Debug.Print ("  > Checking the DTC " + DTC + " did not raise")
    Debug.Print ("  > Checking the signals " + SignalReactions + " did not change")
End Function

Public Function getInfoTestLine(Line As Integer)

    FailureType = FailureRangeD.Cells(D, 1).value 'Frame absent, CRC, Clock, Unavailable value, OutOfRange/NotUsed
    SignalName = SignalNameRangeD.Cells(D, 1).value 'if whole frame, will be "Frame"

End Function
