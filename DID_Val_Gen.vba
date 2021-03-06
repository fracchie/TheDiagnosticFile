Option Explicit


' clean immediate window
'Application.SendKeys "^g ^a {DEL}"
' in immediate window, at the end of the macro it outputs this line of code in between two marker lines.
' place the cursor directly next to it and press ok to clean immediate window

'================ Rules ==============
' Data need to be sorted by DID first, then by Start byte and then for Bit offset. all smaller first
' The whole parameter tab needs to be sorted before launching the Macro
' Formats defined in the parameter tab must be met (spaces, dots, naming, Xs etc. )

'=============== The macro =============
' The macro reads values contained in parameters tab and threats them as Ranges.
' Then, iterating inside these ranges, actions are performed,
' to recognise different DIDs, their definition, and perform read and write operations
' in session 1 and 2 just reading it is allowed
' in session 3 also writing. first, the min value will be tested, then the max (checking also if DID can go out of range)
' then, if checked and true, an outofrange value is tested
' finally, default value (as defined in DID list) is written again in the DID

' PValXML is the main. For each session, DIDs are analysed and needed operations (act/deact via switches in param tab) are performed
' via DIDValStep, that is the function writing a line of the validation (a step) in the arrival sheet.
' Depending on the given arguments (i.e. rw=[WRITE, READ, CHECK], val=[MIN,MAX,DEF,OUTOFRANGE], res=[FORBIDDEN, OUTOFRANGE...]
' it uses the specific functions MinimumValue etc. to compose the binary message, to be translated in hex, to be sent

'============================================================================================================================================================================================================================================================================================
'      Global Variables Declaration
'============================================================================================================================================================================================================================================================================================

Public A As Integer
Public D As Integer

Public HeadersRangeD As Range
Public NameRangeD As Range
Public DIDRangeD As Range
Public LengthRangeD As Range
Public SizeRangeD As Range
Public WriteRangeD As Range
Public ReadRangeD As Range
Public SnapRangeD As Range
Public StartRangeD As Range
Public BitOffsetRangeD As Range
Public DefaultRangeD As Range
Public NumericRangeD As Range
Public ListRangeD As Range
Public MinRangeD As Range
Public MaxRangeD As Range
Public ResRangeD As Range
Public CodingRangeD As Range
Public SignRangeD As Range
Public OffsetRangeD As Range
Public IgnoreDefRangeD As Range
Public AsciRangeD As Range

Public Const FORBIDDEN = "FORBIDDEN"
Public Const OUTOFRANGE = "OUTOFRANGE"
Public Const MAX = "MAX"
Public Const MIN = "MIN"
Public Const NegRepr = 2

Public signed As Boolean

'Public DT As Integer
Public ServiceColA As Integer
Public SIDColA As Integer
Public IDColA As Integer
Public SessionColA As Integer
Public RequestColA As Integer
Public ResponseColA As Integer
Public HeadersRangeA As Range
Public list As String, value As String, Label As String
Public Color
Public session As Integer
Public Sheet As Worksheet
Public Cell As Range
Public WriteCheck As Boolean
Public i As Integer
Public Bit As Integer
Public IgnoreDef As Boolean
Public DIDNumber As String
Public DIDName As String

Public DIDdefValueBin As String
Public DIDLength As Integer

Public ButtonSession1 As Shape
Public ButtonSession2 As Shape
Public ButtonSession3 As Shape
Public ButtonRWSession1 As Shape
Public ButtonRWSession2 As Shape
Public ButtonRWSession3 As Shape
Public ButtonXML As Shape
Public ButtonReset As Shape


Sub PValXML()

'============================================================================================================================================================================================================================================================================================
'           Setup
'============================================================================================================================================================================================================================================================================================

'    Workbooks("PVal macro").Activate 'use that when debugging with several workbooks open
    ThisWorkbook.Activate
    '----------------------------------------------------------------------------------------------------
    'Variables declaration and init
    Dim temp As String
    '----------------------------------------------------------------------------------------------------

    Worksheets("Parameters").Activate

    'Set HeadersRangeD = Range("HeadersRangeD", Range("HeadersRangeD").End(xlToRight).Address)
    Set HeadersRangeD = Range("Name", Range("Name").End(xlToRight).Address)
    HeadersRangeD.Select
    'would like to format the whole thing as a tab, and maybe formatting the headers as text
    'Find the needed columns in the header list. By default is NOT CASE SENSITIVE
    Set NameRangeD = Range(HeadersRangeD.Find("Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Name", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set DIDRangeD = Range(HeadersRangeD.Find("DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set LengthRangeD = Range(HeadersRangeD.Find("Length (Byte)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Length (Byte)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set WriteRangeD = Range(HeadersRangeD.Find("Write by DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Write by DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set ReadRangeD = Range(HeadersRangeD.Find("Read by DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Read by DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set SizeRangeD = Range(HeadersRangeD.Find("Size (bit)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Size (bit)", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set DefaultRangeD = Range(HeadersRangeD.Find("Default Value", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Default Value", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set NumericRangeD = Range(HeadersRangeD.Find("Numeric", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Numeric", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set MinRangeD = Range(HeadersRangeD.Find("min", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("min", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set MaxRangeD = Range(HeadersRangeD.Find("max", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("max", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set ResRangeD = Range(HeadersRangeD.Find("resolution", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("resolution", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set SignRangeD = Range(HeadersRangeD.Find("sign", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("sign", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set OffsetRangeD = Range(HeadersRangeD.Find("Value offset", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Value offset", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set ListRangeD = Range(HeadersRangeD.Find("List", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("List", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set StartRangeD = Range(HeadersRangeD.Find("Start Byte", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Start Byte", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set BitOffsetRangeD = Range(HeadersRangeD.Find("Bit Offset", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Bit Offset", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set CodingRangeD = Range(HeadersRangeD.Find("Coding", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("Coding", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set IgnoreDefRangeD = Range(HeadersRangeD.Find("IgnoreDef DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("IgnoreDef DID", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))
    Set AsciRangeD = Range(HeadersRangeD.Find("ASCII|HEXA", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).Address, HeadersRangeD.Find("ASCII|HEXA", LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True).End(xlDown))

    Set ButtonSession1 = Worksheets("Parameters").Shapes("ButtonSession1")
    Set ButtonSession2 = Worksheets("Parameters").Shapes("ButtonSession2")
    Set ButtonSession3 = Worksheets("Parameters").Shapes("ButtonSession3")
    Set ButtonReset = Worksheets("Parameters").Shapes("ButtonReset")
    Set ButtonRWSession1 = Worksheets("Parameters").Shapes("ButtonRWSession1")
    Set ButtonRWSession2 = Worksheets("Parameters").Shapes("ButtonRWSession2")
    Set ButtonRWSession3 = Worksheets("Parameters").Shapes("ButtonRWSession3")
    Set ButtonXML = Worksheets("Parameters").Shapes("ButtonXML")

    '----------------------------------------------------------------------------------------------------
    '"Arrival" sheet : PVal
    '----------------------------------------------------------------------------------------------------

    Call CreateNewTab("PVal")
    'Define  the columns containing the data, by default
    A = 1
    ServiceColA = 2
    SIDColA = 3
    IDColA = 4
    SessionColA = 8
    RequestColA = 9
    ResponseColA = 10
    Set HeadersRangeA = Range(Cells(A, ServiceColA), Cells(A, ResponseColA))
    HeadersRangeA.Interior.Color = RGB(255, 192, 0)
    HeadersRangeA.RowHeight = 30
    HeadersRangeA.Font.Bold = 1
    HeadersRangeA.HorizontalAlignment = xlCenter
    HeadersRangeA.VerticalAlignment = xlCenter
    'Format:Borders
    HeadersRangeA.borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    HeadersRangeA.borders(xlEdgeLeft).Color = RGB(0, 0, 0)
    HeadersRangeA.borders(xlEdgeRight).Color = RGB(0, 0, 0)
    HeadersRangeA.borders(xlEdgeTop).Color = RGB(0, 0, 0)
    HeadersRangeA.borders(xlInsideVertical).Color = RGB(0, 0, 0)

    HeadersRangeA.WrapText = True
    Columns("A").HorizontalAlignment = xlCenter
    Columns("C:J").HorizontalAlignment = xlCenter
    Columns("A:J").NumberFormat = "@"

    '----------------------------------------------------------------------------------------
    'Headers
    '----------------------------------------------------------------------------------------
    A = 1
    Columns("A").ColumnWidth = 5
    Cells(A, ServiceColA).value = "Service name"
    Cells(A, SIDColA).value = "SID"
    Cells(A, IDColA).value = "2d Byte (not always a DID)"
    Cells(A, SessionColA).value = "Authorised sessions before sending the request for validation "
    Cells(A, RequestColA).value = "Request sent for validation"
    Cells(A, ResponseColA).value = "Positive response expected"


    'Format: Columns width

    Columns(ServiceColA).ColumnWidth = 100
    'Columns("A:J").NumberFormat = "@"
    Range(Columns(5), Columns(7)).ColumnWidth = 1
    Columns(SIDColA).ColumnWidth = 6
    Columns(IDColA).ColumnWidth = 7
    Columns(SessionColA).ColumnWidth = 7
    Columns(RequestColA).ColumnWidth = 40
    Columns(ResponseColA).ColumnWidth = 40
    'Format:interior color

    '================ .xml file declaration ==================
    If ButtonXML.TextFrame.Characters.text = "ON" Then
        Dim filePath As String
        Dim fileName As String
        Dim objShell As Object, objFolder As Object, objFolderItem As Object
        Set objShell = CreateObject("Shell.Application")
        Set objFolder = objShell.BrowseForFolder(&H0&, "Choose file's path", &H1&)
        Set objFolderItem = objFolder.Items.Item
        filePath = objFolderItem.Path
        Dim Aold As Integer: Aold = 2
        Dim TempByteSent As String
        Dim MyFSO As New FileSystemObject
        If MyFSO.FolderExists(filePath) Then
            'MsgBox "The Folder already exists"
        Else
            MyFSO.CreateFolder (filePath) '<- Here the
        End If
        Dim FileOut As TextStream
        '+ specific for this case -> Tab PVAL
        Dim ServiceRangeA As Range: Set ServiceRangeA = Range("B2", Range("B2").End(xlDown).Address)
        Dim CommandSentRangeA As Range: Set CommandSentRangeA = Range("I2", Range("I2").End(xlDown).Address)
        Dim ResponseExpectedRangeA As Range: Set ResponseExpectedRangeA = Range("J2", Range("J2").End(xlDown).Address)
        Dim numberOfFiles As Integer: numberOfFiles = 0
    End If
    '================ .xml file declaration ==================

    A = 2
    D = 2
    Dim j As Integer


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
    For session = 1 To 3

'====== Session 1
        If ((session = 1 And ButtonSession1.TextFrame.Characters.text = "ON") Or (session = 2 And ButtonSession2.TextFrame.Characters.text = "ON")) Then
            StartDiagSession (session)
            If ButtonReset.TextFrame2.TextRange.Characters.text = "ON" Then
                Call ResetECU
            End If

'========== And for each line of Departure tab, perform required actions for each DID found
            For D = 2 To NameRangeD.Cells.Count

'=============== New DID
                If DIDRangeD.Cells(D, 1) <> DIDRangeD.Cells(D - 1, 1) Then

                    '=== get DID name and number
                    DIDNumber = Right(DIDRangeD.Cells(D, 1).value, 4)
                    If (InStr(NameRangeD.Cells(D, 1), ".") <> 0) Then
                        DIDName = Left(NameRangeD.Cells(D, 1), InStr(NameRangeD.Cells(D, 1), ".") - 1)
                        Debug.Print ("-------- DID: " + DIDNumber + " - " + DIDName + "---------")
                    Else
                        DIDName = NameRangeD.Cells(D, 1)
                        Debug.Print ("-------- DID: " + DIDNumber + " - " + DIDName + "---------")
                    End If

                    If (ReadRangeD.Cells(D, 1).value = "X") Then 'some DIDs are not even RO, but only snapshots

                        Call DIDValStep("READ", "DEF") 'TODO const READ etc

                        '== 1st if RW, also a WRITE in session1 - expected to be forbidden
                        If ((session = 1 And ButtonRWSession1.TextFrame.Characters.text = "ON") Or (session = 2 And ButtonRWSession2.TextFrame.Characters.text = "ON")) Then
                            Call DIDValStep("WRITE", "DEF", "FORBIDDEN")
                        End If

                        If ButtonXML.TextFrame.Characters.text = "ON" Then
                            fileName = "DID_Val_" + DIDNumber + "_" + DIDName + ".xml"
                            Debug.Print ("creating file: " + fileName)
                            Set FileOut = MyFSO.CreateTextFile(filePath + "\" + fileName, True, True)
                            FileOut.WriteLine ("<Scenario description=" + Chr(34) + "DID Validation generated by DiagnosticFile " + Chr(34) + ">")
                            FileOut.WriteLine ("      <test name=" + Chr(34) + fileName + Chr(34) + " methode=" + Chr(34) + "2" + Chr(34) + " mode=" + Chr(34) + "PointToPoint" + Chr(34) + " >")
                            FileOut.WriteLine ("            <request name=" + Chr(34) + "StartDiagnosticSession" + Chr(34) + ">")
                            FileOut.WriteLine ("               <Check startbytes=" + Chr(34) + "5003" + Chr(34) + "/>")
                            FileOut.WriteLine ("               <byte min=" + Chr(34) + "10" + Chr(34) + "/>")
                            FileOut.WriteLine ("               <byte min=" + Chr(34) + "03" + Chr(34) + "/>")
                            FileOut.WriteLine ("            </request>")
                            Do While Aold < A - 1
                                FileOut.WriteLine ("            <request name=" + Chr(34) + ServiceRangeA.Cells(Aold, 1).value + Chr(34) + ">")

                                temp = ResponseExpectedRangeA.Cells(Aold, 1).value
                                If (InStr(temp, "ERROR") <> 0) Then
                                    temp = Right(Left(temp, 10), 4)
                                    FileOut.WriteLine ("               <Check codeErr=" + Chr(34) + temp + Chr(34) + "/>")
                                    'FileOut.WriteLine ("               <Check codeErr=" + Chr(34) + "2046" + Chr(34) + " codeErr2=" + Chr(34) + "2032" + Chr(34) + " CodeErr3=" + Chr(34) + "2048" + Chr(34) + " CodeErr4=" + Chr(34) + "2056" + Chr(34) + "/>")
                                Else
                                    FileOut.WriteLine ("               <Check startbytes=" + Chr(34) + ResponseExpectedRangeA.Cells(Aold, 1).value + Chr(34) + "/>")
                                    'FileOut.WriteLine ("               <Check startbytes=" + Chr(34) + ResponseExpectedRangeA.Cells(i, 1).Value + Chr(34) + "/>")
                                    'Did not get why does not take responseExpectedRange, it keeps saying invalide type. Last time it was. Tried to debug, at the end i worked around in that way
                                End If

                                TempByteSent = CommandSentRangeA.Cells(Aold, 1).value
                                For j = 1 To Len(TempByteSent) Step 2
                                    FileOut.WriteLine ("               <byte min=" + Chr(34) + Mid(TempByteSent, j, 2) + Chr(34) + "/>")
                                Next j

                                FileOut.WriteLine ("            </request>")
                                Aold = Aold + 1
                            Loop
                            FileOut.WriteLine ("     </test>")
                            FileOut.WriteLine ("</Scenario>")
                            'Close file, otherwise it will leave the reference to the file, and will not allow you to re-write this file if launching another macro with the same file name
                            numberOfFiles = numberOfFiles + 1
                            FileOut.Close
                            Set MyFSO = Nothing
                            Set FileOut = Nothing
                        End If

                    End If

                    '===============================
                    ' TODO Decide if, with RO setting, you want to try to write just once and expect forbidden response or not even that.
                    ' Put writeCheck = False at the beginning of the session for cycle
                    ' And repeat the code for other sessions. Check here below, should be the code was working before in other v
    '                If WriteCheck = False Then
    '                    Call DIDValStep("WRITE", "DEF", "Forbidden")
    '                    'Receive the negative response should be enough
    '                    WriteCheck = True
    '                End If
                End If

'
            Next D

'======= Execute session 3 if active
        ElseIf (session = 3 And ButtonSession3.TextFrame.Characters.text = "ON") Then
            StartDiagSession (session)
            If ButtonReset.TextFrame2.TextRange.Characters.text = "ON" Then
                Call ResetECU
            End If

'========== For each line of Departure tab, perform required actions for each DID found
            For D = 2 To NameRangeD.Cells.Count

'============== New DID
                If DIDRangeD.Cells(D, 1) <> DIDRangeD.Cells(D - 1, 1) Then

                    '== get DID name and number
                    DIDNumber = Right(DIDRangeD.Cells(D, 1).value, 4)
                    If (InStr(NameRangeD.Cells(D, 1), ".") <> 0) Then
                        DIDName = Left(NameRangeD.Cells(D, 1), InStr(NameRangeD.Cells(D, 1), ".") - 1)
                        Debug.Print ("-------- DID: " + DIDNumber + " - " + DIDName + "---------")
                    Else
                        DIDName = NameRangeD.Cells(D, 1)
                        Debug.Print ("-------- DID: " + DIDNumber + " - " + DIDName + "---------")
                    End If

                    DIDLength = LengthRangeD.Cells(D, 1).value

                    If ReadRangeD.Cells(D, 1).value = "X" Then
                        Call DIDValStep("READ", "DEF")
                    Else
                        'just snapshot, what to do with them?
                    End If

                    ' 2nd: Execute Write operation if active. If possible, try to write min, max, OutOfRange value,checking everytime
                    If ButtonRWSession3.TextFrame.Characters.text = "ON" Then

                        'writable DID. Try writing min, max, out of range(flag set in when comuting max) and then def again
                        If WriteRangeD.Cells(D, 1).value = "X" Then

                            Call DIDValStep("WRITE", "DEF")
                            'TODO check if the new MinMaxValueLoop works as well as the other, but more detailed
                            'Call DIDValStep("WRITE", "MIN")
                            'Call DIDValStep("CHECK", "MIN")
                            Call MinMaxValueLoop(DIDdefValueBin, "MIN")

                            Call OutOfRangeLoop(DIDdefValueBin, "DOWN")

                            'TODO write outOfRange min

                            ' Write Max, setting flag Outofrange if needed (<- from DIIValStep WRITE MAX))
                            'TODO check if the new MinMaxValueLoop works as well as the other, but more detailed
                            'Call DIDValStep("WRITE", "MAX")
                            'Call DIDValStep("CHECK", "MAX")
                            Call MinMaxValueLoop(DIDdefValueBin, "MAX")
                            ' If can go out of range (<- Public OutOrRange), test it, writing max +1
                            Call OutOfRangeLoop(DIDdefValueBin, "UP")

                            'Call DIDValStep("WRITE", "OUTOFRANGEUP", "OUTOFRANGE")
                            'Call DIDValStep("CHECK", "OUTOFRANGEUP", "IGNORE")

                            Call ListValueLoop(DIDdefValueBin)
                            Call OutOfRangeLoop(DIDdefValueBin, "LIST")

                            Call DIDValStep("WRITE", "DEF")
                            Call DIDValStep("CHECK", "DEF")

                            ' not writable DID, try writing and expect out of range - don't know why they don't use "subfunction not allowed" or something similar that exists
                        ElseIf ReadRangeD.Cells(D, 1).value = "X" Then
                            Call DIDValStep("WRITE", "DEF", "READONLY")
                        Else 'just a snapshot? do nothing
                        End If
                    End If
                    If ButtonXML.TextFrame.Characters.text = "ON" Then
                        fileName = "DID_Val_" + DIDNumber + "_" + DIDName + ".xml"
                        Debug.Print ("creating file: " + fileName)
                        Set FileOut = MyFSO.CreateTextFile(filePath + "\" + fileName, True, True)
                        FileOut.WriteLine ("<Scenario description=" + Chr(34) + "DID Validation generated by DiagnosticFile " + Chr(34) + ">")
                        FileOut.WriteLine ("      <test name=" + Chr(34) + fileName + Chr(34) + " methode=" + Chr(34) + "2" + Chr(34) + " mode=" + Chr(34) + "PointToPoint" + Chr(34) + " >")
                        FileOut.WriteLine ("            <request name=" + Chr(34) + "StartDiagnosticSession" + Chr(34) + ">")
                        FileOut.WriteLine ("               <Check startbytes=" + Chr(34) + "5003" + Chr(34) + "/>")
                        FileOut.WriteLine ("               <byte min=" + Chr(34) + "10" + Chr(34) + "/>")
                        FileOut.WriteLine ("               <byte min=" + Chr(34) + "03" + Chr(34) + "/>")
                        FileOut.WriteLine ("            </request>")
                        Do While Aold < A - 1
                            FileOut.WriteLine ("            <request name=" + Chr(34) + ServiceRangeA.Cells(Aold, 1).value + Chr(34) + ">")

                            temp = ResponseExpectedRangeA.Cells(Aold, 1).value
                            If (InStr(temp, "ERROR") <> 0) Then
                                temp = Right(Left(temp, 10), 4)
                                FileOut.WriteLine ("               <Check codeErr=" + Chr(34) + temp + Chr(34) + "/>")
                                'FileOut.WriteLine ("               <Check codeErr=" + Chr(34) + "2046" + Chr(34) + " codeErr2=" + Chr(34) + "2032" + Chr(34) + " CodeErr3=" + Chr(34) + "2048" + Chr(34) + " CodeErr4=" + Chr(34) + "2056" + Chr(34) + "/>")
                            Else
                                FileOut.WriteLine ("               <Check startbytes=" + Chr(34) + ResponseExpectedRangeA.Cells(Aold, 1).value + Chr(34) + "/>")
                                'FileOut.WriteLine ("               <Check startbytes=" + Chr(34) + ResponseExpectedRangeA.Cells(i, 1).Value + Chr(34) + "/>")
                                'Did not get why does not take responseExpectedRange, it keeps saying invalide type. Last time it was. Tried to debug, at the end i worked around in that way
                            End If

                            TempByteSent = CommandSentRangeA.Cells(Aold, 1).value
                            For j = 1 To Len(TempByteSent) Step 2
                                FileOut.WriteLine ("               <byte min=" + Chr(34) + Mid(TempByteSent, j, 2) + Chr(34) + "/>")
                            Next j

                            FileOut.WriteLine ("            </request>")
                            Aold = Aold + 1
                        Loop
                        FileOut.WriteLine ("     </test>")
                        FileOut.WriteLine ("</Scenario>")
                        'Close file, otherwise it will leave the reference to the file, and will not allow you to re-write this file if launching another macro with the same file name
                        numberOfFiles = numberOfFiles + 1
                        FileOut.Close
                        Set MyFSO = Nothing
                        Set FileOut = Nothing
                    End If
                End If
            Next D
        End If
    Next session

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

    Range("A2", "J" + CStr(A - 1)).Select

    If ButtonXML.TextFrame.Characters.text = "ON" Then MsgBox "created " & numberOfFiles & " files in " & filePath

    Debug.Print ("==================================================")
    Debug.Print ("Application.SendKeys " + Chr(34) + "^g ^a {DEL}")
    Debug.Print ("==================================================")

End Sub

Function ComputeContent(what As String, Optional ByVal returnAs As String = "Hexa") As String

    Dim bin As String
    Dim res As Double
    Dim off As Double
    Dim dec As Double
    Dim Dt As Double 'Temp starting from <-Public D, first data in new DID
    Dim size As Integer
    Dim list() As String
    Dim DataName As String
    Dim i As Integer

    bin = ""
    Bit = 0
    Dt = D
    DataName = ""

'=== loop into param of DID
'D starts from 2, because 1 is the header. The number of param is the count of cells in range -1, so that, starting from 2, DT = X means that we are at the X-1th parameter
    Do While ((Dt <= DIDRangeD.Cells.Count) And DIDRangeD.Cells(Dt, 1).value = ("$" + DIDNumber)) 'putting the condition on length first it avoids to go in overflow (out of table))

        DataName = NameRangeD.Cells(Dt, 1).value
        Debug.Print ("...." + DataName + "....")

        'pre-padding
        Do While Bit < ((StartRangeD.Cells(Dt, 1).value - 1) * 8) + BitOffsetRangeD.Cells(Dt, 1)
            Bit = Bit + 1
            bin = bin + "0"
        Loop

        ' Numeric ================================
        If NumericRangeD.Cells(Dt, 1).value <> 0 Then
            res = ResRangeD.Cells(Dt, 1).value
            off = OffsetRangeD.Cells(Dt, 1).value
            size = CInt(SizeRangeD.Cells(Dt, 1).value)
            'If SignRangeD.Cells(Dt, 1).value = "s" Then
            '    size = size - 1
            '    If (CDbl(dec) < 0) Then
            '        bin = bin + "1"
            '        dec = Str(Abs(CDbl(dec)))
            '    Else
            '        bin = bin + "0"
            '    End If
            'End If
            Select Case what
                Case "MIN"
                    dec = MinRangeD.Cells(Dt, 1).value
                Case "MAX"
                    dec = MaxRangeD.Cells(Dt, 1).value
                Case "DEF"
                    If InStr(DefaultRangeD.Cells(Dt, 1).value, " ") = 0 Then
                        dec = DefaultRangeD.Cells(Dt, 1)
                    Else
                        dec = Left(DefaultRangeD.Cells(Dt, 1).value, (InStr(DefaultRangeD.Cells(Dt, 1).value, " ")) - 1)
                    End If
                'Case "OUTOFRANGE"
                    'dec = CDbl(MaxRangeD.Cells(Dt, 1).value) + res
            End Select
            bin = bin + NewDecToBin(Str(dec), size, SignRangeD.Cells(Dt, 1) = "s", NegRepr, res, off)
            'bin = bin + DecToBin(Str(dec), size, , , res, off)
            '====================================================

        ' List ===========================================
        ElseIf ListRangeD.Cells(Dt, 1) <> 0 Then
            size = SizeRangeD.Cells(Dt, 1).value
            Dim value As String
            list = Split(CodingRangeD.Cells(Dt, 1), vbLf) 'dividing in lines of kind "0 = asd" ; "1 = asd" etc.

            'Commented the whole thing down here, because was doing min = first element of the list, max = last element of the list.
            ' but now decided new format: the whole list of values representable is listed, with x : Not Used for all the in valid values. min max nosense
            ' there is also a little macro to generate the list "x : Not Used" according to the size of the param. no excuses
            value = DefaultRangeD.Cells(Dt, 1).value
            If InStr(value, " ") <> 0 Then 'TODO replace : with " " to allow both : and = ?
                value = Left(value, InStr(value, " ") - 1)
            ElseIf value = "" Then
                value = "0"
            Else
                value = DefaultRangeD.Cells(Dt, 1).value
            End If

            dec = CDbl(value)
            bin = bin + NewDecToBin(value, size)

        ElseIf AsciRangeD.Cells(Dt, 1) <> 0 Then
            For i = 0 To size
                bin = bin + "0"
            Next i
        End If

        Debug.Print ("      -> dec: " + Str(dec) + " -> bin: " + bin)

        Bit = Bit + SizeRangeD.Cells(Dt, 1) 'TODO watchout if using negative number represtation msb, check if did size-1
        Dt = Dt + 1

    Loop

'    Final padding
    Do While Bit < (DIDLength * 8)

        bin = bin + "0"
        Bit = Bit + 1
    Loop

    If what = "DEF" Then DIDdefValueBin = bin

    If returnAs = "bin" Then
        ComputeContent = bin
    Else 'returnAs = "Hexa"
        ComputeContent = BinToHex(bin)
    End If

End Function

Function DIDValStepOutOfRange(parameter As String, UpDownNotUsed As String, request As String)
    'TODO specifi function for the out of range val step. when calling, a specific parameter is written out or range
    'up down to say if testing the lower or upper threshold, and notUsed for the lists
    'TODO first the upper threshold

End Function

Function DIDValStep(operation As String, what As String, Optional ByVal response As String = "", Optional ByVal valueDec As String = "")
    value = ""

    Debug.Print ("-------------- " + operation + " " + what + " --------------")

    Cells(A, 1).value = A - 1 'number of step: needed for script with pris diag -> scriptEditor
    Cells(A, ServiceColA).value = operation + " value " + what + " " + valueDec + " in " + DIDName 'Write the kind of operation to perform
    Cells(A, IDColA).value = DIDRangeD.Cells(D, 1).value
    Cells(A, SessionColA).value = "100" + CStr(session)

    'compose the content of the DID, either for the request or for the response. With NA in default cell we indicate the DID for which no specific def value is required, hence no check to be performed, just read the content
    If DefaultRangeD.Cells(D, 1) <> "NA" Then value = ComputeContent(what)

    Select Case operation
        Case "WRITE"
            'Write request
            Cells(A, RequestColA).value = "2E" + DIDNumber + value
            Cells(A, SIDColA).value = "$2E"
            'write response
            If response = "FORBIDDEN" Then
                Cells(A, ResponseColA).value = UDS_NACK_13_WrongFormat
                Range(Cells(A, ResponseColA).Address).HorizontalAlignment = xlLeft
            ElseIf response = "READONLY" Then
                Cells(A, ResponseColA).value = UDS_NACK_31_OutOfRange
                Range(Cells(A, ResponseColA).Address).HorizontalAlignment = xlLeft
            ElseIf response = "OUTOFRANGE" Then
                Cells(A, ResponseColA).value = UDS_NACK_31_OutOfRange
                Range(Cells(A, ResponseColA).Address).HorizontalAlignment = xlLeft
            Else
                Cells(A, ResponseColA).value = "6E" + DIDNumber
            End If

        Case Else 'Read or Check
            'Write request
            Cells(A, RequestColA).value = "22" + DIDNumber
            Cells(A, SIDColA).value = "$22"
            'Write response
            If (response = "IGNORE" Or response = "OUTOFRANGE") Then
                Cells(A, ResponseColA).value = ""
            Else
                Cells(A, ResponseColA).value = "62" + DIDNumber + value
            End If

    End Select

'==Next (-> Public A)
    A = A + 1

End Function

Function StartDiagSession(session As Integer)

        Cells(A, ServiceColA).value = "StartDiagnosticSession"
        Cells(A, SIDColA).value = "$10"
        Cells(A, SessionColA).value = "100" + CStr(session)
        Cells(A, IDColA).value = "$0" + CStr(session)
        If (ButtonXML.TextFrame.Characters.text = "OFF") Then
            'For scriptGenerator, no command for start diag session should be sent
        Else
            Cells(A, RequestColA).value = "100" + CStr(session)
        End If

        Cells(A, ResponseColA).value = "500" + CStr(session) + "003201F4"
        Cells(A, 1).value = A - 1
        A = A + 1

End Function

Function ResetECU()

            Cells(A, 1).value = A - 1
            Cells(A, ServiceColA).value = "ECUReset"
            Cells(A, SessionColA).value = "100" + CStr(session)
            Cells(A, RequestColA).value = "1101" 'verify the mechanism of the script
            A = A + 1

End Function

Function ChooseFolder() As String
'TODO generalise

    Dim fldr As FileDialog
    Dim sItem As String

    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With

NextCode:
    ChooseFolder = sItem
    Set fldr = Nothing
End Function
Public Function MinMaxValueLoop(DIDdefValueBin As String, MinMax As String)
    Dim ParamName As String
    Dim Dt As Integer
    Dim res As Double
    Dim off As Double
    Dim dec As Double
    Dim size As Integer
    Dt = D
    Dim inBin As String
    inBin = ""
    Dim out As String
    Dim bitOff As Integer
    Dim ByteStart As Integer

    Do While Right(DIDRangeD.Cells(Dt, 1).value, 4) = DIDNumber
        inBin = ""
        If NumericRangeD.Cells(Dt, 1) <> 0 Then
            ParamName = NameRangeD.Cells(Dt, 1).value
            size = CDbl(SizeRangeD.Cells(Dt, 1).value)
            bitOff = BitOffsetRangeD.Cells(Dt, 1).value
            ByteStart = StartRangeD.Cells(Dt, 1).value
            res = ResRangeD.Cells(Dt, 1).value
            'res = CDbl(ResRangeD.Cells(Dt, 1).value)
            off = OffsetRangeD.Cells(Dt, 1).value
            If (MinMax = "MAX") Then
                dec = MaxRangeD.Cells(Dt, 1).value
            Else 'UpDownList = "MIN"
                dec = MinRangeD.Cells(Dt, 1).value
            End If
            Cells(A, ServiceColA).value = "WRITE value " + MinMax + " " + Str(dec) + " in " + ParamName + ": size = " + Str(size) + "; res = " + Str(res) + "; offset = " + Str(off)


            inBin = inBin + NewDecToBin(Str(dec), size, SignRangeD.Cells(Dt, 1).value = "s", NegRepr, res, off)
            out = replaceInString(DIDdefValueBin, inBin, (ByteStart - 1) * 8 + bitOff)
            Cells(A, 1).value = A - 1
            Cells(A, SIDColA).value = "$2E"
            Cells(A, IDColA).value = "$" + DIDNumber
            Cells(A, ServiceColA).value = Cells(A, ServiceColA).value + " -> " + inBin

            Cells(A, SessionColA).value = "100" + CStr(session)
            Cells(A, RequestColA).value = "2E" + DIDNumber + BinToHex(out)
            Cells(A, ResponseColA).value = "6E" + DIDNumber
            A = A + 1
            Cells(A, 1).value = A - 1
            Cells(A, SIDColA).value = "$22"
            Cells(A, IDColA).value = "$" + DIDNumber
            Cells(A, ServiceColA).value = "CHECK value " + MinMax + " in " + ParamName
            Cells(A, SessionColA).value = "100" + CStr(session)
            Cells(A, RequestColA).value = "22" + DIDNumber
            Cells(A, ResponseColA).value = "62" + DIDNumber + BinToHex(out)
            A = A + 1
        End If
        Dt = Dt + 1
    Loop
End Function

Public Function OutOfRangeLoop(DIDdefValueBin As String, UpDownList As String)
'DIDdefValueBin is a reference good value used as mask for other parameters. will be the default value stored when first computing def for each DID, in public var
    Dim ParamName As String
    Dim Dt As Integer
    Dim res As Double
    Dim off As Double
    Dim dec As Double
    Dim size As Integer
    Dim lowLimitDec As Double
    Dim HighLimitDec As Double
    Dt = D
    Dim inBin As String
    Dim out As String
    Dim bitOff As Integer
    Dim ByteStart As Integer

    Do While Right(DIDRangeD.Cells(Dt, 1).value, 4) = DIDNumber
        inBin = ""
        ParamName = NameRangeD.Cells(Dt, 1).value
        'If InStr(paramName, ".") <> 0 Then
        '    paramName = Right(paramName, Len(paramName) - InStr(paramName, "."))
        'End If
        Select Case UpDownList
            Case "LIST"
                If ListRangeD.Cells(Dt, 1) <> 0 Then
                    Dim temp() As String
                    Dim i As Integer
                    Dim val As Double
                    size = SizeRangeD.Cells(Dt, 1).value
                    temp = Split(CodingRangeD.Cells(Dt, 1).value, vbLf)
                    For i = 0 To UBound(temp)
                        If InStr(temp(i), "Not Used") <> 0 Then
                            val = CDbl(Left(temp(i), InStr(temp(i), ":") - 1))

                            inBin = NewDecToBin(Str(val), size)
                            'Debug.Print (inBin)
                            out = replaceInString(DIDdefValueBin, inBin, (ByteStart - 1) * 8 + bitOff)
                            'Debug.Print (out)
                            Cells(A, 1).value = A - 1
                            Cells(A, SIDColA).value = "$2E"
                            Cells(A, IDColA).value = "$" + DIDNumber
                            Cells(A, SessionColA).value = "100" + CStr(session)
                            Cells(A, ServiceColA).value = "WRITE value NOTUSED " + Str(val) + " in " + ParamName + ": size = " + Str(size) + " -> " + inBin
                            Cells(A, RequestColA).value = "2E" + DIDNumber + BinToHex(out)
                            Cells(A, ResponseColA).value = UDS_NACK_31_OutOfRange
                            A = A + 1
                        End If
                    Next i
                End If
            Case Else ' "UP", "DOWN"
                If NumericRangeD.Cells(Dt, 1) <> 0 Then
                    size = CDbl(SizeRangeD.Cells(Dt, 1).value)
                    bitOff = BitOffsetRangeD.Cells(Dt, 1).value
                    ByteStart = StartRangeD.Cells(Dt, 1).value
                    res = CDbl(ResRangeD.Cells(Dt, 1).value)
                    off = CDbl(OffsetRangeD.Cells(Dt, 1).value)
                    ' Signed values, msb sign carry - TODO check if they use 2-complement instead 'TODO manage negative, decide if 2compl or msb
                    If (UpDownList = "UP") Then
                        dec = CDbl(MaxRangeD.Cells(Dt, 1).value)
                        If SignRangeD.Cells(Dt, 1).value = "s" Then
                            HighLimitDec = CDbl((2 ^ (size - 1) - 1 + off) * res)
                        Else
                            HighLimitDec = CDbl((2 ^ size - 1 + off) * res)
                        End If
                    Else 'UpDownList = "Down"
                        dec = CDbl(MinRangeD.Cells(Dt, 1).value)
                        If SignRangeD.Cells(Dt, 1).value = "s" Then
                            If NegRepr = 2 Then
                                lowLimitDec = CDbl((2 ^ (size - 1) + off) * res) * -1
                            Else
                                lowLimitDec = CDbl((2 ^ (size - 1) - 1 + off) * res) * -1
                            End If
                        Else
                            lowLimitDec = off
                        End If
                    End If


                    'If SignRangeD.Cells(Dt, 1).value = "s" Then
                    '    size = size - 1
                    '    If (dec < 0) Then
                    '        inBin = "1"
                    '        dec = -1 * dec 'abs
                    '    Else
                    '        inBin = "0"
                    '    End If
                    'End If
                    If UpDownList = "UP" Then
                        If dec < HighLimitDec Then 'Can go out of range
                            dec = dec + res
                            inBin = inBin + NewDecToBin(Str(dec), size, SignRangeD.Cells(Dt, 1).value = "s", NegRepr, res, off)
                            out = replaceInString(DIDdefValueBin, inBin, (ByteStart - 1) * 8 + bitOff)
                            Cells(A, 1).value = A - 1
                            Cells(A, SIDColA).value = "$2E"
                            Cells(A, IDColA).value = "$" + DIDNumber
                            Cells(A, SessionColA).value = "100" + CStr(session)
                            Cells(A, ServiceColA).value = "WRITE value OUTOFRANGE " + Str(dec) + " in " + ParamName + ": size = " + Str(size) + "; res = " + Str(res) + "; offset = " + Str(off) + " -> " + inBin
                            Cells(A, RequestColA).value = "2E" + DIDNumber + BinToHex(out)
                            Cells(A, ResponseColA).value = UDS_NACK_31_OutOfRange
                            A = A + 1
                        End If
                    Else 'DOWN
                        If dec > lowLimitDec Then
                            dec = dec - res
                            inBin = NewDecToBin(Str(dec), size, SignRangeD.Cells(Dt, 1).value = "s", NegRepr, res, off)
                            out = replaceInString(DIDdefValueBin, inBin, (ByteStart - 1) * 8 + bitOff)
                            Cells(A, 1).value = A - 1
                            Cells(A, SIDColA).value = "$2E"
                            Cells(A, IDColA).value = "$" + DIDNumber
                            Cells(A, SessionColA).value = "100" + CStr(session)
                            Cells(A, ServiceColA).value = "WRITE value OUTOFRANGE " + Str(dec) + " in " + ParamName + ": size = " + Str(size) + "; res = " + Str(res) + "; offset = " + Str(off) + " -> " + inBin
                            Cells(A, RequestColA).value = "2E" + DIDNumber + BinToHex(out)
                            Cells(A, ResponseColA).value = UDS_NACK_31_OutOfRange
                            A = A + 1
                        End If
                    End If
                End If
            End Select
        Dt = Dt + 1
    Loop
End Function

Public Function ListValueLoop(DIDdefValueBin As String)
    Dim Dt As Integer: Dt = D
    Dim ParamName As String
    Dim dec As Double
    Dim size As Integer
    Dim inBin As String
    Dim out As String
    Dim bitOff As Integer
    Dim ByteStart As Integer
    Dim temp() As String
    Dim i As Integer

    Do While Right(DIDRangeD.Cells(Dt, 1).value, 4) = DIDNumber
        If ListRangeD.Cells(Dt, 1).value <> 0 Then

            ParamName = NameRangeD.Cells(Dt, 1).value
            ByteStart = StartRangeD.Cells(Dt, 1).value
            bitOff = BitOffsetRangeD.Cells(Dt, 1).value
            size = SizeRangeD.Cells(Dt, 1).value
            temp = Split(CodingRangeD.Cells(Dt, 1).value, vbLf)
            For i = 0 To UBound(temp)
                If InStr(temp(i), "Not Used") = 0 Then
                    dec = CDbl(Left(temp(i), InStr(temp(i), ":") - 1))
                    inBin = NewDecToBin(Str(dec), size, True, NegRepr)
                    out = replaceInString(DIDdefValueBin, inBin, (ByteStart - 1) * 8 + bitOff)
                    Cells(A, 1).value = A - 1
                    Cells(A, SIDColA).value = "$2E"
                    Cells(A, IDColA).value = "$" + DIDNumber
                    Cells(A, SessionColA).value = "100" + CStr(session)
                    Cells(A, ServiceColA).value = "WRITE value " + Str(dec) + " in " + ParamName + ": size = " + Str(size) + " -> " + inBin
                    Cells(A, RequestColA).value = "2E" + DIDNumber + BinToHex(out)
                    Cells(A, ResponseColA).value = "6E" + DIDNumber
                    A = A + 1
                    Cells(A, 1).value = A - 1
                    Cells(A, SIDColA).value = "$22"
                    Cells(A, IDColA).value = "$" + DIDNumber
                    Cells(A, SessionColA).value = "100" + CStr(session)
                    Cells(A, ServiceColA).value = "CHECK value " + Str(dec) + " in " + ParamName
                    Cells(A, RequestColA).value = "22" + DIDNumber
                    Cells(A, ResponseColA).value = "62" + DIDNumber + BinToHex(out)
                    A = A + 1
                End If
            Next i
        End If
        Dt = Dt + 1
    Loop
End Function
