VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "slpTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Private Variables
Private jmDeBug          As Boolean
Private strLongAppTitle  As String
Private strShortAppTitle As String
Private strCommandLine   As String
Private mvarNoLoad       As Boolean
'************************************
'*      API's                       *
'************************************
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'************************************
'*      Properties                  *
'************************************
Property Get AppTitle(bLong As Boolean) As String
10000    On Error GoTo BSS_ErrorHandler
10005    If bLong Then
10010        AppTitle = strLongAppTitle
10015    Else
10020        AppTitle = strShortAppTitle
10025    End If

10030 Exit Property

10035 BSS_ErrorHandler:

10040    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) slpTools::Property Get AppTitle"
10045    Resume Next
End Property

Property Get bDebug() As Boolean
10050    On Error GoTo BSS_ErrorHandler
10055    bDebug = jmDeBug

10060 Exit Property

10065 BSS_ErrorHandler:

10070    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) slpTools::Property Get bDebug"
10075    Resume Next
End Property

Property Get cmdLine() As String
10080    On Error GoTo BSS_ErrorHandler
10085    cmdLine = strCommandLine

10090 Exit Property

10095 BSS_ErrorHandler:

10100    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) slpTools::Property Get cmdLine"
10105    Resume Next
End Property

Property Get noLoad() As Boolean
10110    On Error GoTo BSS_ErrorHandler
10115    noLoad = mvarNoLoad

10120 Exit Property

10125 BSS_ErrorHandler:

10130    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) slpTools::Property Get noLoad"
10135    Resume Next
End Property
'************************************************************************
'*      Name: GetAppTitle                                               *
'* Developer: Jason K. Monroe                                           *
'*      Date: 10/18/1999                                                *
'*   Purpose: gets the app title and version                            *
'*    Inputs: bLong as boolean                                          *
'*   Returns: Appname & version/build "MyApp v1.1 Build 8"              *
'*          : Appname & Version "MyApp v1.1"                            *
'* Algorithm: None                                                      *
'************************************************************************
Private Function GetAppTitle(bLong As Boolean) As String
10140    On Error GoTo BSS_ErrorHandler
10145    Dim strOut As String
    
10150    On Error GoTo errhandler
    
10155    strOut = App.Title & " v" & App.Major & "." & App.Minor
10160    If bLong = True Then
10165        strOut = strOut & " Build " & App.Revision
10170    End If
10175    If jmDeBug Then strOut = strOut & " {DEBUG}"
10180    GetAppTitle = strOut
10185    Exit Function
10190 errhandler:
10195    jmErrMsg "GetAppTitle", Err.Number

10200 Exit Function

10205 BSS_ErrorHandler:

10210    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) slpTools::Function GetAppTitle"
10215    Resume Next
End Function

Private Sub Class_Initialize()
10220    On Error GoTo BSS_ErrorHandler
    '******************************************************
    '* Check to see if we are running in the VB5/6 IDE..  *
    '* if so, set the debug flag                          *
    '******************************************************
10225    Dim strfilename As String
10230    Dim lngCount As Long
10235    strfilename = String(255, 0)
10240    lngCount = GetModuleFileName(App.hInstance, strfilename, 255)
10245    strfilename = Left(strfilename, lngCount)
10250    If (UCase(Right(strfilename, 7)) = "VB5.EXE") Or (UCase(Right(strfilename, 7)) = "VB6.EXE") Then
10255        jmDeBug = True
10260    Else
10265        jmDeBug = False
10270    End If
    
    '********************************
    '* Setup our AppName Variables  *
    '********************************
10275    strLongAppTitle = GetAppTitle(True)
10280    strShortAppTitle = GetAppTitle(False)
    
    '***********************************************
    '* Store our command line params for retrieval *
    '***********************************************
10285    strCommandLine = UCase(GetToken(Command, " "))
10290    If UCase(GetToken("", " ")) = "NOLOAD" Then
10295        mvarNoLoad = True
10300    Else
10305        mvarNoLoad = False
10310    End If
    
10315 Exit Sub

10320 BSS_ErrorHandler:

10325    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) slpTools::Sub Class_Initialize"
10330    Resume Next
End Sub

'********************************************************************
'*      Name: CenterForm                                            *
'* Developer: Jason K. Monroe                                       *
'*      Date: 10/18/1999                                            *
'*   Purpose: Centers a form on the screen                          *
'*    Inputs: The form to center                                    *
'*   Returns: None                                                  *
'********************************************************************
Public Sub Centerform(frm As Form)
10335    On Error GoTo BSS_ErrorHandler
10340    Dim X As Long
10345    Dim Y As Long
10350    X = (Screen.Width \ 2) - (frm.Width \ 2)
10355    Y = (Screen.Height \ 2) - (frm.Height \ 2)
10360    frm.Move X, Y

10365 Exit Sub

10370 BSS_ErrorHandler:

10375    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) slpTools::Sub Centerform"
10380    Resume Next
End Sub

'************************************************************************************
'*      Name: SetMousePointer                                                       *
'* Developer: Jason K. Monroe                                                       *
'*      Date: 10/18/1999                                                            *
'*   Purpose: Used to track and set the mouse pointer state                         *
'*    Inputs: the state you would like the pointer to be (vbHourglass or vbDefault) *
'*   Returns: N/A                                                                   *
'* Algorithm: Incriments counter +1 for each hourglass setting                      *
'*          : Decriments counter -1 for each default setting if counter > 0         *
'*          : if counter > 0 sets mouse pointer = vbHourglass                       *
'*          : if counter = 0 sets mouse pointer = vbDefault                         *
'************************************************************************************
Public Sub SetMousePointer(state As Integer)
10385    On Error GoTo BSS_ErrorHandler
10390    Static iCounter As Integer
10395    On Error GoTo errhandler
10400    If state = vbHourglass Then iCounter = iCounter + 1
10405    If state = vbDefault Then
10410        If iCounter > 0 Then iCounter = iCounter - 1
10415    End If
10420    If iCounter > 0 Then
10425        Screen.MousePointer = vbHourglass
10430    Else
10435        Screen.MousePointer = vbDefault
10440    End If
10445    Exit Sub
10450 errhandler:
10455    jmErrMsg "SetMousePointer(" & state & ")", Err.Number

10460 Exit Sub

10465 BSS_ErrorHandler:

10470    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) slpTools::Sub SetMousePointer"
10475    Resume Next
End Sub

'********************************************************************************
'*      Name: CryptString                                                       *
'* Developer: Jason K. Monroe                                                   *
'*      Date: 10/18/1999                                                        *
'*   Purpose: Used for encrypt and decrypt a string                             *
'*    Inputs: string to work, EnCrypt as boolean                                *
'*   Returns: encrypted or decrypted string depending on EnCrypt                *
'* Algorithm: funky little add/subtract/multiply/divide/reversal twist thingy   *
'********************************************************************************
Public Function CryptString(txtString As String, EnCrypt As Boolean) As String
10480    On Error GoTo BSS_ErrorHandler
10485    On Error GoTo errhandler
10490    Dim X As Integer
10495    Dim outString As String
10500    Dim lLen As Long

10505    lLen = Len(txtString)
10510    outString = String(lLen, 0)

10515    If EnCrypt Then
10520        For X = 1 To Len(txtString)
10525            Mid$(outString, (lLen + 1) - X, 1) = Chr$((Asc(Mid$(txtString, X, 1)) * 2) + 3)
10530        Next X
10535     Else
10540        For X = 1 To Len(txtString)
10545            Mid$(outString, (lLen + 1) - X, 1) = Chr$((Asc(Mid$(txtString, X, 1)) - 3) * 0.5)
10550        Next X
10555    End If

10560    CryptString = outString
    
10565    Exit Function
10570 errhandler:
10575    jmErrMsg "CryptString(" & txtString & ")", Err.Number

10580 Exit Function

10585 BSS_ErrorHandler:

10590    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) slpTools::Function CryptString"
10595    Resume Next
End Function

'************************************************************************************
'* Name: QueryExit                                                                  *
'* Developer: Jason K. Monroe                                                       *
'* Date: 10/18/1999                                                                 *
'* Purpose: Used ask user if he/she wants to exit the application                   *
'* Inputs: none                                                                     *
'* Returns: 0 for Yes, 1 for no. Compaitable with "Cancel" flag in queryunload      *
'* Algorithm: Always returns 0 if running in VB IDE                                 *
'************************************************************************************
Public Function QueryExit() As Integer
10600    On Error GoTo BSS_ErrorHandler
10605    Dim Result As Integer
    
    ' First check to see if we are in the VB6 IDE..
10610    If jmDeBug Then
10615        QueryExit = 0
10620        Exit Function
10625    Else
10630        Result = MsgBox("Are you sure?", vbYesNo, AppTitle(False))
10635    End If
    
    ' Now check the return value of stuff
10640    If Result = vbYes Then
10645        QueryExit = 0
10650    Else
10655        QueryExit = 1
10660    End If

10665 Exit Function

10670 BSS_ErrorHandler:

10675    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) slpTools::Function QueryExit"
10680    Resume Next
End Function

'********************************************************************************
'*      Name: jmErrMsg                                                          *
'* Developer: Jason K. Monroe                                                   *
'*      Date: 10/18/1999                                                        *
'*   Purpose: Used to standardize the error message returns                     *
'*    Inputs: string to include in error message (Usually the err gen routine)  *
'*   Returns: none                                                              *
'* Algorithm: none                                                              *
'********************************************************************************
Public Sub jmErrMsg(txt As String, Optional errCode As Long)
10685    On Error GoTo BSS_ErrorHandler
10690    Dim sOut As String
10695    sOut = "Routine: " & txt & vbCrLf & IIf((errCode), errCode, Err.Number) & " == " & Err.Description
10700    MsgBox sOut
10705    If jmDeBug = True Then
10710        Debug.Print sOut
10715    End If

10720 Exit Sub

10725 BSS_ErrorHandler:

10730    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) slpTools::Sub jmErrMsg"
10735    Resume Next
End Sub

'********************************************************************
'*      Name: GetToken                                              *
'* Developer: Jason K. Monroe                                       *
'*      Date: 10/19/1999                                            *
'*   Purpose: C style string tokenizer                              *
'*    Inputs: string to work, delimiter                             *
'*   Returns: parsed string token                                   *
'* Algorithm: if passed string, reloads the tokenizer               *
'*          : empty call gives next token                           *
'********************************************************************
Public Function GetToken(new_txt As String, delimiter As String) As String
10740    On Error GoTo BSS_ErrorHandler
10745    Static txt As String
10750    Dim pos As Integer

    ' Save new text.
10755    If new_txt <> "" Then txt = new_txt

10760    pos = InStr(txt, delimiter)
10765    If pos < 1 Then pos = Len(txt) + 1
10770    GetToken = Left$(txt, pos - 1)
10775    pos = Len(txt) - pos + 1 - Len(delimiter)
10780    If pos < 1 Then
10785        txt = ""
10790    Else
10795        txt = Right$(txt, pos)
10800    End If

10805 Exit Function

10810 BSS_ErrorHandler:

10815    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) slpTools::Function GetToken"
10820    Resume Next
End Function

Public Sub AlwaysOnTop(frm As Form)
10825    On Error GoTo BSS_ErrorHandler
10830    SetWindowPos frm.hWnd, _
             -1, _
             frm.Left / Screen.TwipsPerPixelX, _
             frm.Top / Screen.TwipsPerPixelY, _
             frm.Width / Screen.TwipsPerPixelX, _
             frm.Height / Screen.TwipsPerPixelY, _
             &H10 Or &H40

10835 Exit Sub

10840 BSS_ErrorHandler:

10845    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) slpTools::Sub AlwaysOnTop"
10850    Resume Next
End Sub

Public Sub SnoozeButton(lMilliseconds As Long)
10855    On Error GoTo BSS_ErrorHandler
10860    Sleep lMilliseconds

10865 Exit Sub

10870 BSS_ErrorHandler:

10875    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) slpTools::Sub SnoozeButton"
10880    Resume Next
End Sub

Public Sub ExitClean()
10885    On Error GoTo BSS_ErrorHandler
10890    Dim frm As Form
10895    For Each frm In Forms
10900        Unload frm
10905    Next

10910 Exit Sub

10915 BSS_ErrorHandler:

10920    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) slpTools::Sub ExitClean"
10925    Resume Next
End Sub
