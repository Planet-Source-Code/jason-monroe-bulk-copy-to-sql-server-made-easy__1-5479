Attribute VB_Name = "GenCode"
Option Explicit

Public cTools As slpTools

Sub main()
10000    On Error GoTo BSS_ErrorHandler
10005    Set cTools = New slpTools

10010 Exit Sub

10015 BSS_ErrorHandler:

10020    If Err.Number > 0 then ProjectErrorHandler  "(Module) GenCode::Sub main"
10025    Resume Next
End Sub

'******************************************************************************
'*      Name: MsgBox                                                          *
'* Developer: Jason K. Monroe                                                 *
'*      Date: 1/7/2000                                                        *
'*    Inputs: All Inputs map to the standard MsgBox function defined in VBA   *
'*   Purpose: This MessageBox wrapper function will determine if the critical *
'*          : flag is set, if it is.. then it will log the message to a log   *
'*          : file specified by app.path and app.exename                      *
'******************************************************************************
Public Function MsgBox(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String, Optional HelpFile As String, Optional Context As Single) As VbMsgBoxResult
10030    On Error GoTo BSS_ErrorHandler

10035    Dim strErrorLog As String
10040    Dim iFileHandle As Integer
10045    Dim strErrorTitle As String
10050    Dim iResult As Integer
    
10055    iFileHandle = FreeFile
10060    strErrorTitle = App.EXEName & " " & Title
10065    strErrorLog = App.Path & "\" & App.EXEName & ".log"
    
10070    If (Buttons And vbCritical) Then
10075        Open strErrorLog For Append As #iFileHandle
10080        Print #iFileHandle, Now, Prompt
10085        Close #iFileHandle
10090    End If
    
10095    iResult = VBA.MsgBox(Prompt, Buttons, strErrorTitle, HelpFile, Context)
10100    MsgBox = iResult

10105 Exit Function

10110 BSS_ErrorHandler:

10115    If Err.Number > 0 then ProjectErrorHandler  "(Module) GenCode::Function MsgBox"
10120    Resume Next
End Function

Public Sub ProjectErrorHandler(MyMethod As String)
    
    Dim sErrStr As String
    Dim uResult As VbMsgBoxResult
    
    sErrStr = "Error " & Trim$(Str$(Err)) & " in " & MyMethod    
    If Erl Then
        sErrStr = sErrStr & " (Line #: " & Erl & ")"
    End If
        
    sErrStr = sErrStr & vbCrLf & "while running " & App.EXEName & ".exe v" & Format$(App.Major, "#") & "." & Format$(App.Minor, "0#")
    sErrStr = sErrStr & " (Build " & Format$(App.Revision, "#0") & ")" & vbCrLf & vbCrLf & "Error = '" & Error$ & "'"
    
    uResult = MsgBox(sErrStr, vbOKOnly + vbCritical + vbApplicationModal, "Error")
    
End Sub
