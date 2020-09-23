VERSION 5.00
Object = "{47FF0A28-F219-11CE-9541-00AA0044FE32}#1.0#0"; "VBSQL.OCX"
Begin VB.Form frmBCP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BCP Process"
   ClientHeight    =   2370
   ClientLeft      =   5520
   ClientTop       =   3630
   ClientWidth     =   3390
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   Begin VbsqlLib.Vbsql Vbsql1 
      Left            =   900
      Top             =   1140
      _Version        =   65536
      _ExtentX        =   2672
      _ExtentY        =   1296
      _StockProps     =   0
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label lblTable 
      Caption         =   "Label1"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   3315
   End
End
Attribute VB_Name = "frmBCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mVar_Batchcount As Long
Private mVar_Target As String

Public Property Get BatchCount() As Long
10000    On Error GoTo BSS_ErrorHandler
10005    BatchCount = mVar_Batchcount

10010 Exit Property

10015 BSS_ErrorHandler:

10020    If Err.Number > 0 then ProjectErrorHandler  "(Form) frmBCP::Property Get BatchCount"
10025    Resume Next
End Property
Public Property Let BatchCount(vData As Long)
10030    On Error GoTo BSS_ErrorHandler
10035    mVar_Batchcount = vData

10040 Exit Property

10045 BSS_ErrorHandler:

10050    If Err.Number > 0 then ProjectErrorHandler  "(Form) frmBCP::Property Let BatchCount"
10055    Resume Next
End Property

Public Property Get Target() As String
10060    On Error GoTo BSS_ErrorHandler
10065    Target = mVar_Target

10070 Exit Property

10075 BSS_ErrorHandler:

10080    If Err.Number > 0 then ProjectErrorHandler  "(Form) frmBCP::Property Get Target"
10085    Resume Next
End Property
Public Property Let Target(vData As String)
10090    On Error GoTo BSS_ErrorHandler
10095    mVar_Target = vData
10100    lblTable.Caption = mVar_Target

10105 Exit Property

10110 BSS_ErrorHandler:

10115    If Err.Number > 0 then ProjectErrorHandler  "(Form) frmBCP::Property Let Target"
10120    Resume Next
End Property

Private Sub Form_Load()
10125    On Error GoTo BSS_ErrorHandler
10130    cTools.Centerform Me
10135    lblTable.Caption = mVar_Target

10140 Exit Sub

10145 BSS_ErrorHandler:

10150    If Err.Number > 0 then ProjectErrorHandler  "(Form) frmBCP::Sub Form_Load"
10155    Resume Next
End Sub

Private Sub Vbsql1_Error(ByVal SqlConn As Long, ByVal Severity As Long, ByVal ErrorNum As Long, ByVal ErrorStr As String, ByVal OSErrorNum As Long, ByVal OSErrorStr As String, RetCode As Long)
10160    On Error GoTo BSS_ErrorHandler
10165    Dim strMessage As String
    Select Case ErrorNum
        Case 10050
10170            List1.AddItem mVar_Batchcount & " Rows Copied.."
10175            List1.Refresh
        Case Else
10180            strMessage = "   SqlConn: " & SqlConn & vbCrLf & _
                         "  Severity: " & Severity & vbCrLf & _
                         "  ErrorNum: " & ErrorNum & vbCrLf & _
                         "  ErrorStr: " & ErrorStr & vbCrLf & _
                         "OSErrorNum: " & OSErrorNum & vbCrLf & _
                         "OSErrorStr: " & OSErrorStr & vbCrLf & _
                         "   RetCode: " & RetCode
10185            MsgBox strMessage, vbOKOnly
10190    End Select

10195 Exit Sub

10200 BSS_ErrorHandler:

10205    If Err.Number > 0 then ProjectErrorHandler  "(Form) frmBCP::Sub Vbsql1_Error"
10210    Resume Next
End Sub

Private Sub Vbsql1_Message(ByVal SqlConn As Long, ByVal Message As Long, ByVal state As Long, ByVal Severity As Long, ByVal MsgStr As String, ByVal ServerNameStr As String, ByVal ProcNameStr As String, ByVal Line As Long)
10215    On Error GoTo BSS_ErrorHandler
10220    Dim strMessage As String
    Select Case Message
        Case 0, 5701
            '
        Case Else
10225            strMessage = "   SQLConn: " & SqlConn & vbCrLf & _
                         "  Message#: " & Message & vbCrLf & _
                         "     State: " & state & vbCrLf & _
                         "  Severity: " & Severity & vbCrLf & _
                         "    MsgStr: " & MsgStr & vbCrLf & _
                         "    Server: " & ServerNameStr & vbCrLf & _
                         "  ProcName: " & ProcNameStr & vbCrLf & _
                         "      Line: " & Line
10230            MsgBox strMessage, vbOKOnly, cTools.AppTitle(False)
10235    End Select

10240 Exit Sub

10245 BSS_ErrorHandler:

10250    If Err.Number > 0 then ProjectErrorHandler  "(Form) frmBCP::Sub Vbsql1_Message"
10255    Resume Next
End Sub
