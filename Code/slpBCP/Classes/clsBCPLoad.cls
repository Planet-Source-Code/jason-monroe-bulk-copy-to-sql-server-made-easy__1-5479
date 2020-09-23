VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBCPLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Private Variables
Private mVar_TargetTable As String
Private mVar_SourceFile As String
Private mVar_TargetDB As String
Private mVar_TargetServer As String
Private mVar_SQLLogin As String
Private mVar_SQLPassword As String
Private mVar_BCPBatchCount As Integer
Private mVar_BCPColumnCount As Integer
Private mVar_RowsCopied As Long
Private mVar_frm As frmBCP
Private mVar_TruncateData As Boolean
Private mVar_ColLengths() As Long

' Private Constant declarations
Private Const ModuleName = "clsBCPLoad"
Private Const mVar_Def_BCPBatchCount = 1000
Private Const mVar_Def_TruncateData = True

'Error codes
Private Enum SQLErrorCodes
    ieLogin = 5001
    ieBCPLogin = 5002
    ieSQLOpen = 5003
    ieSourceFileNotFound = 5004
    ieBCPInitFailed = 5005
    ieBCPControl = 5006
    ieBCPColumns = 5007
    ieBCPColFormat = 5008
    ieBCPExec = 5009
    ieCMDFail = 5010
    ieSQLExec = 5011
    ieSQLResults = 5012
End Enum

' Event Descriptions

' Property Declarations
Public Property Get RowsCopied() As Long
10000    On Error GoTo BSS_ErrorHandler
10005    RowsCopied = mVar_RowsCopied

10010 Exit Property

10015 BSS_ErrorHandler:

10020    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Get RowsCopied"
10025    Resume Next
End Property

Public Property Get TruncateData() As Boolean
10030    On Error GoTo BSS_ErrorHandler
10035    TruncateData = mVar_TruncateData

10040 Exit Property

10045 BSS_ErrorHandler:

10050    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Get TruncateData"
10055    Resume Next
End Property
Public Property Let TruncateData(vData As Boolean)
10060    On Error GoTo BSS_ErrorHandler
10065    mVar_TruncateData = vData

10070 Exit Property

10075 BSS_ErrorHandler:

10080    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Let TruncateData"
10085    Resume Next
End Property

Public Property Get TargetTable() As String
10090    On Error GoTo BSS_ErrorHandler
10095    TargetTable = mVar_TargetTable

10100 Exit Property

10105 BSS_ErrorHandler:

10110    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Get TargetTable"
10115    Resume Next
End Property
Public Property Let TargetTable(vData As String)
10120    On Error GoTo BSS_ErrorHandler
10125    mVar_TargetTable = vData

10130 Exit Property

10135 BSS_ErrorHandler:

10140    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Let TargetTable"
10145    Resume Next
End Property

Public Property Get TargetDB() As String
10150    On Error GoTo BSS_ErrorHandler
10155    TargetDB = mVar_TargetDB

10160 Exit Property

10165 BSS_ErrorHandler:

10170    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Get TargetDB"
10175    Resume Next
End Property
Public Property Let TargetDB(vData As String)
10180    On Error GoTo BSS_ErrorHandler
10185    mVar_TargetDB = vData

10190 Exit Property

10195 BSS_ErrorHandler:

10200    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Let TargetDB"
10205    Resume Next
End Property

Public Property Get TargetServer() As String
10210    On Error GoTo BSS_ErrorHandler
10215    TargetServer = mVar_TargetServer

10220 Exit Property

10225 BSS_ErrorHandler:

10230    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Get TargetServer"
10235    Resume Next
End Property
Public Property Let TargetServer(vData As String)
10240    On Error GoTo BSS_ErrorHandler
10245    mVar_TargetServer = vData

10250 Exit Property

10255 BSS_ErrorHandler:

10260    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Let TargetServer"
10265    Resume Next
End Property

Public Property Get Login() As String
10270    On Error GoTo BSS_ErrorHandler
10275    Login = mVar_SQLLogin

10280 Exit Property

10285 BSS_ErrorHandler:

10290    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Get Login"
10295    Resume Next
End Property
Public Property Let Login(vData As String)
10300    On Error GoTo BSS_ErrorHandler
10305    mVar_SQLLogin = vData

10310 Exit Property

10315 BSS_ErrorHandler:

10320    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Let Login"
10325    Resume Next
End Property

Public Property Get Password() As String
10330    On Error GoTo BSS_ErrorHandler
10335    Password = mVar_SQLPassword

10340 Exit Property

10345 BSS_ErrorHandler:

10350    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Get Password"
10355    Resume Next
End Property
Public Property Let Password(vData As String)
10360    On Error GoTo BSS_ErrorHandler
10365    mVar_SQLPassword = vData

10370 Exit Property

10375 BSS_ErrorHandler:

10380    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Let Password"
10385    Resume Next
End Property

Public Property Get SourceFile() As String
10390    On Error GoTo BSS_ErrorHandler
10395    SourceFile = mVar_SourceFile

10400 Exit Property

10405 BSS_ErrorHandler:

10410    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Get SourceFile"
10415    Resume Next
End Property
Public Property Let SourceFile(vData As String)
10420    On Error GoTo BSS_ErrorHandler
10425    mVar_SourceFile = vData

10430 Exit Property

10435 BSS_ErrorHandler:

10440    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Let SourceFile"
10445    Resume Next
End Property

Public Property Get BCPBatchCount() As Integer
10450    On Error GoTo BSS_ErrorHandler
10455    BCPBatchCount = mVar_BCPBatchCount

10460 Exit Property

10465 BSS_ErrorHandler:

10470    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Get BCPBatchCount"
10475    Resume Next
End Property
Public Property Let BCPBatchCount(vData As Integer)
10480    On Error GoTo BSS_ErrorHandler
10485    mVar_BCPBatchCount = vData

10490 Exit Property

10495 BSS_ErrorHandler:

10500    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Let BCPBatchCount"
10505    Resume Next
End Property

Public Property Get BCPColumnCount() As Integer
10510    On Error GoTo BSS_ErrorHandler
10515    BCPColumnCount = mVar_BCPColumnCount

10520 Exit Property

10525 BSS_ErrorHandler:

10530    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Get BCPColumnCount"
10535    Resume Next
End Property
Public Property Let BCPColumnCount(vData As Integer)
10540    On Error GoTo BSS_ErrorHandler
10545    mVar_BCPColumnCount = vData

10550 Exit Property

10555 BSS_ErrorHandler:

10560    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Property Let BCPColumnCount"
10565    Resume Next
End Property

Public Function StartBulkCopy() As Boolean
10570    On Error GoTo BSS_ErrorHandler
10575    Set mVar_frm = New frmBCP
10580    Load mVar_frm
10585    mVar_frm.Target = mVar_TargetDB & ".." & mVar_TargetTable
10590    mVar_frm.BatchCount = mVar_BCPBatchCount
10595    mVar_frm.Show
10600    DoEvents
    ' First we need to truncate our tables if needed
10605    If mVar_TruncateData Then
10610        DoTruncateTable
10615    End If
    ' We have to get our column length data
10620    Call GetColLengths
    ' Now we need to actually do our BulkCopy
10625    Call DoBCP
    ' Wow.. we are still here.. let's set the return flag to true then :-)
10630    StartBulkCopy = True

10635 Exit Function

10640 BSS_ErrorHandler:

10645    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Function StartBulkCopy"
10650    Resume Next
End Function

Private Sub Class_Initialize()
10655    On Error GoTo BSS_ErrorHandler
10660    Dim iResult As Integer
    ' The very first thing is init the SQL Library..
10665    If SqlInit() = "" Then
10670        Err.Raise 5000, "clsBCPLoad_Initialize", "Unable to Init the SQL Low Level Library"
10675    End If
    ' Next setup our constants and such
10680    mVar_BCPBatchCount = mVar_Def_BCPBatchCount
10685    mVar_TruncateData = mVar_Def_TruncateData

10690 Exit Sub

10695 BSS_ErrorHandler:

10700    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Sub Class_Initialize"
10705    Resume Next
End Sub

Private Sub Class_Terminate()
10710    On Error GoTo BSS_ErrorHandler
10715    SqlExit
10720    SqlWinExit
10725    Unload mVar_frm

10730 Exit Sub

10735 BSS_ErrorHandler:

10740    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Sub Class_Terminate"
10745    Resume Next
End Sub

Private Sub DoTruncateTable()
10750    On Error GoTo BSS_ErrorHandler
10755    Dim iResult As Long
10760    Dim iLogin As Long
10765    Dim iSQLConn As Long
10770    Dim strSQL As String
10775    Const ProcName = ModuleName & ".DoTruncateTable"
    ' Establish a login record for this process
10780    iLogin = SqlLogin()
10785    If iLogin = FAIL Then
10790        Call RaiseError(ieLogin, ProcName)
10795    End If
    ' Set the UesrName
10800    iResult = SqlSetLUser(iLogin, mVar_SQLLogin)
10805    If iResult = FAIL Then
10810        Call RaiseError(ieLogin, ProcName)
10815    End If
    ' Set the Password
10820    iResult = SqlSetLPwd(iLogin, mVar_SQLPassword)
10825    If iResult = FAIL Then
10830        Call RaiseError(ieLogin, ProcName)
10835    End If
    ' Open the SQL Connection
10840    iSQLConn = SqlOpen(iLogin, mVar_TargetServer)
10845    If iSQLConn = FAIL Then
10850        Call RaiseError(ieSQLOpen, ProcName)
10855    End If
    ' Setup the SQL Syntax
10860    strSQL = "Truncate table " & mVar_TargetDB & ".." & mVar_TargetTable
    ' Add the SQL Statement to the command stack
10865    iResult = SqlCmd(iSQLConn, strSQL)
10870    If iResult = FAIL Then
10875        Call RaiseError(ieCMDFail, ProcName, strSQL)
10880    End If
    ' Execute the SQL procedure
10885    iResult = SqlExec(iSQLConn)
10890    If iResult = FAIL Then
10895        Call RaiseError(ieSQLExec, ProcName)
10900    End If
    ' We must call our Results..
10905    iResult = SqlResults(iSQLConn)
10910    If iResult = FAIL Then
10915        Call RaiseError(ieSQLResults, ProcName)
10920    End If
    ' Now to close our SQL Connection
10925    SqlClose (iSQLConn)

10930 Exit Sub

10935 BSS_ErrorHandler:

10940    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Sub DoTruncateTable"
10945    Resume Next
End Sub

Private Sub DoBCP()
10950    On Error GoTo BSS_ErrorHandler
10955    Dim iResult As Long
10960    Dim iSQLConn As Long
10965    Dim iLogin As Long
10970    Dim X As Integer
10975    Dim iFile As Integer
10980    Dim sErrFile As String
10985    Const ProcName = ModuleName & ".DoBCP"
    ' Establish a login record for this BCP Process
10990    iLogin = SqlLogin()
10995    If iLogin = FAIL Then
11000        Call RaiseError(ieLogin, ProcName)
11005    End If
    ' Set the user name
11010    iResult = SqlSetLUser(iLogin, mVar_SQLLogin)
11015    If iResult = FAIL Then
11020        Call RaiseError(ieLogin, ProcName)
11025    End If
    ' Set the Login Password
11030    iResult = SqlSetLPwd(iLogin, mVar_SQLPassword)
11035    If iResult = FAIL Then
11040        Call RaiseError(ieLogin, ProcName)
11045    End If
    ' Enable BCP for this Login
11050    iResult = SqlBCPSetL(iLogin, True)
11055    If iResult = FAIL Then
11060        Call RaiseError(ieBCPLogin, ProcName)
11065    End If
    ' Open up a connection to our server
11070    iSQLConn = SqlOpen(iLogin, mVar_TargetServer)
11075    If iSQLConn = FAIL Then
11080        Call RaiseError(ieSQLOpen, ProcName)
11085    End If
    ' Now it's time to INIT DBLib for BCP processing
11090    If Dir(mVar_SourceFile) = "" Then       ' Make sure our source file is there
11095        Call RaiseError(ieSourceFileNotFound, ProcName)
11100    End If
11105    sErrFile = Left(mVar_SourceFile, InStr(1, mVar_SourceFile, ".") - 1) & ".err"
11110    If Dir(sErrFile) <> "" Then ' check for our error file and
11115        Kill sErrFile           ' delete it if found
11120    End If
11125    iFile = FreeFile
11130    Open sErrFile For Output As #iFile  ' Create an empty error file
11135        Print #iFile, Chr(0)
11140    Close #iFile
11145    iResult = SqlBCPInit(iSQLConn, mVar_TargetDB & ".." & mVar_TargetTable, mVar_SourceFile, sErrFile, DBIN)
11150    If iResult = FAIL Then
11155        Call RaiseError(ieBCPInitFailed, ProcName)
11160    End If
    ' Set the number of Rows per batch
11165    iResult = SqlBCPControl(iSQLConn, BCPBATCH, mVar_BCPBatchCount)
11170    If iResult = FAIL Then
11175        Call RaiseError(ieBCPControl, ProcName)
11180    End If
    ' Set the number of Columns for this Process
11185    iResult = SqlBCPColumns(iSQLConn, mVar_BCPColumnCount)
11190    If iResult = FAIL Then
11195        Call RaiseError(ieBCPColumns, ProcName)
11200    End If
    ' Set the column format.. we are using Fixed Length files.. this has to be done for each column
11205    For X = 1 To mVar_BCPColumnCount - 1
        'iResult = SqlBCPColfmt(iSQLConn, x, 0, -1, -1, Chr(0), 0, x)
        'iResult = SqlBCPColfmt(iSQLConn, x, 0, 0, -1, Chr(0), 0, x)
11210        iResult = SqlBCPColfmt(iSQLConn, X, SQLCHAR, 0, mVar_ColLengths(X), Chr(0), 0, X)
11215        If iResult = FAIL Then
11220            Call RaiseError(ieBCPColFormat, ProcName)
11225        End If
11230    Next X
11235    iResult = SqlBCPColfmt(iSQLConn, mVar_BCPColumnCount, SQLCHAR, 0, mVar_ColLengths(mVar_BCPColumnCount), Chr(13) & Chr(10), 2, mVar_BCPColumnCount)
11240    If iResult = FAIL Then
11245        Call RaiseError(ieBCPColFormat, ProcName)
11250    End If
    ' Actually do the BCP now.
11255    iResult = SqlBCPExec(iSQLConn, mVar_RowsCopied)
11260    If iResult = FAIL Then
11265        Call RaiseError(ieBCPExec, ProcName)
11270    End If
11275    SqlClose iSQLConn

11280 Exit Sub

11285 BSS_ErrorHandler:

11290    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Sub DoBCP"
11295    Resume Next
End Sub

Private Sub RaiseError(numError As SQLErrorCodes, ProcName As String, Optional ExtraInfo As Variant)
11300    On Error GoTo BSS_ErrorHandler
11305    Dim strErrMessage As String
    Select Case numError
        Case ieLogin
11310            strErrMessage = "DBLib could not allocate login record"
        Case ieBCPLogin
11315            strErrMessage = "DBLib count not Enable BCP for this login"
        Case ieSQLOpen
11320            strErrMessage = "DBLib count not open connection to server mvar_targetserver"
        Case ieSourceFileNotFound
11325            strErrMessage = mVar_SourceFile & vbCrLf & "Not found on system"
        Case ieBCPInitFailed
11330            strErrMessage = "DBLib unable to Init the BCP functionality"
        Case ieBCPControl
11335            strErrMessage = "DBLib unable to set Batch Row Count"
        Case ieBCPColumns
11340            strErrMessage = "DBLib unable to set the BCP Column Count : " & mVar_BCPColumnCount
        Case ieBCPColFormat
11345            strErrMessage = "DBLib unable to set the Column Format for column"
        Case ieBCPExec
11350            strErrMessage = "DBLib Execute BCP Failed"
        Case ieCMDFail
11355            strErrMessage = "DBLib Failed to put the SQL Command on the Stack" & vbCrLf & ExtraInfo
        Case ieSQLExec
11360            strErrMessage = "DBLib SQL Execute Failed"
        Case ieSQLResults
11365            strErrMessage = "DBLib Get of SQL Results failed"
        
        Case Else
11370            strErrMessage = "Unknown error has occured in " & ProcName
11375    End Select
11380    Err.Raise numError, ProcName, strErrMessage

11385 Exit Sub

11390 BSS_ErrorHandler:

11395    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Sub RaiseError"
11400    Resume Next
End Sub

Private Sub GetColLengths()
11405    On Error GoTo BSS_ErrorHandler
11410    Dim iResult As Long
11415    Dim iLogin As Long
11420    Dim iSQLConn As Long
11425    Dim strSQL As String
11430    Dim X As Long
11435    Const ProcName = ModuleName & ".GetColLengths"
    ' Establish a login record for this process
11440    iLogin = SqlLogin()
11445    If iLogin = FAIL Then
11450        Call RaiseError(ieLogin, ProcName)
11455    End If
    ' Set the UesrName
11460    iResult = SqlSetLUser(iLogin, mVar_SQLLogin)
11465    If iResult = FAIL Then
11470        Call RaiseError(ieLogin, ProcName)
11475    End If
    ' Set the Password
11480    iResult = SqlSetLPwd(iLogin, mVar_SQLPassword)
11485    If iResult = FAIL Then
11490        Call RaiseError(ieLogin, ProcName)
11495    End If
    ' Open the SQL Connection
11500    iSQLConn = SqlOpen(iLogin, mVar_TargetServer)
11505    If iSQLConn = FAIL Then
11510        Call RaiseError(ieSQLOpen, ProcName)
11515    End If
    ' Setup the SQL Syntax
11520    strSQL = "select cols.length, cols.colid " & _
             "from " & mVar_TargetDB & "..syscolumns COLS, " & _
             mVar_TargetDB & "..sysobjects objs " & _
             "Where objs.Name = '" & mVar_TargetTable & "' " & _
             "and cols.id = objs.id "
    ' Add the SQL Statement to the command stack
11525    iResult = SqlCmd(iSQLConn, strSQL)
11530    If iResult = FAIL Then
11535        Call RaiseError(ieCMDFail, ProcName, strSQL)
11540    End If
    ' Execute the SQL procedure
11545    iResult = SqlExec(iSQLConn)
11550    If iResult = FAIL Then
11555        Call RaiseError(ieSQLExec, ProcName)
11560    End If
    ' We must call our Results..
11565    iResult = SqlResults(iSQLConn)
11570    If iResult = FAIL Then
11575        Call RaiseError(ieSQLResults, ProcName)
11580    End If
    ' Setup our Array for use
11585    ReDim Preserve mVar_ColLengths(mVar_BCPColumnCount)
11590    X = 1
11595    Do Until SqlNextRow(iSQLConn) = NOMOREROWS
11600        mVar_ColLengths(X) = SqlData(iSQLConn, 1)
11605        X = X + 1
11610    Loop
    ' Now to close our SQL Connection
11615    SqlClose (iSQLConn)

11620 Exit Sub

11625 BSS_ErrorHandler:

11630    If Err.Number > 0 then ProjectErrorHandler  "(Class Module) clsBCPLoad::Sub GetColLengths"
11635    Resume Next
End Sub
