VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSubs 
   Caption         =   "Subs and Functions."
   ClientHeight    =   6885
   ClientLeft      =   2430
   ClientTop       =   1530
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   9615
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   6510
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd 
      Caption         =   "View Res"
      Height          =   375
      Index           =   6
      Left            =   3960
      TabIndex        =   18
      ToolTipText     =   "View Result File where all subs/functions were added"
      Top             =   6120
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3120
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFilePath 
      Caption         =   "..."
      Height          =   285
      Left            =   3480
      TabIndex        =   17
      ToolTipText     =   "Select File to Add New Subs"
      Top             =   6120
      Width           =   375
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   360
      TabIndex        =   16
      Text            =   "txtFilePath"
      Top             =   6120
      Width           =   3135
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Add Sub"
      Height          =   375
      Index           =   5
      Left            =   4920
      TabIndex        =   14
      ToolTipText     =   "Adds Sub to File"
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "View Sub"
      Height          =   375
      Index           =   4
      Left            =   5880
      TabIndex        =   13
      ToolTipText     =   "Veiw Selected Sub"
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "View File"
      Height          =   375
      Index           =   3
      Left            =   6840
      TabIndex        =   12
      ToolTipText     =   "View Selected File"
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Clean "
      Height          =   375
      Index           =   2
      Left            =   7800
      TabIndex        =   11
      ToolTipText     =   "Delete ALL Subs/Functions from DB List"
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Check"
      Height          =   375
      Index           =   1
      Left            =   8760
      TabIndex        =   10
      ToolTipText     =   "Check Selected Path/All Files with frm, bas suffix"
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox txtSub 
      Height          =   285
      Left            =   6120
      TabIndex        =   9
      Text            =   "txtSub"
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox cboView 
      Height          =   315
      ItemData        =   "frmSubs.frx":0000
      Left            =   8160
      List            =   "frmSubs.frx":001F
      TabIndex        =   6
      Text            =   "ALL"
      ToolTipText     =   "Select view mode"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Text            =   "txtPath"
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdCallF 
      Caption         =   "..."
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      ToolTipText     =   "Select Folder to check"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Text            =   "txtFile"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Data dcSubs 
      Caption         =   "dcSubs"
      Connect         =   "Access"
      DatabaseName    =   "D:\0_FileLib\Data\SystemFileList.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSDBGrid.DBGrid grdSubs 
      Bindings        =   "frmSubs.frx":0092
      Height          =   5535
      Left            =   0
      OleObjectBlob   =   "frmSubs.frx":00A7
      TabIndex        =   0
      Top             =   480
      Width           =   9615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File:"
      Height          =   195
      Index           =   7
      Left            =   0
      TabIndex        =   15
      Top             =   6120
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sub:"
      Height          =   195
      Index           =   4
      Left            =   5640
      TabIndex        =   8
      Top             =   120
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "View:"
      Height          =   195
      Index           =   6
      Left            =   7680
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Path:"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File:"
      Height          =   195
      Index           =   3
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   285
   End
End
Attribute VB_Name = "frmSubs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private msSql           As String
Private msSqlSel        As String
Private msSqlOrder      As String
Private mlFilesNum      As Long
Private msSqlSelView    As String

Private Const miS_FILE          As Integer = 1
Private Const miS_SUB           As Integer = 2
Private Const miS_TOTAL         As Integer = 3

Private Sub CheckView()
On Error GoTo ErrHandler
    Select Case cboView
    Case "ALL":                     msSqlSelView = ""
    Case "Subs":                    msSqlSelView = " and SubType='S'"
    Case "Functions":               msSqlSelView = " and SubType='F'"
    Case "Public Subs":             msSqlSelView = " and SubType='S'  and SubScope='G'"
    Case "Public Functions":        msSqlSelView = " and SubType='F'  and SubScope='G'"
    Case "Private Subs":            msSqlSelView = " and SubType='S'  and SubScope='P'"
    Case "Private Functions":       msSqlSelView = " and SubType='F'  and SubScope='P'"
    Case "Public API":              msSqlSelView = " and SubType='A'  and SubScope='G'"
    Case "Private API":             msSqlSelView = " and SubType='A'  and SubScope='P'"
    Case Else:          msSqlSelView = ""
    End Select
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "CheckView")
End Sub

Private Sub cboView_Click()
On Error GoTo ErrHandler
    Dim sSql As String, rst As DAO.Recordset
    
    Screen.MousePointer = vbHourglass
    
    CheckView
    sSql = msSql & msSqlSel & msSqlSelView & msSqlOrder
    dcSubs.RecordSource = sSql
    dcSubs.Refresh
    
    sSql = "select Count(SubID) as Num "
    sSql = sSql & " FROM FileList"
    sSql = sSql & " INNER JOIN Subs"
    sSql = sSql & " ON FileList.FileID =Subs.FileID"
    sSql = sSql & msSqlSel & msSqlSelView
    Set rst = SQLOpenRecordset(gdbFL, sSql)
    
    sb.Panels(miS_TOTAL).Text = rst!Num
    DoEvents
    Screen.MousePointer = vbDefault
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "cboView_Click")
End Sub

Private Sub cmd_Click(Index As Integer)
On Error GoTo ErrHandler
    Dim sSql        As String
    Dim rst         As Recordset
    Dim rstCheck    As Recordset
    
    Dim sFilePath   As String
    Dim iFile       As Integer
    Dim s           As String
    Dim sType       As String
    Dim sScope      As String
    Dim sSub        As String
    Dim sMess       As String
    Dim i           As Integer
    Dim l           As Long
    Dim iFound      As Integer
    Dim b           As Boolean
    
    Const iB_CHECK          As Integer = 1
    Const iB_CLEAR          As Integer = 2
    Const iB_VIEW           As Integer = 3
    Const iB_VIEW_SUB       As Integer = 4
    Const iB_ADD_SUB        As Integer = 5
    Const iB_VIEW_RESULT    As Integer = 6
    
    Select Case Index
    Case iB_VIEW_RESULT
        If txtFilePath = "" Then
            MsgBox "Please define the Target File."
            cmdFilePath_Click
            Exit Sub
        End If
        l = Shell("notepad.exe " & txtFilePath, vbNormalFocus)
        
    Case iB_ADD_SUB
        If txtFilePath = "" Then
            MsgBox "Please define the Target File."
            cmdFilePath_Click
            Exit Sub
        End If
        
        Call SaveSetting("FL", "FllePaths", "LastFllePathSub", txtFilePath)
        
        iFound = 0
        sFilePath = dcSubs.Recordset!FilePath & "\" & dcSubs.Recordset!FileName
        If Dir(sFilePath) <> "" And sFilePath <> "" Then
            iFile = FreeFile
            Open sFilePath For Input Shared As #iFile
            sb.Panels(miS_FILE).Text = "File:" & sFilePath
            DoEvents
            Do Until EOF(iFile)
                Line Input #iFile, s
                s = s
                If iFound = 1 Then
                    If InStr(1, s, "End Sub") > 0 Or InStr(1, s, "End Function") > 0 Then
                        iFound = 0
                        sSub = sSub & s & vbCrLf
                        b = AppendToTextFile(txtFilePath, sSub)
                        Exit Do
                    Else
                        sSub = sSub & s & vbCrLf
                    End If
                End If
                If iFound = 2 Then
                    If InStr(1, s, ")") > 0 Then
                        iFound = 0
                        sSub = sSub & s & vbCrLf
WriteToFile2:
                        b = AppendToTextFile(txtFilePath, sSub)
                        Exit Do
                    Else
                        sSub = sSub & s & vbCrLf
                    End If
                End If
                
                If dcSubs.Recordset!SubType = "A" Then
                    If InStr(1, LCase(s), " sub " & dcSubs.Recordset!SubName) > 0 Or _
                     InStr(1, LCase(s), " function " & dcSubs.Recordset!SubName) > 0 Then
                        sSub = sSub & s & vbCrLf
                        iFound = 2
                        If InStr(1, s, ")") > 0 Then
                            GoTo WriteToFile2:
                            iFound = 0
                        End If
                    End If
                Else
                    If InStr(1, LCase(s), " sub " & dcSubs.Recordset!SubName) > 0 Or _
                     InStr(1, LCase(s), " function " & dcSubs.Recordset!SubName) > 0 Then
                        sSub = sSub & s & vbCrLf
                        iFound = 1
                    End If
                End If
            Loop
        End If
        
    Case iB_VIEW_SUB
        iFound = 0
        sFilePath = dcSubs.Recordset!FilePath & "\" & dcSubs.Recordset!FileName
        If Dir(sFilePath) <> "" And sFilePath <> "" Then
            iFile = FreeFile
            Open sFilePath For Input Shared As #iFile
            sb.Panels(miS_FILE).Text = "File:" & sFilePath
            DoEvents
            Do Until EOF(iFile)
                Line Input #iFile, s
                s = s
                If iFound = 1 Then
                    If InStr(1, s, "End Sub") > 0 Or InStr(1, s, "End Function") > 0 Then
                        iFound = 0
                        sSub = sSub & s & vbCrLf
                        Call WriteToFile("c:\LastSub.txt", sSub)
                        l = Shell("notepad.exe 'c:\LastSub.txt'" & sFilePath, vbNormalFocus)
                        Exit Do
                    Else
                        sSub = sSub & s & vbCrLf
                    End If
                End If
                If iFound = 2 Then
                    If InStr(1, s, ")") > 0 Then
                        iFound = 0
                        sSub = sSub & s & vbCrLf
WriteToFile1:
                        Call WriteToFile("c:\LastSub.txt", sSub)
                        l = Shell("notepad.exe 'c:\LastSub.txt'" & sFilePath, vbNormalFocus)
                        Exit Do
                    Else
                        sSub = sSub & s & vbCrLf
                    End If
                End If
                If dcSubs.Recordset!SubType = "A" Then
                    If InStr(1, LCase(s), " sub " & dcSubs.Recordset!SubName) > 0 Or _
                     InStr(1, LCase(s), " function " & dcSubs.Recordset!SubName) > 0 Then
                        sSub = sSub & s & vbCrLf
                        iFound = 2
                        If InStr(1, s, ")") > 0 Then
                            GoTo WriteToFile1:
                            iFound = 0
                        End If
                    End If
                Else
                    If InStr(1, LCase(s), " sub " & dcSubs.Recordset!SubName) > 0 Or _
                     InStr(1, LCase(s), " function " & dcSubs.Recordset!SubName) > 0 Then
                        sSub = sSub & s & vbCrLf
                        iFound = 1
                    End If
                End If
            Loop
        End If
        
    Case iB_VIEW
        sFilePath = dcSubs.Recordset!FilePath & "\" & dcSubs.Recordset!FileName
        l = Shell("notepad.exe " & sFilePath, vbNormalFocus)
    
    Case iB_CLEAR
        i = MsgBox("Do you want to delete ALL records from Subs List?", vbYesNo)
        If i <> vbYes Then Exit Sub
        sSql = "delete from subs"
        Call SQLExecute(gdbFL, sSql)
        
        dcSubs.Refresh
        
    Case iB_CHECK
        
        mlFilesNum = 0
        If txtPath = "" Then
            sSql = "select * from FileList where FileSuffix='frm' or FileSuffix='bas'"
        Else
            sSql = "select * from FileList where FilePath like '" & txtPath & "*'"
            sSql = sSql & " and (FileSuffix='frm' or FileSuffix='bas')"
        End If
        
        Set rst = SQLOpenRecordset(gdbFL, sSql)
        
        Do Until rst.EOF
            iFile = FreeFile
            sFilePath = rst!FilePath & "\" & rst!FileName
            If Dir(sFilePath) <> "" And sFilePath <> "" Then
                Open sFilePath For Input Shared As #iFile
                sb.Panels(miS_FILE).Text = "File:" & sFilePath
                DoEvents
                Do Until EOF(iFile)
ReadNext:
                    Line Input #iFile, s
                    s = LCase(s)
                    If InStr(1, s, "private sub") > 0 Then
                        sType = "S"
                        sScope = "P"
                        If InStr(1, s, "(") <> 0 Then
                            s = (Left(s, InStr(1, s, "(") - 1))
                        End If
                        sSub = Trim(Replace(s, "private sub", ""))
                        GoTo WriteSub
                    End If
                    
                    If InStr(1, (s), "private function") > 0 Then
                        sType = "F"
                        sScope = "P"
                        If InStr(1, s, "(") <> 0 Then
                            s = (Left(s, InStr(1, s, "(") - 1))
                        End If
                        sSub = Trim(Replace(s, "private function", ""))
                        GoTo WriteSub
                    End If
                    
                    If InStr(1, (s), "public sub") > 0 Then
                        sType = "S"
                        sScope = "G"
                        If InStr(1, s, "(") <> 0 Then
                            s = (Left(s, InStr(1, s, "(") - 1))
                        End If
                        sSub = Trim(Replace(s, "public sub", ""))
                        GoTo WriteSub
                    End If
                    
                    If InStr(1, (s), "public function") > 0 Then
                        sType = "F"
                        sScope = "G"
                        If InStr(1, s, "(") <> 0 Then
                            s = (Left(s, InStr(1, s, "(") - 1))
                        End If
                        sSub = Trim(Replace(s, "public function", ""))
                        GoTo WriteSub
                    End If
                    
                    If InStr(1, (s), "public declare function") > 0 Then
                        sType = "A"
                        sScope = "G"
                        If InStr(1, s, "lib") <> 0 Then
                            s = (Left(s, InStr(1, s, "lib") - 1))
                        End If
                        sSub = Trim(Replace(s, "public declare function", ""))
                        GoTo WriteSub
                    End If
                    
                    If InStr(1, (s), "private declare function") > 0 Then
                        sType = "A"
                        sScope = "P"
                        If InStr(1, s, "lib") <> 0 Then
                            s = (Left(s, InStr(1, s, "lib") - 1))
                        End If
                        sSub = Trim(Replace(s, "private declare function", ""))
                        GoTo WriteSub
                    End If
                    
                    If InStr(1, (s), "public declare sub") > 0 Then
                        sType = "A"
                        sScope = "G"
                        If InStr(1, s, "lib") <> 0 Then
                            s = (Left(s, InStr(1, s, "lib") - 1))
                        End If
                        sSub = Trim(Replace(s, "public declare sub", ""))
                        GoTo WriteSub
                    End If
                    
                    If InStr(1, (s), "private declare sub") > 0 Then
                        sType = "A"
                        sScope = "P"
                        If InStr(1, s, "lib") <> 0 Then
                            s = (Left(s, InStr(1, s, "lib") - 1))
                        End If
                        sSub = Trim(Replace(s, "private declare sub", ""))
                        GoTo WriteSub
                    End If
                    
                    If InStr(1, (s), "msgbox") > 0 Then
                        sMess = Trim(Replace((s), "msgbox", ""))
                        GoTo WriteMess
                    End If
                    GoTo NextStep
WriteSub:

                    sb.Panels(miS_SUB).Text = "Sub:" & sSub
                    mlFilesNum = mlFilesNum + 1
                    sb.Panels(miS_TOTAL).Text = "Total:" & mlFilesNum
                    DoEvents
                    sSql = "select * from Subs where "
                    sSql = sSql & " FileID=" & rst!FileID
                    sSql = sSql & " and SubName=" & SQLCheck(sSub)
                    Set rstCheck = SQLOpenRecordset(gdbFL, sSql)
                    If Not rstCheck Is Nothing Then
                        If rstCheck.EOF Then
                            ' INSERT STATEMENT
                            sSql = "insert into Subs ("
                            sSql = sSql & "Checked,"
                            sSql = sSql & "FileID,"
                            sSql = sSql & "StatusID,"
                            sSql = sSql & "SubName,"
                            sSql = sSql & "SubScope,"
                            sSql = sSql & "SubType"
                            sSql = sSql & ") values("
                            sSql = sSql & 1 & ","
                            sSql = sSql & rst!FileID & ","
                            sSql = sSql & 1 & ","
                            sSql = sSql & SQLCheck(sSub) & ","
                            sSql = sSql & SQLCheck(sScope) & ","
                            sSql = sSql & SQLCheck(sType) & ""
                            sSql = sSql & ")"
                        
                            Call SQLExecute(gdbFL, sSql)
                        End If
                    End If
                            
                    GoTo NextStep
WriteMess:

                    sSql = "select * from Messages where "
                    sSql = sSql & " FileID=" & rst!FileID
                    sSql = sSql & " and MessageTxt=" & SQLCheck(sMess)
                    Set rstCheck = SQLOpenRecordset(gdbFL, sSql)
                    If Not rstCheck Is Nothing Then
                        If rstCheck.EOF Then
                            ' INSERT STATEMENT
                            sSql = "insert into Messages ("
                            sSql = sSql & "Checked,"
                            sSql = sSql & "FileID,"
                            sSql = sSql & "MessageTxt,"
                            sSql = sSql & "StatusID"
                            sSql = sSql & ") values("
                            sSql = sSql & 1 & ","
                            sSql = sSql & rst!FileID & ","
                            sSql = sSql & SQLCheck(sMess) & ","
                            sSql = sSql & 1 & ""
                            sSql = sSql & ")"
                        
                            Call SQLExecute(gdbFL, sSql)
                        End If
                    End If
NextStep:
                Loop
                Close #iFile
            End If
            rst.MoveNext
        Loop
        dcSubs.Refresh
        sb.Panels(miS_FILE).Text = "File:" & "FINISHED"
        sb.Panels(miS_SUB).Text = "Sub:" & ""
        
    End Select
    
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "cmd_Click")
End Sub

Private Sub cmdCallF_Click()
On Error GoTo ErrHandler
    txtPath = OpenDirectoryTV(Me, "Please, select Path.")

ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "cmdCallF_Click")
End Sub

Private Sub cmdFilePath_Click()
On Error GoTo ErrHandler

    cd.Flags = cdlOFNHideReadOnly   ' Set filters
    cd.CancelError = True
    cd.Filter = "All Files (*.*)|*.*"
    cd.FilterIndex = 1              ' Display the Open dialog box
'    cd.FileName = ms_FILE_NAME      ' Display file name of selected file
    cd.ShowSave                     ' Display name of selected file
    txtFilePath = cd.FileName

ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "cmdFilePath_Click")
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    Dim sSql As String
' SELECT DISTINCTROW FileList.FileID, FileList.FilePath, FileList.FileName, Subs.SubID, Subs.SubName, Subs.SubType, Subs.SubScope, Subs.StatusID, Subs.Checked FROM FileList INNER JOIN Subs ON FileList.FileID =Subs.FileID
    Screen.MousePointer = vbHourglass
    
    sSql = "SELECT DISTINCTROW FileList.FileID,"
    sSql = sSql & "        FileList.FilePath,"
    sSql = sSql & "        FileList.FileName,"
    sSql = sSql & "        Subs.SubID,"
    sSql = sSql & "        Subs.SubName,"
    sSql = sSql & "        Subs.SubType,"
    sSql = sSql & "        Subs.SubScope,"
    sSql = sSql & "        Subs.StatusID,"
    sSql = sSql & "        Subs.Checked"
    sSql = sSql & " FROM FileList"
    sSql = sSql & " INNER JOIN Subs"
    sSql = sSql & " ON FileList.FileID =Subs.FileID"
    msSql = sSql
    
    txtPath = ""
    txtFile = ""
    txtFilePath = ""
    txtSub = ""
    sb.Panels(miS_SUB).Text = "Sub:" & ""
    sb.Panels(miS_FILE).Text = "File:" & ""
    sb.Panels(miS_TOTAL).Text = "Total:" & 0
    
    txtFilePath = GetSetting("FL", "FllePaths", "LastFllePathSub", "")
    
    msSqlSel = " where FileList.FileID>=0 "
    msSqlOrder = " order by FilePath, FileName "
    dcSubs.RecordSource = msSql & msSqlSel & msSqlOrder
    dcSubs.DatabaseName = gsDBpath
    dcSubs.Refresh
    Screen.MousePointer = vbDefault

ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "Form_Load")
End Sub


Private Sub grdSubs_DblClick()
On Error GoTo ErrHandler
    cmd_Click (4)
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "grdSubs_DblClick")
End Sub

Private Sub grdSubs_HeadClick(ByVal ColIndex As Integer)
On Error GoTo ErrHandler
    Dim sSql As String, rst As DAO.Recordset
    Select Case ColIndex
    Case 0:         msSqlOrder = " Order by FilePath "
    Case 1:         msSqlOrder = " Order by FileName "
    Case 2:         msSqlOrder = " Order by SubName "
    Case 3:         msSqlOrder = " Order by SubType "
    Case 4:         msSqlOrder = " Order by SubScope "
    Case Else:      msSqlOrder = " Order by FileName "
    End Select
    If grdSubs.Tag <> " ASC" Then
        grdSubs.Tag = " ASC"
    Else
        grdSubs.Tag = " DESC"
    End If
    Me.MousePointer = vbHourglass
    
    sSql = msSql & msSqlSel & msSqlSelView & msSqlOrder & grdSubs.Tag
    dcSubs.RecordSource = sSql
    dcSubs.Refresh
    
    sSql = "select Count(SubID) as Num "
    sSql = sSql & " FROM FileList"
    sSql = sSql & " INNER JOIN Subs"
    sSql = sSql & " ON FileList.FileID =Subs.FileID"
    sSql = sSql & msSqlSel & msSqlSelView
    Set rst = SQLOpenRecordset(gdbFL, sSql)
    
    sb.Panels(miS_TOTAL).Text = "Total:" & rst!Num
    DoEvents
    Me.MousePointer = vbDefault
    
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "grdSubs_HeadClick")
End Sub

Private Sub txtFile_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandler
    If KeyAscii = 13 Then
        RefreshData
    End If
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "txtFile_KeyPress")
End Sub

Private Sub txtPath_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandler
    If KeyAscii = 13 Then
        RefreshData
    End If
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "txtPath_KeyPress")
End Sub

Private Sub txtSub_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandler
    If KeyAscii = 13 Then
        RefreshData
    End If
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "txtSub_KeyPress")
End Sub
Private Sub RefreshData()
On Error GoTo ErrHandler
    Dim sSql As String, rst As DAO.Recordset
    
    Screen.MousePointer = vbHourglass

    msSqlOrder = " Order by FilePath ASC,FileName ASC, SubName ASC"
    msSqlSel = " Where FilePath like '" & txtPath & "*'"
    If txtFile <> "" Then msSqlSel = msSqlSel & " and FileName like '" & txtFile & "*'"
    If txtSub <> "" Then msSqlSel = msSqlSel & " and SubName like '" & txtSub & "*'"
    sSql = msSql & msSqlSel & msSqlSelView & msSqlOrder
    dcSubs.RecordSource = sSql
    dcSubs.Refresh
    
    sSql = "select Count(SubID) as Num "
    sSql = sSql & " FROM FileList"
    sSql = sSql & " INNER JOIN Subs"
    sSql = sSql & " ON FileList.FileID =Subs.FileID"
    sSql = sSql & msSqlSel & msSqlSelView
    Set rst = SQLOpenRecordset(gdbFL, sSql)
    
    sb.Panels(miS_TOTAL).Text = "Total:" & rst!Num
    DoEvents
    
    Screen.MousePointer = vbDefault

ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "RefreshData")
End Sub

