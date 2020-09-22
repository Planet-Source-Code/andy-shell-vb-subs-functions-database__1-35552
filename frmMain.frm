VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "File List"
   ClientHeight    =   7020
   ClientLeft      =   2460
   ClientTop       =   1545
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   11445
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   6645
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Folder:"
            TextSave        =   "Folder:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "File:"
            TextSave        =   "File:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Total:"
            TextSave        =   "Total:"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Help"
      Height          =   375
      Index           =   11
      Left            =   960
      TabIndex        =   21
      ToolTipText     =   "Open Help File"
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Vote"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   20
      ToolTipText     =   "Vote for this Application"
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Subs"
      Height          =   375
      Index           =   10
      Left            =   1920
      TabIndex        =   19
      ToolTipText     =   "Open Subs/Functions List"
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "New List"
      Height          =   375
      Index           =   9
      Left            =   3840
      TabIndex        =   18
      ToolTipText     =   "Open New File List Form"
      Top             =   6240
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboView 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   10320
      List            =   "frmMain.frx":000D
      TabIndex        =   17
      Text            =   "ALL"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Print"
      Height          =   375
      Index           =   8
      Left            =   2880
      TabIndex        =   15
      ToolTipText     =   "Print Selected List"
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Copy File"
      Height          =   375
      Index           =   6
      Left            =   5760
      TabIndex        =   14
      ToolTipText     =   "Copies Selected File "
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Copy Dir"
      Height          =   375
      Index           =   7
      Left            =   4800
      TabIndex        =   13
      ToolTipText     =   "Copies Selected Folder"
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Move Dir"
      Height          =   375
      Index           =   5
      Left            =   6720
      TabIndex        =   12
      ToolTipText     =   "Moves Selected Folder"
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Move File"
      Height          =   375
      Index           =   4
      Left            =   7680
      TabIndex        =   11
      ToolTipText     =   "Moves Selected File "
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox txtFileExt 
      Height          =   285
      Left            =   9240
      TabIndex        =   10
      Text            =   "txtFileExt"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   5640
      TabIndex        =   9
      Text            =   "txtFile"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Timer tmrNewRun 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   5760
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Clean DB"
      Height          =   375
      Index           =   3
      Left            =   8640
      TabIndex        =   6
      ToolTipText     =   "Deletes ALL Entries in Database"
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmdCallF 
      Caption         =   "..."
      Height          =   285
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Delete"
      Height          =   375
      Index           =   2
      Left            =   9600
      TabIndex        =   4
      ToolTipText     =   "Deletes Selected File/Folder"
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Text            =   "txtPath"
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Check"
      Height          =   375
      Index           =   1
      Left            =   10560
      TabIndex        =   1
      ToolTipText     =   "Check Selected Path"
      Top             =   6240
      Width           =   855
   End
   Begin VB.Data dcFiles 
      Caption         =   "dcFiles"
      Connect         =   "Access"
      DatabaseName    =   "D:\0_FileLib\Data\SystemFileList.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "FileList"
      Top             =   5880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSDBGrid.DBGrid grdFiles 
      Bindings        =   "frmMain.frx":0026
      Height          =   5655
      Left            =   0
      OleObjectBlob   =   "frmMain.frx":003C
      TabIndex        =   0
      Top             =   480
      Width           =   11415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "View:"
      Height          =   195
      Index           =   6
      Left            =   9840
      TabIndex        =   16
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Ext.:"
      Height          =   195
      Index           =   4
      Left            =   8640
      TabIndex        =   8
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File:"
      Height          =   195
      Index           =   3
      Left            =   5280
      TabIndex        =   7
      Top             =   120
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Path:"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlFilesNum              As Long
Private msSql                   As String
Private msSqlSel                As String
Private msSqlSelView            As String
Private msSqlOrder              As String
Private miFileNum               As Integer
Private msFilePath              As String
Private Const ms_FILE_NAME      As String = "FileList01.txt"

Private Const miS_FOLDER        As Integer = 1
Private Const miS_FILE          As Integer = 2
Private Const miS_TOTAL         As Integer = 3

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cboView_Click()
On Error GoTo ErrHandler
    Dim sSql As String, rst As DAO.Recordset
    
    CheckView
    sSql = msSql & msSqlSel & msSqlSelView & msSqlOrder
    dcFiles.RecordSource = sSql
    dcFiles.Refresh
    
    sSql = "select Count(FileID) as Num from FileList "
    sSql = sSql & msSqlSel & msSqlSelView
    Set rst = SQLOpenRecordset(gdbFL, sSql)
    
    sb.Panels(miS_TOTAL).Text = "Total:" & rst!Num
    DoEvents
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "cboView_Click")
End Sub

Private Sub cmd_Click(Index As Integer)
On Error GoTo ErrHandler
    Dim s           As String
    Dim sSql        As String
    Dim rst         As DAO.Recordset
    Dim sFilePath   As String
    Dim i           As Integer
    Dim l           As Long
    Dim lFileID     As Long
    Dim frm         As Form
    
    Const iB_V                  As Integer = 0
    Const iB_CHECK_SELECTED     As Integer = 1
    Const iB_DELETE_SELECTED    As Integer = 2
    Const iB_CLEAN_DB           As Integer = 3
    Const iB_MOVE_FILE          As Integer = 4
    Const iB_MOVE_FOLDER        As Integer = 5
    Const iB_COPY_FILE          As Integer = 6
    Const iB_COPY_FOLDER        As Integer = 7
    Const iB_PRINT              As Integer = 8
    Const iB_NEW_LIST           As Integer = 9
    Const iB_SUBS               As Integer = 10
    Const iB_HELP               As Integer = 11
    
    Select Case Index
    Case iB_V:
        s = "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=35552&lngWId=1"
        ShellExecute hWnd, "Open", s, 0&, "C:\", 1
    
    Case iB_SUBS:       frmSubs.Show
    Case iB_HELP:
        l = Shell("notepad.exe " & App.Path & "\ReadMe.txt", vbNormalFocus)
        
    Case iB_NEW_LIST:
        Set frm = New frmMain
        With frm
            .Show
            .Top = Top + 400
            .Left = Left + 400
        End With
        
    Case iB_PRINT:
        cd.Flags = cdlOFNHideReadOnly   ' Set filters
        cd.CancelError = True
        cd.Filter = "All Files (*.*)|*.*"
        cd.FilterIndex = 1              ' Display the Open dialog box
        cd.FileName = ms_FILE_NAME      ' Display file name of selected file
        cd.ShowSave                     ' Display name of selected file
        msFilePath = cd.FileName
        
        miFileNum = FreeFile
        Open msFilePath For Output As #miFileNum
        
        sSql = msSql & msSqlSel & msSqlSelView & msSqlOrder
        
        Set rst = SQLOpenRecordset(gdbFL, sSql)
        Print #miFileNum, "File Path", , , "File Size", "File Version", "File Create Date"
        Print #miFileNum, "========================================================================="
        Do Until rst.EOF
            Print #miFileNum, rst!FilePath & "\" & rst!FileName, rst!FileSize, rst!FileVersion, rst!FileCreateDate
            rst.MoveNext
        Loop
        Print #miFileNum, "========================================================================="
        Close #miFileNum
        
        sSql = "notepad.exe " & msFilePath
        l = Shell(sSql, vbNormalFocus)
    
    
    Case iB_COPY_FILE:
        If dcFiles.Recordset!FileName = "" Then
            MsgBox "It is a Folder. Not a file."
            Exit Sub
        End If
        With frmCopy
            .Show
            .CopyFile = 1
            .txtPath(0) = CheckPath(dcFiles.Recordset!FilePath & "\" & dcFiles.Recordset!FileName)
        End With
        
    Case iB_COPY_FOLDER:
        With frmCopy
            .Show
            .CopyFile = 0
            .txtPath(0) = dcFiles.Recordset!FilePath
        End With
        
    Case iB_MOVE_FILE:
        If dcFiles.Recordset!FileName = "" Then
            MsgBox "It is a Folder. Not a file."
            Exit Sub
        End If
        With frmMove
            .Show
            .MoveFile = 1
            .txtPath(0) = CheckPath(dcFiles.Recordset!FilePath & "\" & dcFiles.Recordset!FileName)
        End With
        
    Case iB_MOVE_FOLDER:
        With frmMove
            .Show
            .MoveFile = 0
            .txtPath(0) = dcFiles.Recordset!FilePath
        End With
        
    Case iB_CLEAN_DB:
        i = MsgBox("Do you want to delete ALL records from Database?", vbYesNo)
        If i <> vbYes Then Exit Sub
        sSql = "delete from FileList"
        Call SQLExecute(gdbFL, sSql)
        dcFiles.Refresh
        
    Case iB_CHECK_SELECTED:
        mlFilesNum = 0
        If txtPath = "" Or Len(txtPath) = 1 Then
            MsgBox "Please Enter Drive\Folder to be checked and try again."
            Exit Sub
        End If
        txtPath = CheckPath(txtPath)
        If Right(Trim(txtPath), 1) <> "\" Then txtPath = txtPath & "\"
        
        sSql = "update FileList set StatusID=0, Checked=0 where FilePath like " & SQLCheck(txtPath & "*")
        Call SQLExecute(gdbFL, sSql)
        
        Call CheckSelected(txtPath)
        
    Case iB_DELETE_SELECTED
        i = MsgBox("Do you want to delete selected file?", vbYesNo)
        If i <> vbYes Then Exit Sub
        sFilePath = dcFiles.Recordset!FilePath & "\" & dcFiles.Recordset!FileName
        sFilePath = CheckPath(sFilePath)
        If Dir(sFilePath) <> "" Then Kill sFilePath
        
        sSql = "delete from FileList "
        sSql = sSql & " where FIleID=" & dcFiles.Recordset!FileID
        Call SQLExecute(gdbFL, sSql)
        dcFiles.Recordset.MoveNext
        If Not dcFiles.Recordset.EOF Then lFileID = dcFiles.Recordset!FileID
        dcFiles.Refresh
        dcFiles.Recordset.FindFirst "FileID=" & lFileID
        
    End Select
    dcFiles.Refresh
    
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "cmd_Click")
End Sub

Private Sub cmdCallF_Click()
On Error GoTo ErrHandler
    txtPath = OpenDirectoryTV(Me, "Please, select Path.")

ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "cmdCallF_Click")
End Sub

Private Sub dcFiles_Reposition()
On Error GoTo ErrHandler
    If dcFiles.Recordset.EOF Then
        CmdCheck False
    Else
        CmdCheck True
    End If
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "dcFiles_Reposition")
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    Dim i As Integer
    
    txtPath = ""
    txtFile = ""
    txtFileExt = ""
    
    sb.Panels(miS_FOLDER).Text = "Folder:"
    
    sb.Panels(miS_FILE).Text = "File:"
    
    sb.Panels(miS_TOTAL).Text = "Total:"
    
    mlFilesNum = 0
    gsDBpath = App.Path & "\" & "SystemFileList.mdb"
    If Dir(gsDBpath) = "" Then Call Create_Database(gsDBpath)
    msSql = "select * from FileList "
    msSqlSel = " where FileID>=0 "
    
    Set gdbFL = DBEngine.OpenDatabase(gsDBpath, False, False)
    
    dcFiles.DatabaseName = gsDBpath
    dcFiles.Refresh
    
    If dcFiles.Recordset.EOF Then
        CmdCheck False
    Else
        CmdCheck True
    End If
    If Dir(gs_LAST_FILE_LOG) <> "" Then Kill gs_LAST_FILE_LOG
    
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "Form_Load")
End Sub

Private Sub CmdCheck(bShow As Boolean)
On Error GoTo ErrHandler
    Dim i As Integer
    For i = 2 To 10
        cmd(i).Enabled = bShow
    Next

ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "CmdCheck")
End Sub

Private Sub CheckSelected(sFolderPath As String)
On Error GoTo ErrHandler
    Dim sSql        As String
    Dim rst         As Recordset
    Dim sStr()      As String
    Dim iFolders    As Integer
    Dim iFiles      As Integer
    Dim sFileName   As String
    Dim sFileSuffix As String
    Dim i           As Integer
    Dim s           As String
    Dim sPath       As String
    Dim sVersion    As String
    Dim iLevel      As Integer
    Dim lSize       As Double
    Dim dCreateDate As Date
    Dim dModDate    As Date
    Dim dAccessDate As Date
    
    sPath = sFolderPath
    sStr = Split(GetFolderList(sFolderPath, "||"), "||")
    
    For iFolders = 0 To UBound(sStr)
        s = sFolderPath & "\" & sStr(iFolders)
        s = CheckPath(s)
        iLevel = GetFolderLevelDepthFSO(s)
        lSize = GetFolderSize(s)
        dCreateDate = GetFolderDateCreate(s)
        dModDate = GetFolderDateModified(s)
        dAccessDate = GetFolderDateAccessed(s)
        sb.Panels(miS_FOLDER).Text = "Folder:" & s & " Size " & lSize

        mlFilesNum = mlFilesNum + 1
        sb.Panels(miS_TOTAL).Text = "Total:" & mlFilesNum
        DoEvents
        Call SaveFolder(s _
                        , dCreateDate _
                        , dModDate _
                        , dAccessDate _
                        , lSize _
                        , iLevel _
                        )
    Next
    
    sStr = Split(GetFileList(sFolderPath, "||"), "||")
    

    For iFiles = 0 To UBound(sStr)
        If Trim(sStr(iFiles)) = "" Then GoTo NextFile
        s = sFolderPath & "\" & sStr(iFiles)
        s = CheckPath(s)
        
        sb.Panels(miS_FOLDER).Text = "Folder:" & sFolderPath

        sb.Panels(miS_FILE).Text = "File:" & sStr(iFiles)
        
        mlFilesNum = mlFilesNum + 1
        
        sb.Panels(miS_TOTAL).Text = "Total:" & mlFilesNum
        
        DoEvents
        
        sVersion = GetFileVersion(s)
        lSize = GetFileSize(s)
        sFileSuffix = GetFileExt(s)
        dCreateDate = GetFileDateCreate(s)
        dModDate = GetFileDateModified(s)
        dAccessDate = GetFileDateAccessed(s)
        
        sb.Panels(miS_FILE).Text = "File:" & sStr(iFiles) & " Size " & lSize
        
        DoEvents
        Call SaveFile(sStr(iFiles) _
                        , sFolderPath _
                        , sFileSuffix _
                        , dCreateDate _
                        , dModDate _
                        , dAccessDate _
                        , sVersion _
                        , lSize _
                        , iLevel _
                        )
NextFile:
    Next
    
    UpdateFolderChecked sFolderPath
    
    sb.Panels(miS_FOLDER).Text = "Folder: FINISHED"
    
    sb.Panels(miS_FILE).Text = "File:"
    
    DoEvents
    tmrNewRun.Enabled = True

ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "CheckSelected")
End Sub

Private Sub SaveFolder(sFolderPath As String _
                        , dCreateDate As Date _
                        , dModDate As Date _
                        , dAccessDate As Date _
                        , lSize As Double _
                        , iLevel As Integer _
                        )
On Error GoTo ErrHandler
    Dim sSql As String, rst As DAO.Recordset
    
    If Right(sFolderPath, 1) = "\" Then
        sFolderPath = Left(sFolderPath, Len(sFolderPath) - 1)
    End If
    
    sSql = "select * from FileList where "
    sSql = sSql & " FilePath=" & SQLCheck(sFolderPath)
    Set rst = SQLOpenRecordset(gdbFL, sSql)
        
    If rst.EOF Then
        ' INSERT STATEMENT
        sSql = "insert into FileList ("
        sSql = sSql & "FilePath,"
        sSql = sSql & "FileAccessDate,"
        sSql = sSql & "FileCreateDate,"
        sSql = sSql & "FileLevel,"
        sSql = sSql & "FileModifyDate,"
        sSql = sSql & "FileSize,"
        sSql = sSql & "Checked,"
        sSql = sSql & "FileType,"
        sSql = sSql & "StatusID"
        sSql = sSql & ") values("
        sSql = sSql & SQLCheck(LCase(sFolderPath)) & ","
        sSql = sSql & AMDateTime(dAccessDate) & ","
        sSql = sSql & AMDateTime(dCreateDate) & ","
        sSql = sSql & iLevel & ","
        sSql = sSql & AMDateTime(dModDate) & ","
        sSql = sSql & lSize & ","
        sSql = sSql & 0 & ","
        sSql = sSql & 0 & ","
        sSql = sSql & 1 & ""
        sSql = sSql & ")"
    
        Call SQLExecute(gdbFL, sSql)
    Else
        If rst!StatusID = 1 Then Exit Sub
        If rst!FileModifyDate <> dModDate Or rst!FileSize <> lSize Then
            Call UpdateFolder(rst!FileID _
                            , dCreateDate _
                            , dModDate _
                            , dAccessDate _
                            , lSize _
                            , iLevel _
                            )
        End If
    End If
    Set rst = Nothing

ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "SaveFolder")
End Sub

Private Sub UpdateFolder(lFolderID As Long _
                        , dCreateDate As Date _
                        , dModDate As Date _
                        , dAccessDate As Date _
                        , lSize As Double _
                        , iLevel As Integer _
                        )
On Error GoTo ErrHandler
    Dim sSql As String
    
    sSql = "Update FileList set "
    sSql = sSql & " FileAccessDate=" & AMDateTime(dAccessDate)
    sSql = sSql & ", FileCreateDate=" & AMDateTime(dCreateDate)
    sSql = sSql & ", FileLevel=" & iLevel
    sSql = sSql & ", FileModifyDate=" & AMDateTime(dModDate)
    sSql = sSql & ", FileSize=" & lSize
    sSql = sSql & ", Checked=1"
    sSql = sSql & ", StatusID=1"
    sSql = sSql & " where FileID=" & lFolderID
    
    Call SQLExecute(gdbFL, sSql)

ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "UpdateFolder")
End Sub

Private Sub UpdateFile(lFileID As Long _
                        , dCreateDate As Date _
                        , dModDate As Date _
                        , dAccessDate As Date _
                        , sVersion As String _
                        , lSize As Double _
                        , iLevel As Integer _
                        )
On Error GoTo ErrHandler
    Dim sSql As String
    
    If sVersion = "" Then sVersion = " "
    sSql = "Update FileList set "
    sSql = sSql & " FileAccessDate=" & AMDateTime(dAccessDate)
    sSql = sSql & ", FileCreateDate=" & AMDateTime(dCreateDate)
    sSql = sSql & ", FileLevel=" & iLevel
    sSql = sSql & ", FileModifyDate=" & AMDateTime(dModDate)
    sSql = sSql & ", FileVersion=" & SQLCheck(sVersion)
    sSql = sSql & ", FileSize=" & lSize
    sSql = sSql & ", Checked=1"
    sSql = sSql & ", StatusID=1"
    sSql = sSql & " where FileID=" & lFileID
    
    Call SQLExecute(gdbFL, sSql)

ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "UpdateFile")
End Sub


Private Sub UpdateFolderChecked(sFolderPath As String)
On Error GoTo ErrHandler
    Dim sSql As String
    
    sSql = "update FileList set StatusID=1, Checked=1"
    sSql = sSql & " where FilePath=" & SQLCheck(sFolderPath)
    sSql = sSql & " and FileType =0"

    Call SQLExecute(gdbFL, sSql)
    
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "UpdateFolderChecked")
End Sub

Private Sub SaveFile(sFileName As String _
                        , sFolderPath As String _
                        , sFileSuffix As String _
                        , dCreateDate As Date _
                        , dModDate As Date _
                        , dAccessDate As Date _
                        , sVersion As String _
                        , lSize As Double _
                        , iLevel As Integer _
                        )
On Error GoTo ErrHandler
    Dim sSql As String, rst As DAO.Recordset
    If Right(Trim(sFileName), 1) = "\" Then Exit Sub
    If Right(Trim(sFileName), 1) = "/" Then Exit Sub
    
    If Right(sFolderPath, 1) = "\" Then sFolderPath = Left(sFolderPath, Len(sFolderPath) - 1)
    
    If sVersion = "" Then sVersion = " "
    sSql = "select * from FileList where "
    sSql = sSql & " FileName=" & SQLCheck(sFileName)
    sSql = sSql & " and FilePath=" & SQLCheck(sFolderPath)
    Set rst = SQLOpenRecordset(gdbFL, sSql)
    
    If rst.EOF Then
        ' INSERT STATEMENT
        sSql = "insert into FileList ("
        sSql = sSql & "FileAccessDate,"
        sSql = sSql & "FileCreateDate,"
        sSql = sSql & "FileLevel,"
        sSql = sSql & "FileModifyDate,"
        sSql = sSql & "FileName,"
        sSql = sSql & "FilePath,"
        sSql = sSql & "FileSize,"
        sSql = sSql & "FileVersion,"
        sSql = sSql & "FileSuffix,"
        sSql = sSql & "FileType,"
        sSql = sSql & "Checked,"
        sSql = sSql & "StatusID"
        sSql = sSql & ") values("
        sSql = sSql & AMDateTime(dAccessDate) & ","
        sSql = sSql & AMDateTime(dCreateDate) & ","
        sSql = sSql & iLevel & ","
        sSql = sSql & AMDateTime(dModDate) & ","
        sSql = sSql & SQLCheck(LCase(sFileName)) & ","
        sSql = sSql & SQLCheck(LCase(sFolderPath)) & ","
        sSql = sSql & lSize & ","
        sSql = sSql & SQLCheck(sVersion) & ","
        sSql = sSql & SQLCheck(LCase(sFileSuffix)) & ","
        sSql = sSql & 1 & ","
        sSql = sSql & 1 & ","
        sSql = sSql & 1 & ""
        sSql = sSql & ")"
    
        Call SQLExecute(gdbFL, sSql)
    Else
        If rst!StatusID = 1 Then Exit Sub
        Call UpdateFile(rst!FileID _
                        , dCreateDate _
                        , dModDate _
                        , dAccessDate _
                        , sVersion _
                        , lSize _
                        , iLevel _
                        )
    End If
    Set rst = Nothing

ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "SaveFile")
End Sub

Private Sub Form_Terminate()
On Error GoTo ErrHandler
    If Forms.Count <= 0 Then
        gdbFL.Close
        Set gdbFL = Nothing
        End
    End If
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "Form_Terminate")
End Sub

Private Sub grdFiles_DblClick()
On Error GoTo ErrHandler
    Dim l           As Long
    Dim sFilePath   As String
    
    If Not dcFiles.Recordset.EOF Then
        Clipboard.SetText dcFiles.Recordset!FilePath
        sFilePath = dcFiles.Recordset!FilePath & "\" & dcFiles.Recordset!FileName
        If Dir(sFilePath) <> "" Then
            l = Shell("notepad.exe " & sFilePath, vbNormalFocus)
        End If
    End If
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "grdFiles_DblClick")
End Sub

Private Sub grdFiles_HeadClick(ByVal ColIndex As Integer)
On Error GoTo ErrHandler
    
    Dim sSql As String, rst As DAO.Recordset
    
    Screen.MousePointer = vbHourglass
    
    Select Case ColIndex
    Case 0:         msSqlOrder = " Order by FilePath "
    Case 1:         msSqlOrder = " Order by FileName "
    Case 2:         msSqlOrder = " Order by FileSuffix "
    Case 3:         msSqlOrder = " Order by FileLevel "
    Case 4:         msSqlOrder = " Order by FileSize "
    Case 5:         msSqlOrder = " Order by FileVersion "
    Case Else:      msSqlOrder = " Order by FilePath "
    End Select
    If grdFiles.Tag <> " ASC" Then
        grdFiles.Tag = " ASC"
    Else
        grdFiles.Tag = " DESC"
    End If
    
    sSql = msSql & msSqlSel & msSqlSelView & msSqlOrder & grdFiles.Tag
    dcFiles.RecordSource = sSql
    dcFiles.Refresh
    
    sSql = "select Count(FileID) as Num from FileList "
    sSql = sSql & msSqlSel & msSqlSelView
    Set rst = SQLOpenRecordset(gdbFL, sSql)
    
    sb.Panels(miS_TOTAL).Text = "Total:" & rst!Num
    
    DoEvents
    Screen.MousePointer = vbDefault
    
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "grdFiles_HeadClick")
End Sub


Private Sub tmrNewRun_Timer()
On Error GoTo ErrHandler
    Dim sSql As String, rst As DAO.Recordset
    tmrNewRun.Enabled = False
    
    sSql = "select * from FileList "
    sSql = sSql & " where Checked=0"
    sSql = sSql & " and FileType=0"
    Set rst = SQLOpenRecordset(gdbFL, sSql)
    If rst.EOF Then Exit Sub
    Do Until rst.EOF
        CheckSelected rst!FilePath
        rst.MoveNext
    Loop
    Set rst = Nothing
    sSql = "delete from FileList where StatusID=0"
    Call SQLExecute(gdbFL, sSql)
    
    dcFiles.Refresh
    sb.Panels(miS_FOLDER).Text = "Folder: FINISHED"
    
    sb.Panels(miS_FILE).Text = "File:"
    
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "tmrNewRun_Timer")
End Sub


Private Sub txtFile_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandler
    If KeyAscii = 13 Then
        RefreshData
    End If
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "txtFile_KeyPress")
End Sub

Private Sub txtFileExt_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandler
    If KeyAscii = 13 Then
        RefreshData
    End If
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "txtFileExt_KeyPress")
End Sub

Private Sub txtPath_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandler
    If KeyAscii = 13 Then
        RefreshData
    End If
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "txtPath_KeyPress")
End Sub

Private Sub CheckView()
On Error GoTo ErrHandler
    Select Case cboView
    Case "ALL":         msSqlSelView = ""
    Case "Files":       msSqlSelView = " and FileType=1"
    Case "Folders":     msSqlSelView = " and FileType=0"
    Case Else:          msSqlSelView = ""
    End Select
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "CheckView")
End Sub

Private Sub RefreshData()
On Error GoTo ErrHandler
    Dim sSql As String, rst As DAO.Recordset
    
    Screen.MousePointer = vbHourglass
    
    msSqlOrder = " Order by FilePath ASC,FileName ASC, FileSuffix ASC"
    msSqlSel = " Where FileID>0 "
    If Trim(txtPath) <> "" Then msSqlSel = msSqlSel & " and FilePath like '" & txtPath & "*'"
    If Trim(txtFile) <> "" Then msSqlSel = msSqlSel & " and FileName like '" & txtFile & "*'"
    If Trim(txtFileExt) <> "" Then msSqlSel = msSqlSel & " and FileSuffix like '" & txtFileExt & "*'"
    sSql = msSql & msSqlSel & msSqlSelView & msSqlOrder
    dcFiles.RecordSource = sSql
    dcFiles.Refresh
    
    sSql = "select Count(FileID) as Num from FileList "
    sSql = sSql & msSqlSel & msSqlSelView
    Set rst = SQLOpenRecordset(gdbFL, sSql)
    
    sb.Panels(miS_TOTAL).Text = "Total: " & rst!Num
    
    DoEvents
    Screen.MousePointer = vbDefault

ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "RefreshData")
End Sub
