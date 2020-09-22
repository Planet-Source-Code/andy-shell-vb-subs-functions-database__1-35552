VERSION 5.00
Begin VB.Form frmMove 
   Caption         =   "Move"
   ClientHeight    =   1425
   ClientLeft      =   7140
   ClientTop       =   4680
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   6330
   Begin VB.CommandButton cmd 
      Caption         =   "Move"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      ToolTipText     =   "Move File/Folder"
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdCallF 
      Caption         =   "..."
      Height          =   285
      Left            =   5880
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Text            =   "txtPath"
      Top             =   480
      Width           =   5175
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Text            =   "txtPath"
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "To:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "From :"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "frmMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private miMoveFile       As Integer

Public Property Get MoveFile() As Integer
On Error GoTo ErrHandler

        MoveFile = miMoveFile

ErrExit:         Exit Property
ErrHandler:      Call ErrorHandler(Name, 0, "MoveFile (Property Get)")
End Property

Public Property Let MoveFile(iMoveFile As Integer)
On Error GoTo ErrHandler

       miMoveFile = iMoveFile

ErrExit:         Exit Property
ErrHandler:      Call ErrorHandler(Name, 0, "MoveFile (Property Let)")
End Property


Private Sub cmd_Click()
On Error GoTo ErrHandler
    Call SaveSetting("SysFiles", "LastPath", "LastPath", txtPath(1))
    If MoveFile = 1 Then
        If Dir(txtPath(0)) <> "" Then
            MoveFileFSO CheckPath(txtPath(0)), CheckPath(txtPath(1))
        Else
            MsgBox "Source File was not found"
        End If
    Else
        If Dir(txtPath(0), vbDirectory) <> "" Then
            If Dir(txtPath(1), vbDirectory) <> "" Then
                MsgBox "Destination Directory already exists. Use Copy."
                Exit Sub
            End If
            MoveFolderFSO CheckPath(txtPath(0)), CheckPath(txtPath(1))
        Else
            MsgBox "Source Folder was not found"
        End If
    End If
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "cmd_Click")
End Sub

Private Sub cmdCallF_Click()
On Error GoTo ErrHandler

    txtPath(1) = OpenDirectoryTV(Me, "Please, select Path.")
    txtPath(1) = txtPath(1) & GetLastFolderFromPath(txtPath(0))
    
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "cmdCallF_Click")
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    txtPath(0) = ""
    txtPath(1) = GetSetting("SysFiles", "LastPath", "LastPath", "")
    
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler(Name, 0, "Form_Load")
End Sub
