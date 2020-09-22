Attribute VB_Name = "modFileList"
Option Explicit

Public gdbFL                    As DAO.Database
Public gsDBpath                 As String
Private gsPathLogs              As String
Public Const gs_LAST_FILE_LOG   As String = "c:\LastFile.txt"


Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
'Private Const MAX_PATH = 260
Private Const MAX_PATH As Long = 260

Private Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Public Function OpenDirectoryTV(odtvOwner As Form, Optional odtvTitle As String) As String
On Error GoTo ErrHandler
   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo
   szTitle = odtvTitle
   With tBrowseInfo
      .hwndOwner = odtvOwner.hWnd
      .lpszTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
   End With
   lpIDList = SHBrowseForFolder(tBrowseInfo)
   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      OpenDirectoryTV = sBuffer
   End If
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFileList", 0, "OpenDirectoryTV")
End Function

'=====================================================================
'Description:       Converts Date/Time into American Format
'Parameters         d           - Date
'                   dbt         - Database Format
'
Public Function AMDateTime(ByVal d As Date) As String
On Error GoTo ErrHandler
    AMDateTime = "#" & Month(d) & "/" & Day(d) & "/" & Year(d) & " " & TimeValue(d) & "#"
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFileList", 0, "AMDateTime")
End Function


Public Sub SQLExecute(DB As DAO.Database, sSql As String)
On Error GoTo ErrHandler
    DB.Execute sSql, dbSeeChanges
ErrExit:      Exit Sub
ErrHandler:   Call ErrorHandler("modFileList", 0, "SQLExecute", sSql)
End Sub
Public Function SQLOpenRecordset(DB As DAO.Database, sSql As String) As DAO.Recordset
On Error GoTo ErrHandler
    
    Set SQLOpenRecordset = DB.OpenRecordset(sSql, dbOpenSnapshot)
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFileList", 0, "SQLOpenRecordset")
End Function

Public Function SQLCheck(sStr As String) As String
On Error GoTo ErrHandler
    Dim s As String
    s = Replace(sStr, "'", "''")
    s = Replace(s, "", """")
    SQLCheck = "'" & s & "'"
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFileList", 0, "SQLCheck")
End Function


'========================================================================
'Description:   Error Handler Process and generate Error messages in Log files
'PARAMETERS:    sFormName       - Form Name
'               ind             - Group or line Index variable (used mainly if Subs with Index parameter
'               SubName         - Procedure number withing a Form
'               sSql            - SQl Statement or another String passed from the form
'               bLogOnly        - Put only in Database Log
'
'
Public Sub ErrorHandler(sFormName As String, ind As Integer, SubName As String, Optional sSql As String, Optional bLogOnly As Boolean)
    Dim iFreeFileN As Integer    ' Next Free File Number
    Dim sPrintLine As String     ' Line to be printed
    Dim sLogName As String       ' Log file name
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
                
    ' Show Indicator
    Screen.MousePointer = vbDefault
    sErrDesc = Err.Description
    lErrNum = Err.Number
    sErrSource = Err.Source
    sPrintLine = " FSOModule - " & sFormName & "  -  " & SubName & " - Section - " & CStr(ind) & " - " & sErrDesc & " Error # " & lErrNum & "  " & sErrSource & " " & sSql
    
    'Forming Log File Name and open it
    If Dir(gsPathLogs, vbDirectory) = "" Or gsPathLogs = "" Then gsPathLogs = App.Path & "\"
    sLogName = gsPathLogs & "log" & Year(Date) & Month(Date) & Day(Date) & ".txt"
    
    iFreeFileN = 22
    Open sLogName For Append Shared As #iFreeFileN
    ' Printing Error Message
    Print #iFreeFileN, sPrintLine, Time
    Close #iFreeFileN
    
    If Len(sPrintLine) > 2000 Then sPrintLine = Left(sPrintLine, 2000)
    
    If bLogOnly Then Exit Sub
    
    DoEvents
End Sub

