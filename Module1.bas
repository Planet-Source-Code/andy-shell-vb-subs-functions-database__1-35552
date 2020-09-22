Attribute VB_Name = "Module1"
Option Explicit


Public gsPathLogs   As String

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
    sPrintLine = SystemVersion & " - " & gsUserName & " - " & gsCompName & " - " & sFormName & "  -  " & SubName & " - Section - " & CStr(ind) & " - " & sErrDesc & " Error # " & lErrNum & "  " & sErrSource & " " & sSql
    
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


