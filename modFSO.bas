Attribute VB_Name = "modFSO"
Option Explicit
Private gsPathLogs                      As String

Public Const gi_DriveTypeUnknown        As Integer = 0
Public Const gi_DriveTypeRemovable      As Integer = 1
Public Const gi_DriveTypeFixed          As Integer = 2
Public Const gi_DriveTypeNetwork        As Integer = 3
Public Const gi_DriveTypeCDROM          As Integer = 4
Public Const gi_DriveTypeRAMDisk        As Integer = 5

Public Const gs_DriveTypeUnknown        As String = "Unknown"
Public Const gs_DriveTypeRemovable      As String = "Removable (FDD)"
Public Const gs_DriveTypeFixed          As String = "Fixed (HDD)"
Public Const gs_DriveTypeNetwork        As String = "Network"
Public Const gs_DriveTypeCDROM          As String = "CD ROM"
Public Const gs_DriveTypeRAMDisk        As String = "RAM Disk"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Constants returned by File.Attributes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const gi_FileAttrNormal          As Integer = 0
Public Const gi_FileAttrReadOnly        As Integer = 1
Public Const gi_FileAttrHidden          As Integer = 2
Public Const gi_FileAttrSystem          As Integer = 4
Public Const gi_FileAttrVolume          As Integer = 8
Public Const gi_FileAttrDirectory       As Integer = 16
Public Const gi_FileAttrArchive         As Integer = 32
Public Const gi_FileAttrAlias           As Integer = 64
Public Const gi_FileAttrCompressed      As Integer = 128

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Constants for opening files
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const gi_OpenFileForReading      As Integer = 1
Public Const gi_OpenFileForWriting      As Integer = 2
Public Const gi_OpenFileForAppending    As Integer = 8


Private Declare Function GetFileVersionInfo _
  Lib "Version.dll" _
  Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, _
    ByVal dwHandle As Long, _
    ByVal dwLen As Long, _
    lpData As Any) _
As Long

Private Declare Function GetFileVersionInfoSize _
  Lib "Version.dll" _
  Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, _
    lpdwHandle As Long) _
  As Long

Private Declare Function VerQueryValue _
  Lib "Version.dll" _
  Alias "VerQueryValueA" _
  (pBlock As Any, _
    ByVal lpSubBlock As String, _
    lplpBuffer As Any, _
    puLen As Long) _
  As Long
  
Private Declare Sub MoveMemory _
  Lib "kernel32" _
  Alias "RtlMoveMemory" _
  (dest As Any, _
    ByVal Source As Long, _
    ByVal length As Long)
    
    
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
'Private Const MAX_PATH = 260
Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer
    dwStrucVersionh As Integer
    dwFileVersionMSl As Integer
    dwFileVersionMSh As Integer
    dwFileVersionLSl As Integer
    dwFileVersionLSh As Integer
    dwProductVersionMSl As Integer
    dwProductVersionMSh As Integer
    dwProductVersionLSl As Integer
    dwProductVersionLSh As Integer
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type


Public Function GetDriveList()
On Error Resume Next
    Dim fso          As FileSystemObject
    Dim d            As Drive
    Dim dc           As Drives
    Dim s            As String
    Dim sName        As String
    Dim sType        As String
    Dim dAvailSpace  As Double
    Dim dFreeSpace   As Double
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set dc = fso.Drives
    For Each d In dc
         sName = ""
         s = s & d.DriveLetter & " - "
         sType = GetDriveTypeName(d.DriveLetter)
         s = s & sType & " - "
    
         If d.DriveType = gi_DriveTypeNetwork Then
            sName = d.ShareName
         ElseIf d.IsReady Then
            sName = d.VolumeName
         End If
         
       dAvailSpace = Format(((d.TotalSize / 1024) / 1024), "###,###.00")
       dFreeSpace = Format(((d.FreeSpace / 1024) / 1024), "###,###.00")
       s = s & Format(dAvailSpace, "###,###.00") & " Mb - "
       s = s & Format(dFreeSpace, "###,###.00") & " Mb - "
       s = s & d.SerialNumber & " - "
       s = s & d.FileSystem & " - "
       s = s & sName & vbCrLf
    Next
    GetDriveList = s
    Set fso = Nothing
    Set dc = Nothing
    
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetDriveList")
End Function

Public Sub GetDriveParam(sDriveLetter As String, _
                        sVolName As String, _
                        sFileSystem As String, _
                        sSerialNumber As String, _
                        iType As Long, _
                        sType As String, _
                        dTotalSpace As Double, _
                        dFreeSpace As Double)
   On Error Resume Next
   Dim fso As FileSystemObject
   Dim d As Drive
   
   Set fso = CreateObject("Scripting.FileSystemObject")
   
   Set d = fso.GetDrive(sDriveLetter)
   iType = d.DriveType
    Select Case iType
       Case gi_DriveTypeUnknown:    sType = "Unknown"
       Case gi_DriveTypeRemovable:  sType = "Removable"
       Case gi_DriveTypeFixed:      sType = "Fixed"
       Case gi_DriveTypeNetwork:    sType = "Network"
       Case gi_DriveTypeCDROM:      sType = "CD-ROM"
       Case gi_DriveTypeRAMDisk:    sType = "RAM Disk"
    End Select
   
    If d.DriveType = 3 Then
       sVolName = d.ShareName
    ElseIf d.IsReady Then
       sVolName = d.VolumeName
    End If
    dTotalSpace = d.TotalSize
    dFreeSpace = d.FreeSpace
    sSerialNumber = d.SerialNumber & " - "
    sFileSystem = d.FileSystem & " - "
    Set fso = Nothing
    Set d = Nothing

ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetDriveParam")
End Sub

Public Function GetDriveSerialNum(sDriveLetter As String) As String
On Error GoTo ErrHandler
    Dim fso  As FileSystemObject
    Dim d    As Drive
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set d = fso.GetDrive(sDriveLetter)
    GetDriveSerialNum = d.SerialNumber
    Set fso = Nothing
    Set d = Nothing

ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetDriveSerialNum")
End Function

Public Function GetDriveFreeSpace(sDriveLetter As String) As Double
On Error GoTo ErrHandler
    Dim fso  As FileSystemObject
    Dim d    As Drive
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set d = fso.GetDrive(sDriveLetter)
    GetDriveFreeSpace = d.FreeSpace
    Set fso = Nothing
    Set d = Nothing

ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetDriveFreeSpace")
End Function

Public Function GetDriveTotalSpace(sDriveLetter As String) As Double
On Error GoTo ErrHandler
    Dim fso  As FileSystemObject
    Dim d    As Drive
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set d = fso.GetDrive(sDriveLetter)
    GetDriveTotalSpace = d.TotalSize
    Set fso = Nothing
    Set d = Nothing

ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetDriveTotalSpace")
End Function

Public Function GetDriveType(sDriveLetter As String) As Long
On Error GoTo ErrHandler
    Dim fso  As FileSystemObject
    Dim d    As Drive
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set d = fso.GetDrive(sDriveLetter)
    GetDriveType = d.DriveType
    Set fso = Nothing
    Set d = Nothing

ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetDriveType")
End Function

Public Function GetDriveTypeName(sDriveLetter As String) As String
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim d   As Drive
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set d = fso.GetDrive(sDriveLetter)
   
    Select Case d.DriveType
       Case gi_DriveTypeUnknown:    GetDriveTypeName = "Unknown"
       Case gi_DriveTypeRemovable:  GetDriveTypeName = "Removable"
       Case gi_DriveTypeFixed:      GetDriveTypeName = "Fixed"
       Case gi_DriveTypeNetwork:    GetDriveTypeName = "Network"
       Case gi_DriveTypeCDROM:      GetDriveTypeName = "CD-ROM"
       Case gi_DriveTypeRAMDisk:    GetDriveTypeName = "RAM Disk"
    End Select
    Set fso = Nothing
    Set d = Nothing
    
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetDriveTypeName")
End Function

Public Function GetFileSize(sFilePath As String) As Double
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim f   As File
    If Right(sFilePath, 1) = "\" Then
        GetFileSize = 0
        Exit Function
    End If
    If Dir(sFilePath) = "" Then
        GetFileSize = 0
        Exit Function
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(sFilePath)
    GetFileSize = f.Size
    Set fso = Nothing
    Set f = Nothing

ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetFileSize")
End Function

Public Function GetFileDateCreate(sFilePath As String) As Date
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim f   As File
    If Right(sFilePath, 1) = "\" Then
        GetFileDateCreate = "1/1/1900"
        Exit Function
    End If
    If Dir(sFilePath) = "" Then
        GetFileDateCreate = "1/1/1900"
        Exit Function
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(sFilePath)
    GetFileDateCreate = f.DateCreated
    Set fso = Nothing
    Set f = Nothing
    
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetFileDateCreate")
End Function


Public Function GetFileDateModified(sFilePath As String) As Date
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim f   As File
    If Right(sFilePath, 1) = "\" Then
        GetFileDateModified = "1/1/1900"
        Exit Function
    End If
    If Dir(sFilePath) = "" Then
        GetFileDateModified = "1/1/1900"
        Exit Function
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(sFilePath)
    GetFileDateModified = f.DateLastModified
    Set fso = Nothing
    Set f = Nothing

ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetFileDateModified")
End Function


Public Function GetFileDateAccessed(sFilePath As String) As Date
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim f   As File
    If Right(sFilePath, 1) = "\" Then
        GetFileDateAccessed = "1/1/1900"
        Exit Function
    End If
    If Dir(sFilePath) = "" Then
        GetFileDateAccessed = "1/1/1900"
        Exit Function
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(sFilePath)
    GetFileDateAccessed = f.DateLastAccessed
    Set fso = Nothing
    Set f = Nothing
   
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetFileDateAccessed")
End Function


Public Function GetFolderSize(sFolderPath As String) As Double
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim f   As Folder
    If Dir(sFolderPath, vbDirectory) = "" Then
        GetFolderSize = 0
        Exit Function
    End If
    
    If Right(sFolderPath, 1) = "\" Then
        GetFolderSize = 0
        Exit Function
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFolder(sFolderPath)
    GetFolderSize = f.Size
    Set fso = Nothing
    Set f = Nothing
   
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetFolderSize")
End Function

Public Function GetFolderDateCreate(sFolderPath As String) As Date
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim f   As Folder
    If Dir(sFolderPath, vbDirectory) = "" Then
        GetFolderDateCreate = "1/1/1900"
        Exit Function
    End If
    
    If Len(sFolderPath) <= 3 Then
        GetFolderDateCreate = "1/1/1900"
        Exit Function
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFolder(sFolderPath)
    GetFolderDateCreate = f.DateCreated
    Set fso = Nothing
    Set f = Nothing
   
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetFolderDateCreate")
End Function

Public Function GetFolderDateAccessed(sFolderPath As String) As Date
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim f   As Folder
    If Dir(sFolderPath, vbDirectory) = "" Then
        GetFolderDateAccessed = "1/1/1900"
        Exit Function
    End If
    If Len(sFolderPath) <= 3 Then
        GetFolderDateAccessed = "1/1/1900"
        Exit Function
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFolder(sFolderPath)
    GetFolderDateAccessed = f.DateLastAccessed
    Set fso = Nothing
    Set f = Nothing
   
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetFolderDateAccessed")
End Function

Public Function GetFolderDateModified(sFolderPath As String) As Date
On Error GoTo ErrHandler
    Dim fso  As FileSystemObject
    Dim f    As Folder
    If Dir(sFolderPath, vbDirectory) = "" Then
        GetFolderDateModified = "1/1/1900"
        Exit Function
    End If
    If Len(sFolderPath) <= 3 Then
        GetFolderDateModified = "1/1/1900"
        Exit Function
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFolder(sFolderPath)
    GetFolderDateModified = f.DateLastModified
    Set fso = Nothing
    Set f = Nothing
   
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetFolderDateModified")
End Function


Public Function GetFileList(sFolderPath As String, Optional sDelimiter As String = vbCrLf) As String
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim f   As Folder
    Dim f1  As File
    Dim fc  As Files
    Dim s   As String
    
    If Dir(sFolderPath, vbDirectory) <> "" Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.GetFolder(sFolderPath)
        Set fc = f.Files
        For Each f1 In fc
           s = s & f1.Name & sDelimiter
        Next
        GetFileList = s
        Set fso = Nothing
        Set f = Nothing
    Else
        GetFileList = ""
    End If
    
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetFileList")
End Function

Public Function GetFileShortName(sFilePath As String) As String
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim f   As File
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(sFilePath)
    GetFileShortName = f.ShortName
    Set fso = Nothing
    Set f = Nothing

ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetFileShortName")
End Function

Public Function GetFileName(sFilePath As String) As String
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim f   As File
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(sFilePath)
    GetFileName = f.Name
    Set fso = Nothing
    Set f = Nothing

ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetFileName")
End Function

Public Function GetFilePath(sFilePath As String) As String
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim f   As File
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(sFilePath)
    GetFilePath = f.Path
    Set fso = Nothing
    Set f = Nothing

ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetFilePath")
End Function

Public Function GetFileVersion(sFilePath As String) As String
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim sFileExt    As String
    
    WriteToFile gs_LAST_FILE_LOG, sFilePath
    
    sFileExt = GetFileExt(sFilePath)
    If sFileExt = "exe" Or _
        sFileExt = "dll" Or _
        sFileExt = "ocx" Then
    
        If Dir(sFilePath) = "" Then
            GetFileVersion = ""
            Exit Function
        End If
    
        If Right(sFilePath, 1) = "\" Then
            GetFileVersion = ""
            Exit Function
        End If
    
        Set fso = CreateObject("Scripting.FileSystemObject")
        GetFileVersion = fso.GetFileVersion(sFilePath)
    
        Set fso = Nothing
    Else
        GetFileVersion = ""
    End If
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetFileVersion")
End Function



Public Function GetFolderList(sFolderPath As String, Optional sDelimiter As String = vbCrLf) As String
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim f   As Folder
    Dim f1  As Folder
    Dim s   As String
    Dim sf
    
    If Dir(sFolderPath, vbDirectory) <> "" Then
    
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.GetFolder(sFolderPath)
        Set sf = f.SubFolders
        
        For Each f1 In sf
           s = s & f1.Name & sDelimiter
        Next
        
        GetFolderList = s
        Set fso = Nothing
        Set f = Nothing
        Set sf = Nothing
    Else
        GetFolderList = ""
    End If
    
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetFolderList")
End Function
Public Function GetDriveAvailableSpace(drvPath As String) As Long
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim d As Drive
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set d = fso.GetDrive(fso.GetDriveName(drvPath))
    GetDriveAvailableSpace = d.AvailableSpace
    Set fso = Nothing
    Set d = Nothing

ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetDriveAvailableSpace")
End Function

Public Function GetDriveName(drvPath As String) As String
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim d As Drive
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set d = fso.GetDrive(fso.GetDriveName(drvPath))
    GetDriveName = d.VolumeName
    Set fso = Nothing
    Set d = Nothing

ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetDriveName")
End Function

Public Function GetParentFolderFSO(sFilePath As String) As String
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim f As File
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(sFilePath)
    GetParentFolderFSO = f.ParentFolder
    Set fso = Nothing
    Set f = Nothing

ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetParentFolderFSO")
End Function

Public Function GetFolderLevelDepthFSO(sFilePath As String) As Integer
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim f As Folder
    Dim n As Integer
    
    If Dir(sFilePath, vbDirectory) = "" Then
        GetFolderLevelDepthFSO = 0
        Exit Function
    End If
    
    If Right(sFilePath, 1) = "\" Then
        GetFolderLevelDepthFSO = 0
        Exit Function
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFolder(sFilePath)
    n = 0
    If f.IsRootFolder Then
       GetFolderLevelDepthFSO = n
    Else
       Do Until f.IsRootFolder
          Set f = f.ParentFolder
          n = n + 1
       Loop
       GetFolderLevelDepthFSO = n
    End If
    Set fso = Nothing
    Set f = Nothing
   
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetFolderLevelDepthFSO")
End Function

Public Sub CreateFolderFSO(sFilePath As String)
On Error GoTo ErrHandler
    Dim fso As FileSystemObject, f As Folder
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder (sFilePath)
    Set fso = Nothing
    
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler("modFSO", 0, "CreateFolderFSO")
End Sub

Public Sub DeleteFolderFSO(sFilePath As String)
On Error GoTo ErrHandler
    Dim fso As FileSystemObject, f As Folder
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFolder (sFilePath)
    Set fso = Nothing
    
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler("modFSO", 0, "DeleteFolderFSO")
End Sub

Public Sub MoveFolderFSO(sFolderFromPath As String, sFolderToPath As String)
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.MoveFolder sFolderFromPath, sFolderToPath
    Set fso = Nothing
    
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler("modFSO", 0, "MoveFolderFSO")
End Sub

Public Sub CopyFolderFSO(sFolderFromPath As String, sFolderToPath As String, Optional bOverWrite As Boolean = True)
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFolder sFolderFromPath, sFolderToPath, bOverWrite
    Set fso = Nothing
    
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler("modFSO", 0, "CopyFolderFSO")
End Sub

Public Sub MoveFileFSO(sFileFromPath As String, sFileToPath As String)
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.MoveFile sFileFromPath, sFileToPath
    Set fso = Nothing
    
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler("modFSO", 0, "MoveFileFSO")
End Sub

Public Sub CopyFileFSO(sFileFromPath As String, sFileToPath As String, Optional bOverWrite As Boolean = True)
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile sFileFromPath, sFileToPath, bOverWrite
    Set fso = Nothing

ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler("modFSO", 0, "CopyFileFSO")
End Sub


Public Sub CreateTextFile(sFilePath As String)
On Error GoTo ErrHandler
    Dim fso As FileSystemObject
    Dim f1 As File
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f1 = fso.CreateTextFile(sFilePath, True)
    Set fso = Nothing
    Set f1 = Nothing
    
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler("modFSO", 0, "CreateTextFile")
End Sub

Public Sub WriteToFile(sFilePath As String, sStr As String)
On Error GoTo ErrHandler
    Dim fso As FileSystemObject, tf
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set tf = fso.CreateTextFile(sFilePath, ForWriting)
    tf.Write (sStr)
    tf.Close
    Set fso = Nothing
    Set tf = Nothing
    
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler("modFSO", 0, "WriteToFile")
End Sub

Public Sub AppendingToFile(sFilePath As String, sStr As String)
On Error GoTo ErrHandler
   Dim fso As FileSystemObject, tf
   
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set tf = fso.OpenTextFile(sFilePath, ForAppending)
    tf.Write (sStr)
    tf.Close
    Set fso = Nothing
    Set tf = Nothing
   
ErrExit:         Exit Sub
ErrHandler:      Call ErrorHandler("modFSO", 0, "AppendingToFile")
End Sub

Public Function ReadingFromFile(sFilePath As String) As String
On Error GoTo ErrHandler
    Dim fso As FileSystemObject, tf
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set tf = fso.OpenTextFile(sFilePath, ForReading)
    ReadingFromFile = tf.Read
    tf.Close
    Set fso = Nothing
    Set tf = Nothing
   
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "ReadingFromFile")
End Function

Public Function GetFileExt(sFilePath As String) As String
On Error GoTo ErrHandler
    
    If InStr(1, sFilePath, ".") > 0 Then
        GetFileExt = Right(Trim(sFilePath), 3)
    Else
        GetFileExt = " "
    End If
   
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetFileExt")
End Function

Public Function CheckPath(sSourcePath As String) As String
On Error GoTo ErrHandler
    sSourcePath = Replace(sSourcePath, "/", "\")
    CheckPath = Replace(sSourcePath, "\\", "\")
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "CheckPath")
End Function

Public Function GetLastFolderFromPath(sSourcePath As String) As String
On Error GoTo ErrHandler
    Dim i   As Integer
    Dim s   As String
    i = InStrRev(sSourcePath, "\")
    If i > 0 Then
        GetLastFolderFromPath = Mid(sSourcePath, i, Len(sSourcePath) - i + 1)
    Else
        GetLastFolderFromPath = ""
    End If
    
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetLastFolderFromPath")
End Function

Public Function GetLastFileFromPath(sSourcePath As String) As String
On Error GoTo ErrHandler
    Dim i   As Integer
    
    i = InStrRev(sSourcePath, "\")
    If i > 0 Then
        GetLastFileFromPath = Mid(sSourcePath, i, Len(sSourcePath) - i + 1)
    Else
        GetLastFileFromPath = ""
    End If

ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "GetLastFileFromPath")
End Function

Public Function AppendToTextFile(sFilePath As String, sText As String) As Boolean
On Error GoTo ErrHandler
  Dim iFile As Integer

  ' Get a free file handle
  iFile = FreeFile

  ' Open the file in append mode, write to it, and close it
  Open sFilePath For Append Shared As iFile
  Print #iFile, sText
  Close #iFile

  AppendToTextFile = True

ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("modFSO", 0, "AppendToTextFile")
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
Private Sub ErrorHandler(sFormName As String, ind As Integer, SubName As String, Optional sSql As String, Optional bLogOnly As Boolean)
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

