Attribute VB_Name = "ModFileIO"
Option Explicit

Const MOVEFILE_REPLACE_EXISTING = &H1
Const FILE_ATTRIBUTE_TEMPORARY = &H100
Const FILE_BEGIN = 0
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const CREATE_NEW = 1
Const OPEN_EXISTING = 3
Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const OF_EXIST As Long = &H4000
Private Const OFS_MAXPATHNAME As Long = 128
Private Const HFILE_ERROR As Long = -1

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

Public Function FileExists(FilePath As String) As Boolean
    Dim lRetVal As Long
    Dim OfSt As OFSTRUCT
    
    lRetVal = OpenFile(FilePath, OfSt, OF_EXIST)
    
    If lRetVal <> HFILE_ERROR Then
        FileExists = True
    Else
        FileExists = False
    End If
    
    CloseHandle lRetVal
End Function

Public Function FormatSize(ByVal Size As Currency) As String
    Const Kilobyte As Currency = 1024@
    Const HundredK As Currency = 102400@
    Const ThousandK As Currency = 1024000@
    Const Megabyte As Currency = 1048576@
    Const HundredMeg As Currency = 104857600@
    Const ThousandMeg As Currency = 1048576000@
    Const Gigabyte As Currency = 1073741824@
    Const Terabyte As Currency = 1099511627776@
    
    If Size < Kilobyte Then
        FormatSize = Int(Size) & " bytes"
    ElseIf Size < HundredK Then
        FormatSize = Format(Size / Kilobyte, "#.0") & " KB"
    ElseIf Size < ThousandK Then
        FormatSize = Int(Size / Kilobyte) & " KB"
    ElseIf Size < HundredMeg Then
        FormatSize = Format(Size / Megabyte, "#.0") & " MB"
    ElseIf Size < ThousandMeg Then
        FormatSize = Int(Size / Megabyte) & " MB"
    ElseIf Size < Terabyte Then
        FormatSize = Format(Size / Gigabyte, "#.00") & " GB"
    Else
        FormatSize = Format(Size / Terabyte, "#.00") & " TB"
    End If
End Function

Public Sub API_OpenFile(ByVal FileName As String, ByRef FileNumber As Long, ByRef FileSize As Currency)
Dim FileH As Long
Dim Ret As Long
On Error Resume Next
FileH = CreateFile(FileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0, 0)
If Err.Number > 0 Then
    Err.Clear
    FileNumber = -1
Else
    FileNumber = FileH
    Ret = SetFilePointer(FileH, 0, 0, FILE_BEGIN)
    API_FileSize FileH, FileSize
End If
On Error GoTo 0
End Sub
Public Sub API_FileSize(ByVal FileNumber As Long, ByRef FileSize As Currency)
    Dim FileSizeL As Long
    Dim FileSizeH As Long
    FileSizeH = 0
    FileSizeL = GetFileSize(FileNumber, FileSizeH)
    Long2Size FileSizeL, FileSizeH, FileSize
End Sub

Public Sub API_ReadFile(ByVal FileNumber As Long, ByVal Position As Currency, ByRef BlockSize As Long, ByRef Data() As Byte)
Dim PosL As Long
Dim PosH As Long
Dim SizeRead As Long
Dim Ret As Long
Size2Long Position, PosL, PosH
Ret = SetFilePointer(FileNumber, PosL, PosH, FILE_BEGIN)
Ret = ReadFile(FileNumber, Data(0), BlockSize, SizeRead, 0&)
BlockSize = SizeRead
End Sub

Public Sub API_CloseFile(ByVal FileNumber As Long)
Dim Ret As Long
Ret = CloseHandle(FileNumber)
End Sub

Public Sub API_WriteFile(ByVal FileNumber As Long, ByVal Position As Currency, ByRef BlockSize As Long, ByRef Data() As Byte)
Dim PosL As Long
Dim PosH As Long
Dim SizeWrit As Long
Dim Ret As Long
Size2Long Position, PosL, PosH
Ret = SetFilePointer(FileNumber, PosL, PosH, FILE_BEGIN)
Ret = WriteFile(FileNumber, Data(0), BlockSize, SizeWrit, 0&)
BlockSize = SizeWrit
End Sub

Private Sub Size2Long(ByVal FileSize As Currency, ByRef LongLow As Long, ByRef LongHigh As Long)
'&HFFFFFFFF unsigned = 4294967295
Dim Cutoff As Currency
Cutoff = 2147483647
Cutoff = Cutoff + 2147483647
Cutoff = Cutoff + 1 ' now we hold the value of 4294967295 and not -1
LongHigh = 0
Do Until FileSize < Cutoff
    LongHigh = LongHigh + 1
    FileSize = FileSize - Cutoff
Loop
If FileSize > 2147483647 Then
    LongLow = -CLng(Cutoff - (FileSize - 1))
Else
    LongLow = CLng(FileSize)
End If
End Sub

Private Sub Long2Size(ByVal LongLow As Long, ByVal LongHigh As Long, ByRef FileSize As Currency)
'&HFFFFFFFF unsigned = 4294967295
Dim Cutoff As Currency
Cutoff = 2147483647
Cutoff = Cutoff + 2147483647
Cutoff = Cutoff + 1 ' now we hold the value of 4294967295 and not -1
FileSize = Cutoff * LongHigh
If LongLow < 0 Then
    FileSize = FileSize + (Cutoff + (LongLow + 1))
Else
    FileSize = FileSize + LongLow
End If
End Sub


