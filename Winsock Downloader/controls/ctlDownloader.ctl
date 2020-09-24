VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl ctlDownloader 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ctlDownloader.ctx":0000
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   ToolboxBitmap   =   "ctlDownloader.ctx":0C44
   Windowless      =   -1  'True
   Begin MSWinsockLib.Winsock sckFTP 
      Left            =   360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrSpeed 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   720
      Top             =   0
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "ctlDownloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ctlDownloader.ctl
'=================
'Version:       1.0.0b
'Author:        Daniel Elkins (DigiRev)
'Copyright:     (c) 2008 Daniel Elkins
'Website:       http://www.DigiRev.org
'E-mail:        Daniel@DigiRev.org
'Created:       March 19th, 2008
'Last-updated:  March 19th, 2008

'Description: Download files in your VB6 programs via HTTP or FTP with this control.
'             Support for files greater than 4GB.

'License:     You are free to use this control in your software, both free and
'             commercial. You are allowed to modify the code as you need to. You
'             are NOT allowed to redistribute the code (compiled or otherwise) without
'             express permission from the author! You may NOT sell this code (compiled or otherwise).

'Credits:     Win32 file I/O functions provided by CodeGuru. A couple of other
'             procedures were taken from other sources, and credit to the author
'             is displayed above them.
'
'             File I/O: http://www.codeguru.com/vb/controls/vb_file/directory/print.php/c12917__2/

Option Explicit

Public Event DownloadError(ByVal Number As Long, Description As String)
Public Event DownloadStarted(ByVal FileSize As Currency)
Public Event DownloadProgress(ByVal BytesReceived As Currency, ByVal FileSize As Currency)
Public Event DownloadRedirect(ByVal Location As String)
Public Event DownloadSpeed(ByVal BytesPerSecond As Single)
Public Event DownloadComplete(ByVal BytesReceived As Currency, ByVal FileSize As Currency)
Public Event SocketClose()
Public Event SocketConnect()
Public Event SocketError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

Private Const FILE_BEGIN = 0
Private Const FILE_END = 2
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
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

Public Enum DL_PROTOCOL
    proHTTP = 0
    proFTP = 1
End Enum

Private Enum HTTP_STATE
    intIdle = 0
    intConnecting = 1
    intHEAD = 2
    intGET = 3
    intReceivingFile = 4
    intFileReceived = 5
End Enum

Private Enum FTP_COMMAND
    ftpNONE = 0
    FTPUser = 1
    FTPPass = 2
    ftpCWD = 3
    ftpPWD = 4
    ftpPASV = 5
    ftpRETR = 6
    ftpSIZE = 7
End Enum

Private Declare Function GetInputState Lib "user32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

'Default property constants.
'---------------------------
Private Const DEF_USER_AGENT As String = "WinsockDownloader (DigiRev)"
Private Const DEF_MAX_REDIRECTS As Integer = 5

'Property container variables.
'-----------------------------
Private p_MaxRedirects As Integer   'Max redirects before erroring out.
Private p_UserAgent As String       'User agent to use in HTTP requests.

'Other variables.
'----------------
Private strHost As String           'Remote host of HTTP/FTP server (IP or DNS).
Private lonPort As Long             'Remote port of server.
Private curFileSize As Currency     'Size of file in bytes.
Private lonFH As Long               'File handle.
Private curBytesRec As Currency     'Number of bytes received of file.
Private strSavePath As String       'Path to save file.
Private strURI As String            'Request URI (file path being requested).
Private intRedirects As Integer     'Number of redirects encountered.
Private intHTTPState As HTTP_STATE  'State of HTTP session.
Private intProtocol As DL_PROTOCOL  'Protocol we're using (HTTP/FTP).
Private bytBuffer() As Byte         'TCP data buffer for HTTP session.
Private curDownloadSec As Currency  'Currently downloaded data in a second (for transfer speed).
Private sinDownloadSpeed As Single  'Current download speed.
Private intFTPState As FTP_COMMAND  'Last FTP command sent.
Private bolFTPSending As Boolean    'Sending data to the FTP server? (so packets don't get sent too fast).
Private strUsername As String       'FTP username.
Private strPassword As String       'FTP password.

'Extract info from HTTP header into our variables.
Private Sub ExtractFromHeader(ByRef Data() As Byte)
    Dim strSize As String
    
    If curFileSize <= 0 Then
        strSize = Trim$(HTTPField(Data, "Content-Length:"))
        If IsNumeric(strSize) Then
            If CCur(strSize) <> 0 Then
                curFileSize = CCur(Trim$(HTTPField(Data, "Content-Length:")))
            End If
        End If
    End If
End Sub

'HTTP User-Agent.
Public Property Get UserAgent() As String
    UserAgent = p_UserAgent
End Property

Public Property Let UserAgent(ByRef NewValue As String)
    If Len(NewValue) > 0 Then
        p_UserAgent = NewValue
        PropertyChanged "UserAgent"
    End If
End Property

'Max HTTP redirects before erroring-out.
Public Property Get MaxRedirects() As Integer
    MaxRedirects = p_MaxRedirects
End Property

Public Property Let MaxRedirects(ByVal NewValue As Integer)
    If NewValue > 0 And NewValue < 200 Then
        p_MaxRedirects = NewValue
        PropertyChanged "MaxRedirects"
    Else
        MsgBox "MaxRedirects must be between 1 and 200!", vbExclamation
    End If
End Property

'Reset session information.
Private Sub ResetAll()
    curFileSize = -1
    If lonFH >= 0 Then API_CloseFile lonFH
    lonFH = -1
    curBytesRec = 0
    intHTTPState = intIdle
    intFTPState = ftpNONE
    Erase bytBuffer
    curDownloadSec = 0
    sinDownloadSpeed = 0
End Sub

'Initiate a download via FTP.
Public Sub DownloadFTP(ByRef Hostname As String, ByRef Username As String, _
    ByRef Password As String, ByRef SavePath As String, _
    ByRef FileLocation As String, Optional ByVal Port As Long = 21)
    
    ResetAll
    intProtocol = proFTP
    strSavePath = SavePath
    strHost = Hostname
    strUsername = Username
    strPassword = Password
    strURI = FileLocation
    With sckMain
        .Close
        .RemoteHost = Hostname
        .RemotePort = Port
        .Connect
    End With
End Sub

'Initiate a download via HTTP.
Public Sub DownloadHTTP(ByRef URL As String, ByRef SavePath As String)
    On Error GoTo ErrorHandler
    tmrSpeed.Enabled = False
    intProtocol = proHTTP
    strSavePath = SavePath
    ParseURLHTTP URL, strHost, lonPort, strURI
    ResetAll
    intRedirects = 0
    With sckMain
        .Close
        .RemoteHost = strHost
        .RemotePort = lonPort
        .Connect
        intHTTPState = intConnecting
    End With
    Exit Sub
    
ErrorHandler:
    RaiseEvent DownloadError(Err.Number, Err.Description)
    Exit Sub
End Sub

'Extract information from a URL.
Private Sub ParseURLHTTP(URL As String, _
    Optional ByRef Hostname As String, _
    Optional ByRef Port As Long, _
    Optional ByRef RequestURI As String)
    
    Dim s As String, i As Integer
    Dim strHost As String, strPort As String
    s = Replace(URL, "http://", "", , , vbTextCompare)
    i = InStr(1, s, "/")
    If i > 0 Then strHost = Left$(s, i - 1) Else strHost = s
    i = InStr(1, strHost, ":")
    If i > 0 Then
        strPort = Mid$(strHost, i + 1)
        strHost = Left$(strHost, i - 1)
    Else
        strPort = "80"
    End If
    i = InStr(1, s, "/")
    If i > 0 Then RequestURI = Mid$(s, i) Else RequestURI = "/"
    Hostname = strHost
    Port = CLng(strPort)
End Sub

'Check if a file exists.
Private Function FileExists(FilePath As String) As Boolean
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

'Following API_ functions provided by CodeGuru:
'http://www.codeguru.com/vb/controls/vb_file/directory/print.php/c12917__2/
Private Sub API_OpenFile(ByVal FileName As String, ByRef FileNumber As Long, ByRef FileSize As Currency)
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

Private Sub API_FileSize(ByVal FileNumber As Long, ByRef FileSize As Currency)
    Dim FileSizeL As Long
    Dim FileSizeH As Long
    FileSizeH = 0
    FileSizeL = GetFileSize(FileNumber, FileSizeH)
    Long2Size FileSizeL, FileSizeH, FileSize
End Sub

Private Sub API_CloseFile(ByVal FileNumber As Long)
    Dim Ret As Long
    Ret = CloseHandle(FileNumber)
End Sub

Private Sub API_WriteFile(ByVal FileNumber As Long, ByVal Position As Currency, ByRef BlockSize As Long, ByRef Data() As Byte)
    Dim PosL As Long
    Dim PosH As Long
    Dim SizeWrit As Long
    Dim Ret As Long
    Size2Long Position, PosL, PosH
    Ret = SetFilePointer(FileNumber, PosL, PosH, FILE_END)
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

'FTP data socket is connected.
Private Sub sckFTP_Connect()
    If FileExists(strSavePath) Then
        SafeKill strSavePath
    End If
    Dim i As Integer: i = FreeFile
    Dim curTemp As Currency
    Open strSavePath For Binary As #i
    Close #i
    API_OpenFile strSavePath, lonFH, curTemp
    tmrSpeed.Enabled = True
    WaitSend
    sckMain.SendData "TYPE I" & vbCrLf 'Set transfer mode to binary.
    WaitSend
    If Left$(strURI, 1) <> "/" Then strURI = "/" & strURI 'Prefix path with / if not present.
    sckMain.SendData "RETR " & strURI & vbCrLf 'Send RETR command to download the file.
End Sub

'Error opening data connection.
Private Sub sckFTP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckFTP.Close
    If lonFH > -1 Then API_CloseFile lonFH
    tmrSpeed.Enabled = False
    RaiseEvent DownloadError(Number, "Error opening data connection!")
End Sub

'Connection closed remotely.
Private Sub sckMain_Close()
    If intProtocol = proFTP Then bolFTPSending = False
    RaiseEvent SocketClose
    tmrSpeed.Enabled = False
End Sub

'Main socket is connected.
Private Sub sckMain_Connect()
    Dim strReq As String
    If intProtocol = proHTTP Then 'We're downloading a file via HTTP.
        If intHTTPState = intConnecting Then
            RaiseEvent SocketConnect
            'Send HEAD request.
            strReq = "HEAD " & strURI & " HTTP/1.1" & vbCrLf
            strReq = strReq & "Accept: *.*" & vbCrLf
            strReq = strReq & "Accept-Encoding: gzip,deflate" & vbCrLf
            strReq = strReq & "Host: " & strHost & vbCrLf
            strReq = strReq & "User-Agent: " & p_UserAgent & vbCrLf
            strReq = strReq & "Connection: close" & vbCrLf & vbCrLf
            sckMain.SendData strReq
            intHTTPState = intHEAD
        ElseIf intHTTPState = intHEAD Then
            'Send GET request.
            strReq = "GET " & strURI & " HTTP/1.1" & vbCrLf
            strReq = strReq & "Accept: *.*" & vbCrLf
            strReq = strReq & "Accept-Encoding: gzip,deflate" & vbCrLf
            strReq = strReq & "Host: " & strHost & vbCrLf
            strReq = strReq & "User-Agent: " & p_UserAgent & vbCrLf
            strReq = strReq & "Connection: close" & vbCrLf & vbCrLf
            intHTTPState = intGET
            sckMain.SendData strReq
        End If
    ElseIf intProtocol = proFTP Then 'We're downloading a file via FTP.
        bolFTPSending = False
        RaiseEvent SocketConnect
        'Wait for 220 Welcome message.
    End If
End Sub

'Loop until data is sent.
Private Sub WaitSend()
    Do While bolFTPSending
        If GetInputState Then DoEvents
    Loop
End Sub

'Receiving the file's contents.
'Dump it to disk.
Private Sub sckFTP_DataArrival(ByVal bytesTotal As Long)
    Dim bytData() As Byte, lonUB As Long
    sckFTP.GetData bytData, vbByte + vbArray, bytesTotal

    lonUB = UBound(bytData)
    API_WriteFile lonFH, 0, lonUB + 1, bytData
    curBytesRec = curBytesRec + lonUB + 1
    RaiseEvent DownloadProgress(curBytesRec, curFileSize)
    If curBytesRec = curFileSize Then
        API_CloseFile lonFH
        RaiseEvent DownloadComplete(curBytesRec, curFileSize)
        tmrSpeed.Enabled = False
        RaiseEvent DownloadSpeed(0)
    End If
End Sub

'We're receiving data from the server (HTTP or FTP).
Private Sub sckMain_DataArrival(ByVal bytesTotal As Long)
    'Get received data.
    On Error GoTo ErrorHandler
    
    Dim bytData() As Byte, lonUB As Long
    Dim lonStatusCode As Long, lonDataStart As Long
    Dim curTemp As Currency
    Dim bytFileData() As Byte, lonDataLen As Long
    Dim strLocation As String, strData As String
    Dim intCMD As Integer, strPHost As String, lonPPort As Long
    Dim strSize As String
    
    'If we're downloading an HTTP file...
    If intProtocol = proHTTP Then
        sckMain.GetData bytData, vbByte + vbArray, bytesTotal
        'Append data to buffer.
        lonUB = Byte_UBound(bytBuffer)
        If lonUB < 0 Then
            ReDim bytBuffer(0 To UBound(bytData)) As Byte
            CopyMemory bytBuffer(0), bytData(0), UBound(bytData) + 1
        Else
            ReDim Preserve bytBuffer(0 To UBound(bytBuffer) + UBound(bytData) + 1) As Byte
            CopyMemory bytBuffer(lonUB + 1), bytData(0), UBound(bytData) + 1
        End If
        Erase bytData
        
        If intHTTPState = intHEAD Then 'Sent HEAD request. Get file info.
        
            'Check for end of header info (vbCrLf & vbCrLf).
            If BytInStr(0, bytBuffer, vbCrLf & vbCrLf) >= 0 Then
            
                'Found end of header.
                'Check HTTP status code (404, 200, etc).
                lonStatusCode = HTTPStatusCode(bytBuffer)
                Select Case lonStatusCode
                
                    Case 200 'HTTP OK
                        'Extract file size, etc. from header.
                        ExtractFromHeader bytBuffer
                        
                        'Erase buffer.
                        Erase bytBuffer
                        
                        'Send GET request.
                        intHTTPState = intHEAD
                        sckMain.Close
                        sckMain.Connect
                        'GET request sent. HTTPStatus = intGET
                    
                    Case 300 To 303 'Some sort of redirect
                        strLocation = LTrim$(HTTPField(bytBuffer, "Location:"))
                        If Len(strLocation) > 0 Then
                            intRedirects = intRedirects + 1
                            If MaxRedirects > 0 And intRedirects > MaxRedirects Then
                                RaiseEvent DownloadError(-1, "Redirected too many times.")
                                sckMain.Close
                                intHTTPState = intIdle
                                Erase bytBuffer
                                Exit Sub
                            Else
                                ResetAll
                                intHTTPState = intConnecting
                                ParseURLHTTP strLocation, strHost, lonPort, strURI
                                sckMain.Close
                                sckMain.Connect
                                RaiseEvent DownloadRedirect(strLocation)
                            End If
                        Else
                            RaiseEvent DownloadError(-1, "Redirected but no location specified.")
                            sckMain.Close
                            intHTTPState = intIdle
                            Erase bytBuffer
                            Exit Sub
                        End If
                    
                    Case Else 'Some other response.
                        sckMain.Close
                        intHTTPState = intIdle
                        If lonStatusCode = 400 Then 'Bad request.
                            RaiseEvent DownloadError(400, "Bad request.")
                        ElseIf lonStatusCode = 401 Then 'Unauthorized.
                            RaiseEvent DownloadError(401, "Unauthorized.")
                        ElseIf lonStatusCode = 402 Then 'Payment required.
                            RaiseEvent DownloadError(402, "Payment required.")
                        ElseIf lonStatusCode = 403 Then 'Forbidden.
                            RaiseEvent DownloadError(403, "Forbidden.")
                        ElseIf lonStatusCode = 404 Then 'Not found.
                            RaiseEvent DownloadError(404, "Not found.")
                        ElseIf lonStatusCode = 405 Then 'Method not allowed.
                            RaiseEvent DownloadError(405, "Method not allowed.")
                        ElseIf lonStatusCode = 406 Then 'Not acceptable.
                            RaiseEvent DownloadError(406, "Not acceptable.")
                        Else
                            RaiseEvent DownloadError(lonStatusCode, "Undefined error")
                        End If
                        'Add more here, I'm done. ;)
                        'http://www.w3.org/Protocols/rfc2616/rfc2616-sec10.html
                        
                End Select
                
            Else
                'Haven't received entire header? (rare)
                'It should be processed next time.
                Exit Sub
            End If
            
        ElseIf intHTTPState = intGET Then 'Sent GET request.
        
            'Check for end of header info (vbCrLf & vbCrLf).
            lonDataStart = BytInStr(0, bytBuffer, vbCrLf & vbCrLf)
            
            If lonDataStart >= 0 Then
                lonDataStart = lonDataStart + 4 'Move to end of header (where data starts).
                'Found end of header.
                'Check HTTP status code (404, 200, etc).
                lonStatusCode = HTTPStatusCode(bytBuffer)
                Select Case lonStatusCode
                    Case 200 'HTTP OK
                        'Extract file size, etc. from header.
                        ExtractFromHeader bytBuffer
                        
                        'Open file.
                        If FileExists(strSavePath) Then
                            SafeKill strSavePath
                        End If
                        lonFH = CLng(FreeFile)
                        Open strSavePath For Binary As #lonFH
                        Close #lonFH
                        lonFH = -1
                        'Open strSavePath For Binary Access Write As #lonFH
                        API_OpenFile strSavePath, lonFH, curTemp
                        RaiseEvent DownloadStarted(curFileSize)
                        tmrSpeed.Enabled = True
                        
                        'Check if there is any file data in this packet.
                        If UBound(bytBuffer) > lonDataStart Then
                            'Extract data.
                            lonDataLen = UBound(bytBuffer) - lonDataStart
                            ReDim bytFileData(0 To lonDataLen) As Byte
                            CopyMemory bytFileData(0), bytBuffer(lonDataStart), lonDataLen + 1
                            API_WriteFile lonFH, 0, lonDataLen + 1, bytFileData
                            'Put #lonFH, 1, bytFileData
                            curBytesRec = lonDataLen + 1
                            RaiseEvent DownloadProgress(curBytesRec, curFileSize)
                        End If
                        
                        intHTTPState = intReceivingFile
                        
                        'Done with this packet, erase buffer.
                        Erase bytBuffer
                    
                    Case Else 'Some other response.
                        sckMain.Close
                        intHTTPState = intIdle
                        If lonStatusCode = 400 Then 'Bad request.
                            RaiseEvent DownloadError(400, "Bad request.")
                        ElseIf lonStatusCode = 401 Then 'Unauthorized.
                            RaiseEvent DownloadError(401, "Unauthorized.")
                        ElseIf lonStatusCode = 402 Then 'Payment required.
                            RaiseEvent DownloadError(402, "Payment required.")
                        ElseIf lonStatusCode = 403 Then 'Forbidden.
                            RaiseEvent DownloadError(403, "Forbidden.")
                        ElseIf lonStatusCode = 404 Then 'Not found.
                            RaiseEvent DownloadError(404, "Not found.")
                        ElseIf lonStatusCode = 405 Then 'Method not allowed.
                            RaiseEvent DownloadError(405, "Method not allowed.")
                        ElseIf lonStatusCode = 406 Then 'Not acceptable.
                            RaiseEvent DownloadError(406, "Not acceptable.")
                        Else
                            RaiseEvent DownloadError(lonStatusCode, "Undefined error")
                        End If
                        'Add more here, I'm done. ;)
                        'http://www.w3.org/Protocols/rfc2616/rfc2616-sec10.html
                End Select
            Else
                'Haven't received entire GET response header (rare).
                'Exit and it should be processed next time.
                Exit Sub
            End If
        
        ElseIf intHTTPState = intReceivingFile Then 'Receiving just file data.
            API_WriteFile lonFH, 0, UBound(bytBuffer) + 1, bytBuffer
            'Put #lonFH, CLng(curBytesRec + 1), bytBuffer
            curBytesRec = curBytesRec + UBound(bytBuffer) + 1
            RaiseEvent DownloadProgress(curBytesRec, curFileSize)
            If GetInputState Then DoEvents
            If curBytesRec = curFileSize Then
                RaiseEvent DownloadComplete(curBytesRec, curFileSize)
                API_CloseFile lonFH
                lonFH = -1
                sckMain.Close
                intHTTPState = intFileReceived
                tmrSpeed.Enabled = False
                RaiseEvent DownloadSpeed(0)
            End If
            Erase bytBuffer
        End If
                        
        
    ElseIf intProtocol = proFTP Then
        sckMain.GetData strData, vbString, bytesTotal
        If Len(strData) > 3 Then
            If IsNumeric(Left$(strData, 3)) Then
                intCMD = CInt(Left$(strData, 3))
                
                Select Case intCMD
                    Case 213 'Response to SIZE command.
                        WaitSend
                        strSize = Trim$(Replace(Mid$(strData, 5), vbCrLf, ""))
                        curFileSize = CCur(strSize)
                        RaiseEvent DownloadStarted(curFileSize)
                        sckMain.SendData "PASV" & vbCrLf
                        bolFTPSending = True
                        intFTPState = ftpSIZE
                    Case 220 'Welcome message.
                        WaitSend
                        sckMain.SendData "USER " & strUsername & vbCrLf
                        bolFTPSending = True
                        intFTPState = FTPUser
                    Case 331 'Request for password.
                        WaitSend
                        sckMain.SendData "PASS " & strPassword & vbCrLf
                        bolFTPSending = True
                        intFTPState = FTPPass
                    Case 230 'Logged in.
                        WaitSend
                        sckMain.SendData "SIZE " & strURI & vbCrLf
                        bolFTPSending = True
                        intFTPState = ftpSIZE
                    Case 226 'File transfer OK.
                        'sckFTP.Close
                        'API_CloseFile lonFH
                        'RaiseEvent DownloadComplete(curBytesRec, curFileSize)
                    Case 227 'Entering passive mode.
                        PassiveInfo strData, strPHost, lonPPort
                        With sckFTP
                            .Close
                            .RemoteHost = strPHost
                            .RemotePort = lonPPort
                            .Connect
                        End With
                        intFTPState = ftpRETR
                    Case 530 'Invalid username/password.
                        RaiseEvent DownloadError(530, "Invalid username/password.")
                        sckMain.Close
                        sckFTP.Close
                    Case 550 'Error.
                        RaiseEvent DownloadError(550, Mid$(strData, 5))
                        sckMain.Close
                        sckFTP.Close
                End Select
            End If
        End If
    End If
    Erase bytFileData
    Exit Sub
ErrorHandler:
    If Err.Number = 40006 Then
        sckMain.Close
        Erase bytData
        Erase bytFileData
        tmrSpeed.Enabled = False
        sckFTP.Close
        If lonFH <> -1 Then API_CloseFile lonFH
        RaiseEvent SocketClose
    End If
End Sub

'FTP server provides port with 2 numbers for passive mode.
'ie: Entering passive mode (127,0,0,1,255,255)
'127,0,0,1 is the IP address
'255,255 represents a single port number.
'This function returns the port number (unsigned INT) from those 2 values.
Private Function GetPort(ByRef LeftValue As String, ByRef RightValue As String) As Long
    GetPort = (LeftValue * 256) + RightValue
End Function

'This takes the data from the: Entering passive mode (127,0,0,1,255,255) packet
'and extracts the IP and port.
Private Sub PassiveInfo(ByRef Data As String, _
    ByRef IPAddress As String, ByRef Port As Long)
    
    Dim l As Long, e As Long, strTemp As String
    Dim strInfo() As String
    
    'Entering Passive Mode (127,0,0,1,36,67)
    l = InStr(1, Data, "(")
    If l > 0 Then
        l = l + 1
        e = InStr(l, Data, ")")
        If e > 0 Then
            strTemp = Mid$(Data, l, e - l)
            strInfo = Split(strTemp, ",")
            IPAddress = strInfo(0) & "." & strInfo(1) & "." & strInfo(2) & "." & strInfo(3)
            Port = CLng(GetPort(strInfo(4), strInfo(5)))
        End If
    End If
    Erase strInfo
End Sub

Private Sub SafeKill(ByRef Path As String)
    On Error Resume Next
    Kill Path
    On Error GoTo 0 'Not really needed...
End Sub

'Extract an HTTP field from an HTTP header stored in a byte array. ex:
'Content-type: text/html
'Content-length: 123456

'HTTPField(Data(), "content-length:") = " 123456"
'LTrim$() may be needed after.
'FieldName is case-INsensitive.
Private Function HTTPField(ByRef Data() As Byte, ByRef FieldName As String) As String
    Dim i As Integer, e As Integer
    Dim bytRet() As Byte, lonLen As Long
    
    i = BytInStr(0, Data, FieldName, vbTextCompare)
    If i >= 0 Then
        i = i + Len(FieldName)
        e = BytInStr(i, Data, vbCrLf)
        If e > 0 Then
            lonLen = e - i
            ReDim bytRet(0 To lonLen - 1) As Byte
            CopyMemory bytRet(0), Data(i), lonLen
        End If
    End If
    HTTPField = StrConv(bytRet, vbUnicode)
    Erase bytRet
End Function

'Get the upper-boundaries of a byte array.
'Returns -1 on error (array not initialized).
Private Function Byte_UBound(ByRef Data() As Byte) As Long
    On Error GoTo ErrorHandler
    Byte_UBound = UBound(Data)
    Exit Function
ErrorHandler:
    Byte_UBound = -1
    Exit Function
End Function

'Simply compares 2 byte arrays and returns TRUE if they are the same.
'Can also compare case-sensitive or case-insensitive.
'Used for the BytInStr() function.
Private Function ArraysSame(ByRef Array1() As Byte, ByRef Array2() As Byte, _
    Optional ByVal CompareMethod As VbCompareMethod = vbBinaryCompare) As Boolean
    Dim bolRet As Boolean, l As Long, u As Long
    Dim u2 As Long
    bolRet = True
    u = Byte_UBound(Array1)
    u2 = Byte_UBound(Array2)
    
    If u <> u2 Then
        ArraysSame = False
        Exit Function
    Else
        For l = 0 To u
            If CompareMethod = vbBinaryCompare Then
                If Array1(l) <> Array2(l) Then
                    bolRet = False
                    Exit For
                End If
            ElseIf CompareMethod = vbTextCompare Then
                If BytLCaseSingle(Array1(l)) <> BytLCaseSingle(Array2(l)) Then
                    bolRet = False
                    Exit For
                End If
            End If
        Next l
    End If
    ArraysSame = bolRet
End Function

'Convert a byte value to its lcase (lowercase) equivalent.
Private Function BytLCaseSingle(ByVal Data As Byte) As Byte
    Select Case Data
        Case 65 To 90 'A-Z
            BytLCaseSingle = Data + 32
        Case Else
            BytLCaseSingle = Data
    End Select
End Function

'InStr() for byte array.
Private Function BytInStr(ByVal Start As Long, ByRef Data() As Byte, _
    ByRef FindWhat As String, _
    Optional ByVal CompareMethod As VbCompareMethod = vbBinaryCompare) As Long
    
    Dim l As Long, u As Long, bytTest() As Byte
    Dim bytFind() As Byte, lonLenFind As Long
    Dim lonRet As Long
    lonRet = -1
    u = Byte_UBound(Data)
    If u >= 0 And Start <= u Then
        bytFind = StrConv(FindWhat, vbFromUnicode)
        lonLenFind = UBound(bytFind)
        ReDim bytTest(0 To lonLenFind) As Byte
        For l = Start To u
            If CompareMethod = vbBinaryCompare Then
                If Data(l) = bytFind(0) Then
                    CopyMemory bytTest(0), Data(l), lonLenFind + 1
                    If ArraysSame(bytTest, bytFind) Then
                        lonRet = l
                        Exit For
                    End If
                End If
            ElseIf CompareMethod = vbTextCompare Then
                If BytLCaseSingle(Data(l)) = BytLCaseSingle(bytFind(0)) Then
                    CopyMemory bytTest(0), Data(l), lonLenFind + 1
                    If ArraysSame(bytTest, bytFind, vbTextCompare) Then
                        lonRet = l
                        Exit For
                    End If
                End If
            End If
        Next l
    End If
    Erase bytTest
    Erase bytFind
    BytInStr = lonRet
End Function

'Get status code from HTTP response.
'ex: 200 (OK) 404 (Not found), etc.
Public Function HTTPStatusCode(ByRef Data() As Byte) As Long
    Dim i As Long, e As Long
    Dim bytCode() As Byte, lonLen As Long
    i = BytInStr(0, Data, " ")
    If i >= 0 Then
        i = i + 1
        e = BytInStr(i, Data, " ")
        If e >= 0 Then
            lonLen = e - i
            ReDim bytCode(0 To lonLen) As Byte
            CopyMemory bytCode(0), Data(i), lonLen
            HTTPStatusCode = CLng(StrConv(bytCode, vbUnicode))
        End If
    End If
    Erase bytCode
End Function

'Tells the program we're done sending a packet to the FTP server.
Private Sub sckMain_SendComplete()
    If intProtocol = proFTP Then
        bolFTPSending = False
    End If
End Sub

'Calculate bytes per second of transfer.
Private Sub tmrSpeed_Timer()
    sinDownloadSpeed = CSng(curBytesRec - curDownloadSec)
    RaiseEvent DownloadSpeed(sinDownloadSpeed)
    curDownloadSec = curBytesRec
End Sub

'Initialize default property values.
Private Sub UserControl_InitProperties()
    MaxRedirects = DEF_MAX_REDIRECTS
    UserAgent = DEF_USER_AGENT
End Sub

'Read the properties.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        MaxRedirects = .ReadProperty("MaxRedirects", DEF_MAX_REDIRECTS)
        UserAgent = .ReadProperty("UserAgent", DEF_USER_AGENT)
    End With
End Sub

Private Sub UserControl_Resize()
    With UserControl
        .Width = 32 * Screen.TwipsPerPixelX
        .Height = 32 * Screen.TwipsPerPixelY
    End With
End Sub

'Cleanup.
Private Sub UserControl_Terminate()
    sckMain.Close
    sckFTP.Close
    tmrSpeed.Enabled = False
    If lonFH <> -1 Then API_CloseFile lonFH
    Erase bytBuffer
End Sub

'Write the properties.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "MaxRedirects", p_MaxRedirects, DEF_MAX_REDIRECTS
        .WriteProperty "UserAgent", p_UserAgent, DEF_USER_AGENT
    End With
End Sub
