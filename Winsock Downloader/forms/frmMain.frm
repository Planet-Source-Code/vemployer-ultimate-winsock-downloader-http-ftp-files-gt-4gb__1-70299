VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Winsock Downloader"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   234
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   384
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGo 
      Caption         =   "&GO"
      Height          =   255
      Left            =   2400
      TabIndex        =   17
      Top             =   120
      Width           =   615
   End
   Begin VB.OptionButton optFTP 
      Caption         =   "FTP"
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   120
      Width           =   975
   End
   Begin VB.OptionButton optHTTP 
      Caption         =   "HTTP"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.Frame fraFTP 
      Caption         =   " FTP Download "
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   5535
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Text            =   "/public_html/file.txt"
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txtPass 
         Height          =   285
         Left            =   3480
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   3480
         TabIndex        =   2
         Text            =   "21"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remote file:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pass:"
         Height          =   195
         Index           =   3
         Left            =   3000
         TabIndex        =   11
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Port:"
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   9
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Host:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   450
      End
   End
   Begin VB.Frame fraHTTP 
      Caption         =   "HTTP Download"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5535
      Begin VB.TextBox txtURL 
         Height          =   285
         Left            =   600
         TabIndex        =   6
         Text            =   "http://www.domain.com/file.ext"
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "URL:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   405
      End
   End
   Begin MSComDlg.CommonDialog objCD 
      Left            =   2160
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TestDownloader.FlatProgress objBar 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
   End
   Begin VB.Label lblSpeed 
      AutoSize        =   -1  'True
      Caption         =   "0 bytes/sec"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   1170
   End
   Begin TestDownloader.ctlDownloader objDownload 
      Left            =   5160
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Ready."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   660
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGo_Click()
    On Error GoTo ErrorHandler
    Dim strName As String, strExt As String
    
    If optHTTP.Value Then
        txtURL.Text = LTrim$(txtURL.Text)
        If Len(txtURL.Text) > 0 Then
            With objCD
                .CancelError = True
                .Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
                .DialogTitle = "Save File"
                strName = GetAfterLast(txtURL.Text, "/")
                strExt = GetAfterLast(strName, ".")
                .InitDir = App.Path
                .FileName = strName
                .Filter = UCase$(strExt) & " Files|*." & LCase$(strExt) & "|All Files (*.*)|*.*"
                .ShowSave
                lblStatus.Caption = "Connecting..."
                objDownload.DownloadHTTP txtURL.Text, .FileName
            End With
            objBar.Value = 0
        End If
    Else
        txtHost.Text = Trim$(txtHost.Text)
        txtPort.Text = Trim$(txtPort.Text)
        txtUser.Text = Trim$(txtUser.Text)
        txtPath.Text = Trim$(txtPath.Text)
        If Len(txtHost.Text) > 0 Then
            If Len(txtPort.Text) > 0 Then
                If Len(txtUser.Text) > 0 Then
                    If Len(txtPath.Text) > 0 Then
                        With objCD
                            .CancelError = True
                            .Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
                            .DialogTitle = "Save File"
                            strName = GetAfterLast(txtPath.Text, "/")
                            strExt = GetAfterLast(strName, ".")
                            .InitDir = App.Path
                            .FileName = strName
                            .Filter = UCase$(strExt) & " Files|*." & LCase$(strExt) & "|All Files (*.*)|*.*"
                            .ShowSave
                            lblStatus.Caption = "Connecting..."
                        End With
                        objBar.Value = 0
                        objDownload.DownloadFTP txtHost.Text, txtUser.Text, txtPass.Text, objCD.FileName, txtPath.Text, txtPort.Text
                    Else
                        MsgBox "Enter a file name to download!", vbExclamation
                        txtPath.SetFocus
                    End If
                Else
                    MsgBox "Enter the username of the FTP account!", vbExclamation
                    txtUser.SetFocus
                End If
            Else
                MsgBox "Enter the FTP port!", vbExclamation
                txtPort.SetFocus
            End If
        Else
            MsgBox "Enter the FTP hostname or IP address!", vbExclamation
            txtHost.SetFocus
        End If
    End If
    Exit Sub
    
ErrorHandler:
    Exit Sub
End Sub

Private Sub Form_Load()
    optHTTP_Click
End Sub

Private Sub objDownload_DownloadComplete(ByVal BytesReceived As Currency, ByVal FileSize As Currency)
    lblStatus.Caption = "Download complete!"
    objBar.Value = objBar.Max
End Sub

Private Sub objDownload_DownloadError(ByVal Number As Long, Description As String)
    lblStatus.Caption = Description
End Sub

Private Sub objDownload_DownloadProgress(ByVal BytesReceived As Currency, ByVal FileSize As Currency)
    On Error Resume Next
    With objBar
        .Value = CLng(BytesReceived)
    End With
End Sub

Private Sub objDownload_DownloadRedirect(ByVal Location As String)
    lblStatus.Caption = "Redirecting..."
End Sub

Private Sub objDownload_DownloadSpeed(ByVal BytesPerSecond As Single)
    lblSpeed.Caption = FormatSize(CCur(BytesPerSecond)) & "/sec"
End Sub

Private Sub objDownload_DownloadStarted(ByVal FileSize As Currency)
    lblStatus.Caption = "Receiving file..."
    On Error Resume Next
    With objBar
        .Max = CDbl(FileSize)
        .Min = 0
        .Value = 0
    End With
End Sub

Private Sub objDownload_SocketClose()
    'lblStatus.Caption = "Connection remotely closed."
End Sub

Private Sub objDownload_SocketConnect()
    lblStatus.Caption = "Requesting file..."
End Sub

Private Function GetAfterLast(ByRef Text As String, ByRef AfterWhat As String) As String
    Dim l As Long
    l = InStrRev(Text, AfterWhat)
    If l > 0 Then
        GetAfterLast = Mid$(Text, l + 1)
    End If
End Function

'Thanks Ellis Dee @ www.vbforums.com :)
Private Function FormatSize(ByVal Size As Currency) As String
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

Private Sub objDownload_SocketError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    lblStatus.Caption = Description
End Sub

Private Sub optFTP_Click()
    fraFTP.Visible = True
    fraHTTP.Visible = False
End Sub

Private Sub optHTTP_Click()
    fraHTTP.Visible = True
    fraFTP.Visible = False
End Sub
