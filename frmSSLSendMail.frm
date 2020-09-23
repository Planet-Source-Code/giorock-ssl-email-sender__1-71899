VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SSLForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SSL eMail Sender - Updated 2012"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9060
   Icon            =   "frmSSLSendMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   4995
      TabIndex        =   1
      Top             =   5655
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   1
      Min             =   1e-4
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   4800
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4320
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame frameEmail 
      Height          =   5505
      Left            =   30
      TabIndex        =   3
      Top             =   45
      Width           =   8985
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   5340
         Left            =   75
         ScaleHeight     =   5340
         ScaleWidth      =   8865
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   135
         Width           =   8865
         Begin VB.CommandButton cmdSend 
            Caption         =   "&Send"
            Height          =   375
            Left            =   6960
            TabIndex        =   21
            Top             =   885
            Width           =   1815
         End
         Begin VB.TextBox txtSubject 
            Height          =   285
            Left            =   735
            TabIndex        =   12
            Text            =   "Testing SSL connection"
            Top             =   540
            Width           =   5655
         End
         Begin VB.TextBox txtTo 
            Height          =   285
            Left            =   735
            TabIndex        =   11
            Text            =   "giorock@teletu.it"
            Top             =   900
            Width           =   5655
         End
         Begin VB.TextBox txtFrom 
            Height          =   285
            Left            =   735
            TabIndex        =   10
            Top             =   180
            Width           =   5655
         End
         Begin VB.TextBox txtBCC 
            Height          =   285
            Left            =   735
            TabIndex        =   9
            Text            =   "giorock@libero.it"
            Top             =   1620
            Width           =   5655
         End
         Begin VB.TextBox txtCC 
            Height          =   285
            Left            =   735
            TabIndex        =   8
            Text            =   "rockadmin@teletu.it"
            Top             =   1260
            Width           =   5655
         End
         Begin VB.TextBox txtMessage 
            Height          =   3000
            Left            =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Text            =   "frmSSLSendMail.frx":0622
            Top             =   2220
            Width           =   8520
         End
         Begin VB.CommandButton cmdSetting 
            Caption         =   "&Setting"
            Height          =   375
            Left            =   6960
            TabIndex        =   6
            Top             =   1365
            Width           =   1815
         End
         Begin VB.CommandButton cmdAttachments 
            Caption         =   "Add &Attachment"
            Height          =   375
            Left            =   6960
            TabIndex        =   5
            Top             =   405
            Width           =   1815
         End
         Begin VB.Image Image2 
            Height          =   240
            Left            =   6630
            Picture         =   "frmSSLSendMail.frx":0692
            Top             =   952
            Width           =   240
         End
         Begin VB.Line Line2 
            BorderColor     =   &H8000000C&
            X1              =   6660
            X2              =   8730
            Y1              =   1860
            Y2              =   1860
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "GioRock 2012"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   6525
            MouseIcon       =   "frmSSLSendMail.frx":0CB4
            MousePointer    =   99  'Custom
            TabIndex        =   20
            ToolTipText     =   "go to Author Site..."
            Top             =   60
            Width           =   2235
         End
         Begin VB.Image Image3 
            Height          =   240
            Left            =   6630
            Picture         =   "frmSSLSendMail.frx":0E06
            Top             =   1432
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   6630
            Picture         =   "frmSSLSendMail.frx":1428
            Top             =   472
            Width           =   240
         End
         Begin VB.Label lblAttachments 
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1320
            TabIndex        =   19
            Top             =   1965
            Width           =   7455
         End
         Begin VB.Image Image11 
            Height          =   240
            Left            =   -15
            Picture         =   "frmSSLSendMail.frx":1A4A
            Top             =   3600
            Width           =   240
         End
         Begin VB.Image Image10 
            Height          =   240
            Left            =   -15
            Picture         =   "frmSSLSendMail.frx":206C
            Top             =   1972
            Width           =   240
         End
         Begin VB.Image Image9 
            Height          =   240
            Left            =   -15
            Picture         =   "frmSSLSendMail.frx":268E
            Top             =   562
            Width           =   240
         End
         Begin VB.Image Image8 
            Height          =   240
            Left            =   -15
            Picture         =   "frmSSLSendMail.frx":2CB0
            Top             =   1642
            Width           =   240
         End
         Begin VB.Image Image7 
            Height          =   240
            Left            =   -15
            Picture         =   "frmSSLSendMail.frx":32D2
            Top             =   1282
            Width           =   240
         End
         Begin VB.Image Image6 
            Height          =   240
            Left            =   -15
            Picture         =   "frmSSLSendMail.frx":36F4
            Top             =   922
            Width           =   240
         End
         Begin VB.Image Image5 
            Height          =   240
            Left            =   -15
            Picture         =   "frmSSLSendMail.frx":3D16
            Top             =   202
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "To:"
            Height          =   195
            Left            =   420
            TabIndex        =   18
            Top             =   945
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Sub:"
            Height          =   195
            Left            =   330
            TabIndex        =   17
            Top             =   585
            Width           =   330
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "From:"
            Height          =   195
            Left            =   270
            TabIndex        =   16
            Top             =   225
            Width           =   390
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "BCC:"
            Height          =   195
            Left            =   300
            TabIndex        =   15
            Top             =   1665
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "CC:"
            Height          =   195
            Left            =   405
            TabIndex        =   14
            Top             =   1305
            Width           =   255
         End
         Begin VB.Label Label8 
            Caption         =   "Attachments:"
            Height          =   255
            Left            =   315
            TabIndex        =   13
            Top             =   1965
            Width           =   975
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000009&
            BorderWidth     =   2
            X1              =   6630
            X2              =   8730
            Y1              =   1875
            Y2              =   1875
         End
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5595
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8643
            MinWidth        =   8643
            Text            =   "Status: Ready to send an eMail"
            TextSave        =   "Status: Ready to send an eMail"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8643
            MinWidth        =   8643
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   480
      Left            =   6810
      TabIndex        =   2
      Top             =   1935
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmSSLSendMail.frx":4338
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   6495
      Picture         =   "frmSSLSendMail.frx":43BA
      Top             =   2055
      Width           =   240
   End
End
Attribute VB_Name = "SSLForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************
'*      SSL eMail Sender        *
'********************************
'*   Created by GioRock 2009    *
'*     giorock@libero.it        *
'********************************

'********************************************
'*               UPDATED  2012              *
'********************************************

'******************NEWS**********************
'*    NOW RSA KEY CERTIFIED V3 SUPPORTED    *
'*     ENABLE TO HANDSHAKE WITH SSL3.0      *
'*     SERVER MUST BE ABLE TO FALLBACK      *
'*       AT SSL2.0 PROTOCOL INTERFACE       *
'********************************************

'WARNING:
'THIS PROGRAM IS ONLY TESTED ON HOTMAIL SERVER
'PARAMS:
'       SERVER NAME
'       PORT
'       USERID
'       PASSWORD
'       ACCESS BY SERVER AUTHENTICATION
'       SSL -> TRUE
'NOT DISCLAIMS ARE ACCEPTED USING OTHER CONFIGURATIONS OR DIFFERENT SERVERS

'YOU CAN ADAPT ALL CODE TO WORK ON OTHER SERVERS

'TODO:
'MUCH MORE....

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const CS_DROPSHADOW As Long = &H20000
Private Const GCL_STYLE     As Long = -26

Private Mail As SSLSocket

Private Sub ApplyDropShadow(ByVal hWnd As Long)
    Me.Hide
    DoEvents
    Call SetClassLong(hWnd, GCL_STYLE, GetClassLong(hWnd, GCL_STYLE) Or CS_DROPSHADOW)
    Me.Show
End Sub

Private Function MyDocuments() As String
    Dim r As Long
    Dim IDL As ITEMIDLIST
    Dim Path As String
    'Get the special folder
    r = SHGetSpecialFolderLocation(100, &H5, IDL)
    If r = 0 Then
        'Create a buffer
        Path$ = Space$(512)
        'Get the path from the IDList
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        'Remove the unnecessary chr$(0)'s
        MyDocuments = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If

End Function
Private Sub cmdAttachments_Click()
    Dim pstrFolder As String
    Dim pstrFilePath As String
    Dim pstrFile As String
    Dim pintI As Integer
    Dim pstrLabel As String
    
    Screen.MousePointer = vbHourglass
    pstrFolder = MyDocuments
    
    On Error GoTo ErrorCall
    
    With dlgFile
        .CancelError = True
        .InitDir = MyDocuments
        .ShowOpen
        pstrFilePath = .Filename
    End With
    
    StatusBar1.Panels(1).Text = "Status: Adding File..."
    DoEvents
    
    If Mail Is Nothing Then: Set Mail = New SSLSocket
    
    ProgressBar1.Visible = True
    Call Mail.AddAttachment(pstrFilePath, ProgressBar1)
    
    pintI = InStrRev(pstrFilePath, "\")
    pstrFile = Mid(pstrFilePath, pintI + 1)
    pstrLabel = lblAttachments.Caption
    
    If pstrLabel <> "" Then
        pstrLabel = pstrLabel & ", "
    End If
    
    txtMessage.Top = lblAttachments.Top + lblAttachments.Height
    pstrLabel = pstrLabel & pstrFile
    lblAttachments.Caption = pstrLabel
    ProgressBar1.Visible = False
    
    Screen.MousePointer = vbNormal
    
    Exit Sub

ErrorCall:
    If Err.Number = 32755 Then
        Screen.MousePointer = vbNormal
        Exit Sub
    Else
        Call MsgBox(Err.Number & "' " & Err.Description, vbCritical + vbOKOnly, "Error")
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
End Sub


Private Sub cmdSend_Click()
    Dim pstrFrom As String
    Dim pstrTo() As String
    Dim pstrBCC() As String
    Dim pstrCC() As String
    Dim plngPort As Long
    Dim pstrServer As String
    Dim pblnSSL As Boolean
    Dim pstrSSL As String
    Dim pstrPass As String
    Dim pstrUserID As String
    Dim pintI As Integer
    Dim pstrtmp As String
    Dim pintC As Integer
    Dim pstrFromName As String
    Dim pintSendMailRet As Integer
    
    pstrFrom = GetSetting(App.Title, "SrvConfig", "FromAddress", "xxxxxx@live.it")
    pstrServer = GetSetting(App.Title, "SrvConfig", "ServerName", "smtp.live.com")
    If pstrServer = "" Then
        frmSettings.Show vbModal, Me
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    plngPort = GetSetting(App.Title, "SrvConfig", "Port", "587")
    pstrSSL = GetSetting(App.Title, "SrvConfig", "SSL", "True")
    pstrPass = GetSetting(App.Title, "SrvConfig", "UserPass", "******")
    pstrPass = frmSettings.DecryptPass(pstrPass)
    pstrUserID = GetSetting(App.Title, "SrvConfig", "UserID", "xxxxxx@live.it")
    pstrFromName = GetSetting(App.Title, "SrvConfig", "UserName", "GioRock")
    
    StatusBar1.Panels(1).Text = "Status: Sending eMail..."
    
    cmdAttachments.Enabled = False
    cmdSend.Enabled = False
    cmdSetting.Enabled = False
    
    If LCase(pstrSSL) = "true" Then
        pblnSSL = True
    Else
        pblnSSL = False
    End If
    
    If txtSubject.Text = "" Then
        txtSubject.Text = "(none)"
    End If
    
    ReDim pstrTo(0)
    ReDim pstrBCC(0)
    ReDim pstrCC(0)
    
    
    pintI = 0
    If txtTo.Text <> "" Then
        pstrtmp = txtTo.Text
        Do Until pstrtmp = ""
            pintC = InStr(pstrtmp, ", ")
            If pintC = 0 And pstrtmp <> "" Then
                ReDim Preserve pstrTo(pintI)
                pstrTo(pintI) = pstrtmp
                Exit Do
            End If
            ReDim Preserve pstrTo(pintI)
            pstrTo(pintI) = Left(pstrtmp, pintC - 1)
            pstrtmp = Replace(pstrtmp, pstrTo(pintI) & ", ", "")
            pintI = pintI + 1
        Loop
        
    End If
    
    pintI = 0
    If txtBCC.Text <> "" Then
        pstrtmp = txtBCC.Text
        Do Until pstrtmp = ""
            pintC = InStr(pstrtmp, ", ")
            If pintC = 0 And pstrtmp <> "" Then
                ReDim Preserve pstrBCC(pintI)
                pstrBCC(pintI) = pstrtmp
                Exit Do
            End If
            ReDim Preserve pstrBCC(pintI)
            pstrBCC(pintI) = Left(pstrtmp, pintC - 1)
            pstrtmp = Replace(pstrtmp, pstrBCC(pintI) & ", ", "")
            pintI = pintI + 1
        Loop
        
    End If
    
    pintI = 0
    If txtCC.Text <> "" Then
        pstrtmp = txtCC.Text
        Do Until pstrtmp = ""
            pintC = InStr(pstrtmp, ", ")
            If pintC = 0 And pstrtmp <> "" Then
                ReDim Preserve pstrCC(pintI)
                pstrCC(pintI) = pstrtmp
                Exit Do
            End If
            ReDim Preserve pstrTo(pintI)
            pstrCC(pintI) = Left(pstrtmp, pintC - 1)
            pstrtmp = Replace(pstrtmp, pstrCC(pintI) & ", ", "")
        Loop
    End If
    
    If Mail Is Nothing Then: Set Mail = New SSLSocket

    With Mail
        ProgressBar1.Visible = True
        Call .SetUp(pstrServer, plngPort, _
            pstrFrom, pblnSSL, pstrUserID, _
            pstrPass)
        pintSendMailRet = .SendEmail(pstrFromName, txtSubject.Text, _
            txtMessage.Text, pstrTo, pstrBCC, _
            pstrCC, Winsock1, RichTextBox1, ProgressBar1)
        ProgressBar1.Visible = False
    End With
    
    If pintSendMailRet <> 0 Then
        Dim intRes As Integer
        StatusBar1.Panels(1).Text = "Status: Failed sending eMail(s)"
        intRes = MsgBox("There Is No Connection To The Internet.  Connect Now?", vbYesNo, "Not Connected")
        cmdAttachments.Enabled = True
        cmdSend.Enabled = True
        cmdSetting.Enabled = True
        lblAttachments.Caption = ""
        If intRes = VbMsgBoxResult.vbYes Then
            intRes = Mail.ConnectToInternet
            If intRes <> 0 Then: Set Mail = Nothing
            Select Case intRes
                Case 87
                    Call MsgBox("Bad Parameter for InternetDial - Couldn't Connect", vbOKOnly, "Connection Failed")
                Case 668
                    Call MsgBox("No Connection for InternetDial - Couldn't Connect", vbOKOnly, "Connection Failed")
                Case 631
                    Call MsgBox("User Cancelled Dialup.", vbOKOnly, "Connection Canceled")
                Case 0
                    Call cmdSend_Click
                Case Else
                    Call MsgBox("Unknown InternetDial Error.", vbOKOnly, "Connection Failed")
            End Select
        Else
            Set Mail = Nothing
        End If
    Else
        Set Mail = Nothing
        cmdAttachments.Enabled = True
        cmdSend.Enabled = True
        cmdSetting.Enabled = True
        lblAttachments.Caption = ""
        StatusBar1.Panels(1).Text = "Status: eMail(s) sent"
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdSetting_Click()
    frmSettings.Show vbModal, Me
End Sub


Private Sub Form_Initialize()
Dim iccex As tagInitCommonControlsEx
    
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_USEREX_CLASSES
    End With
    
    InitCommonControlsEx iccex
    
End Sub


Private Sub Form_Load()
    Call ApplyDropShadow(Me.hWnd)
    txtFrom.Text = GetSetting(App.Title, "SrvConfig", "FromAddress", "xxxxxx@live.it")
    Set StatusBar1.Panels(1).Picture = Image4.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Mail = Nothing
    RemoveDropShadow hWnd
    End
    Set SSLForm = Nothing
End Sub


Private Sub Label1_Click()
    ShellExecute hWnd, "Open", "http://digilander.libero.it/giorock/", vbNullString, App.Path, vbNormalFocus
End Sub

Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
' Added by Giorock
'-----------------------------------------------------------'
Dim lPercProgr As Long
Static TotbyteSent As Long
    If bytesRemaining = 0 Then: TotbyteSent = 0
    TotbyteSent = IIf(bytesRemaining <> 0, TotbyteSent + bytesSent, bytesSent)
    lPercProgr = Int((100 / (TotbyteSent + bytesRemaining)) * TotbyteSent)
    StatusBar1.Panels(1).Text = "Status: Bytes Sent = " & TotbyteSent & " - Remaining = " & bytesRemaining ' & IIf(bytesRemaining <> 0, " - " & lPercProgr & "%", "")
    If bytesRemaining <> 0 Then
        ProgressBar1.Max = 100.0001
        ProgressBar1.Value = lPercProgr + 0.0001
    End If
'-----------------------------------------------------------'
End Sub



Private Sub RemoveDropShadow(ByVal hWnd As Long)
    Me.Hide
    DoEvents
    Call SetClassLong(hWnd, GCL_STYLE, GetClassLong(hWnd, GCL_STYLE) Xor CS_DROPSHADOW)
End Sub
