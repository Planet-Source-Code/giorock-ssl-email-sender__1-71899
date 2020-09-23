VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Email Settings"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CheckBox chkSSL 
      Caption         =   "Needs SSL"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtUserID 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   4770
      Y1              =   2265
      Y2              =   2265
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   660
      TabIndex        =   14
      Top             =   1320
      Width           =   465
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   4680
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "From Address:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   435
      TabIndex        =   10
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "User ID:"
      Height          =   195
      Left            =   585
      TabIndex        =   9
      Top             =   2400
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Port:"
      Height          =   195
      Left            =   840
      TabIndex        =   8
      Top             =   600
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server Name:"
      Height          =   195
      Left            =   195
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function EncryptPass(strPass As String) As String
    Dim pstrtmp As String
    Dim pintPos As Integer
    Dim pintC As Integer
    
    For pintC = 1 To Len(strPass)
      pstrtmp = Right(strPass, 1)
      EncryptPass = EncryptPass & Asc(pstrtmp) + 3 & "-"
      strPass = Left(strPass, Len(strPass) - 1)
    Next
    
End Function

Public Function DecryptPass(strPass As String) As String
    Dim pstrtmp As String
    Dim pintPos As Integer
    Dim pintC As Integer
    
    pintPos = InStr(strPass, "-")
    Do Until pintPos = 0
      pstrtmp = Left(strPass, pintPos)
      DecryptPass = DecryptPass & Chr(Val(pstrtmp) - 3)
      strPass = Replace(strPass, pstrtmp, "", 1, 1)
      pintPos = InStr(strPass, "-")
    Loop
    
    DecryptPass = StrReverse(DecryptPass)
    
End Function

Private Sub chkSSL_Click()
    chkSSL.Value = vbChecked
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim pstrPass As String
    
    pstrPass = EncryptPass(Trim(txtPass.Text))
    
    SaveSetting App.Title, "SrvConfig", "ServerName", Trim(txtServer.Text)
    SaveSetting App.Title, "SrvConfig", "UserID", Trim(txtUserID.Text)
    SaveSetting App.Title, "SrvConfig", "UserPass", pstrPass
    SaveSetting App.Title, "SrvConfig", "Port", Trim(txtPort.Text)
    SaveSetting App.Title, "SrvConfig", "UserName", Trim(txtName.Text)
    SaveSetting App.Title, "SrvConfig", "SSL", IIf(chkSSL.Value = vbChecked, True, False)
    SaveSetting App.Title, "SrvConfig", "FromAddress", Trim(txtFrom.Text)
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Dim pstrSSL As String
    
    On Error Resume Next
    
    txtServer.Text = GetSetting(App.Title, "SrvConfig", "ServerName", "smtp.live.com")
    txtUserID.Text = GetSetting(App.Title, "SrvConfig", "UserID", "xxxxxx@live.it")
    txtPass.Text = GetSetting(App.Title, "SrvConfig", "UserPass", "")
    txtPass.Text = DecryptPass(txtPass.Text)
    txtPort.Text = GetSetting(App.Title, "SrvConfig", "Port", "587")
    pstrSSL = GetSetting(App.Title, "SrvConfig", "SSL", "True")
    txtName.Text = GetSetting(App.Title, "SrvConfig", "UserName", "GioRock")
    txtFrom.Text = GetSetting(App.Title, "SrvConfig", "FromAddress", "xxxxxx@live.it")
    
    If pstrSSL = "True" Then
        chkSSL.Value = vbChecked
    Else
        chkSSL.Value = vbUnchecked
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSettings = Nothing
End Sub


