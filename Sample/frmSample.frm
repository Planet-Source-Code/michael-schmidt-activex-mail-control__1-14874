VERSION 5.00
Begin VB.Form frmSample 
   Caption         =   "Sample"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin MailNotify.VBMail MyBox 
      Left            =   2700
      Top             =   510
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.CommandButton cmdRegistry 
      Caption         =   "Remove Regsettings"
      Height          =   315
      Left            =   1860
      TabIndex        =   14
      Top             =   960
      Width           =   1695
   End
   Begin VB.Timer timMinutes 
      Interval        =   60000
      Left            =   3120
      Top             =   510
   End
   Begin VB.PictureBox picMail 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   450
      Picture         =   "frmSample.frx":0E42
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picMail 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   1140
      Picture         =   "frmSample.frx":1284
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMail 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   1650
      Picture         =   "frmSample.frx":1B4E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   2760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMail 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   2190
      Picture         =   "frmSample.frx":2418
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   2820
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtLog 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "frmSample.frx":285A
      Top             =   1740
      Width           =   2655
   End
   Begin VB.TextBox txtDelay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "5"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1665
   End
   Begin VB.TextBox txtServer 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1665
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1860
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   360
      Width           =   1665
   End
   Begin VB.CommandButton cmdCheckMailbox 
      Caption         =   "Check"
      Height          =   315
      Left            =   2760
      TabIndex        =   4
      Top             =   1740
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mail Server"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   1860
      TabIndex        =   7
      Top             =   120
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- Time Delay in Minutes"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   540
      TabIndex        =   6
      Top             =   1380
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   780
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00DFA684&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   0
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   0
      Top             =   0
      Width           =   3615
   End
   Begin VB.Menu mnuMailMenu 
      Caption         =   "MailMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuSetup 
         Caption         =   "Setup"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckMail 
         Caption         =   "Check Mail"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MailTray As New clsTray     ' - Mail Tray
Dim MinutesElapsed As Integer   ' - Mail Check Timer


'================================
'   Form Load
'================================
Private Sub Form_Load()

    MailTray.ShowIcon Me
    MailTray.ChangeIcon Me, picMail.Item(0)
    
    LoadAppSettings
    
    If Len(txtServer) <> 0 Then
        Me.Hide
        cmdCheckMailbox_Click
    End If

End Sub


'================================
'   Form Resize
'================================
Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Me.Hide

End Sub


'================================
'   Form Unload
'================================
Private Sub Form_Unload(Cancel As Integer)

    MailTray.RemoveIcon Me
    SaveAppSettings

End Sub


'================================
'   Check Mailbox Button
'================================
Private Sub cmdCheckMailbox_Click()

    MyBox.CheckNewMail
    MailTray.ChangeIcon Me, picMail.Item(1)
    MailTray.ChangeToolTip Me, "Checking Mail (" & MyBox.Server & ")"

End Sub


'================================
'   Registry Remove Button
'================================
Private Sub cmdRegistry_Click()

    DeleteAppSettings
    txtUser = ""
    txtPassword = ""
    txtServer = ""
    txtDelay = ""

End Sub


'================================
'   Menu Check Mail
'================================
Private Sub mnuCheckMail_Click()
    
    cmdCheckMailbox_Click

End Sub


'================================
'   Menu Setup
'================================
Private Sub mnuSetup_Click()
    Me.WindowState = vbNormal
    Me.Show
    
End Sub


Private Sub mnuExit_Click()
    Unload Me
End Sub

'================================
'   OBJECT Event New Mail
'================================
Private Sub MYBox_NewMail(NumMsgs As Integer)

    MailTray.ChangeIcon Me, picMail.Item(2)
    MailTray.ChangeToolTip Me, NumMsgs & " New Message(s)!"

End Sub


'================================
'   OBJECT Event Noisy
'================================
Private Sub MYBox_Noisy(POPresponse As String)

    txtLog = POPresponse & vbCrLf & txtLog

End Sub


'================================
'   OBJECT Event No Mail
'================================
Private Sub MYBox_NoMail()

    MailTray.ChangeIcon Me, picMail.Item(0)
    MailTray.ChangeToolTip Me, "No New Mail"

End Sub


'================================
'   OBJECT Event Error
'================================
Private Sub MYBox_SockError(ErrorStats As String)

    MailTray.ChangeIcon Me, picMail.Item(3)
    MailTray.ChangeToolTip Me, ErrorStats

End Sub


'================================
'   Timer
'================================
Private Sub timMinutes_Timer()
    '#############################
    '# Every minute this sub is
    '# called, we simply increment
    '# our counter or reset and
    '# check the mail.
    MinutesElapsed = MinutesElapsed + 1

    If MinutesElapsed = txtDelay Then
        cmdCheckMailbox_Click ' Check Mail
        MinutesElapsed = 0    ' Reset Counter
    End If

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Remember..... The value of X will be different if the icon is minimised
' to the system tray.  The values in this case will be as follows,
'       7680   ' MouseMove
'       7695   ' Left MouseDown
'       7710   ' Left MouseUp
'       7725   ' Left DoubleClick
'       7740   ' Right MouseDown
'       7755   ' Right MouseUp
'       7770   ' Right DoubleClick
If MailTray.bRunningInTray Then          'Check to see if form is in the system tray
    Select Case X                           'If it is, use X to get message value
        Case 7755: PopupMenu Me.mnuMailMenu, vbPopupMenuRightButton
        Case 7725: Me.Show: Me.WindowState = vbNormal
    End Select
End If

End Sub


Private Sub txtPassword_Change()
    MyBox.Password = txtPassword
End Sub
Private Sub txtServer_Change()
    MyBox.Server = txtServer
End Sub
Private Sub txtUser_Change()
    MyBox.User = txtUser
End Sub

