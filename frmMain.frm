VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Battle.net Nick Spoofer"
   ClientHeight    =   2385
   ClientLeft      =   1455
   ClientTop       =   1680
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   8040
   Begin VB.PictureBox picSystemTray 
      Height          =   375
      Left            =   0
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame frmColors 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "&Colors"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   4320
      TabIndex        =   11
      Top             =   120
      Width           =   3615
      Begin MSComctlLib.ImageCombo cmbColors 
         Height          =   330
         Left            =   1200
         TabIndex        =   13
         ToolTipText     =   "Select a color to insert into your nickname. You must enable colors in the Options category."
         Top             =   300
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         Text            =   "Select Color..."
         ImageList       =   "ImageList1"
      End
      Begin VB.Label lblColorsFrame 
         Caption         =   "Colors"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   735
      End
      Begin VB.Label lblInsertColor 
         Alignment       =   1  'Right Justify
         Caption         =   "Insert Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Select a color to insert into your nickname. You must enable colors in the Options category."
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   34
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0614
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0990
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1450
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17F0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmOptions 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "&Options"
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   4320
      TabIndex        =   5
      Top             =   1080
      Width           =   3615
      Begin VB.CheckBox chkDisableChecking 
         Caption         =   "Disable Nickname Checking Interval"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   $"frmMain.frx":1B80
         Top             =   960
         Width           =   3375
      End
      Begin VB.CheckBox chkRetainNickname 
         Caption         =   "Retain Nickname Throughout Session"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   $"frmMain.frx":1C2C
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox chkNoClearNickname 
         Caption         =   "Do Not Clear Nickname Box On Change"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "When this is checked, the nickname box will keep the nickname text until you overwrite it."
         Top             =   480
         Width           =   3135
      End
      Begin VB.CheckBox chkAllowBanned 
         Caption         =   "Allow Colors And Banned Characters"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Enables entering characters (including colors) that are banned by Battle.Net. This may have undesired side effects."
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblOptionsFrame 
         Caption         =   "Options"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdAdvanced 
      Appearance      =   0  'Flat
      Caption         =   "&Advanced >>"
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      ToolTipText     =   "Show/Hide Advanced options."
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   2000
      Left            =   0
      Top             =   2040
   End
   Begin VB.Frame frmModify 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "&Modify"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4095
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   240
         MaxLength       =   21
         TabIndex        =   18
         ToolTipText     =   "Enter the nickname you wish to use in a Battle.Net game. "
         Top             =   390
         Width           =   1935
      End
      Begin VB.CommandButton cmdSpoof 
         Appearance      =   0  'Flat
         Caption         =   "&Spoof Nickname"
         Default         =   -1  'True
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         ToolTipText     =   "Change your nickname to the nickname in the text box to the left."
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblModifyFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Modify"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.Frame frmStatus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "S&tatus"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.Label lblStatusFrame 
         Caption         =   "Status"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No text to display in this view."
         Height          =   255
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Your current Battle.Net Nickname."
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Label lbl 
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Menu mnuST 
      Caption         =   "&SystemTray"
      Visible         =   0   'False
      Begin VB.Menu mnuSTRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuSTEnterNickname 
         Caption         =   "E&nter Nickname..."
      End
      Begin VB.Menu mnuSTExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Autocomplete functions
Dim PopulateNicks As Integer, RetreiveNicks As Integer
' Options functions
Dim DisableChecking As Boolean, NoClearNickname As Boolean, RetainNickname As Boolean, _
    AllowBanned As Boolean, enablebanned As Integer
' Form positions
Dim posTop As Long, posLeft As Long
' Parsing functions
Dim DecodeASCII As Integer, ParsedNick As String, ParseChars As Integer, NewNick As String
' NickName variables
Dim NickResult As String, CurrentNick As String, Restoring As Boolean

Public Sub CreateIcon()
Dim Tic As NOTIFYICONDATA
Tic.cbSize = Len(Tic)
Tic.hwnd = picSystemTray.hwnd
Tic.uID = 1&
Tic.uFlags = NIF_DOALL
Tic.uCallbackMessage = WM_MOUSEMOVE
Tic.hIcon = picSystemTray.Picture
Tic.szTip = "Nick Spoofer" & Chr$(0)
Shell_NotifyIcon NIM_ADD, Tic
End Sub

Public Sub DeleteIcon()
Dim Tic As NOTIFYICONDATA
Tic.cbSize = Len(Tic)
Tic.hwnd = picSystemTray.hwnd
Tic.uID = 1&
Shell_NotifyIcon NIM_DELETE, Tic
End Sub

Private Sub PopulateColors()
cmbColors.ComboItems.Add 1, , "Blue", 1
cmbColors.ComboItems.Add 2, , "Green", 2
cmbColors.ComboItems.Add 3, , "Highlighted Green", 3
cmbColors.ComboItems.Add 4, , "Black", 4
cmbColors.ComboItems.Add 5, , "White", 5
cmbColors.ComboItems.Add 6, , "Red", 6
cmbColors.ComboItems.Add 7, , "Invisble"
End Sub
Private Sub LoadLastPos()
posTop = GetSetting("NameSpoofer", "Position", "Top", "0")
posLeft = GetSetting("NameSpoofer", "Position", "Left", "0")
If posTop > Screen.Height Or posLeft > Screen.Width Then Exit Sub
If posTop > 0 And posLeft > 0 Then frmMain.Top = Int(posTop): frmMain.Left = posLeft
End Sub
Private Sub ExpandApp()
    frmMain.Width = 8130
    frmColors.Visible = True
    frmOptions.Visible = True
    cmdAdvanced.Caption = "<< &Basic"
    SaveSetting "NameSpoofer", "Position", "Expanded", "1"
End Sub
Private Sub ShrinkApp()
    frmMain.Width = 4455
    frmColors.Visible = False
    frmOptions.Visible = False
    cmdAdvanced.Caption = "&Advanced >>"
    SaveSetting "NameSpoofer", "Position", "Expanded", "0"
    'Super resource saver: Hide hidden controls
End Sub
Private Sub GetAppSettings()

'Fill the option boxes with the appropriate options, etc
DisableChecking = GetSetting("NameSpoofer", vbNullChar, "DisableInterval", False)
NoClearNickname = GetSetting("NameSpoofer", vbNullChar, "DontClearTextbox", False)
' Since "DoClearTextbox" is the opposite of "NoClearNickname", reverse. In for backwards
' compatibility with Nick Spoofer version 1.04
'If NoClearNickname = 1 Then NoClearNickname = 0 Else NoClearNickname = 1
RetainNickname = GetSetting("NameSpoofer", vbNullChar, "DoRetainName", False)
AllowBanned = GetSetting("NameSpoofer", vbNullChar, "DoShowColors", False)

If DisableChecking Then
    chkDisableChecking.Value = 1
Else
    chkDisableChecking.Value = 0
End If

If NoClearNickname Then
    chkNoClearNickname.Value = 1
Else
    chkNoClearNickname.Value = 0
End If

If RetainNickname Then
    chkRetainNickname.Value = 1
Else
    chkRetainNickname.Value = 0
End If

If AllowBanned Then
    cmbColors.Enabled = True
    chkAllowBanned.Value = 1
Else
    cmbColors.Enabled = False
    chkAllowBanned.Value = 0
End If

If chkDisableChecking.Value = 1 Then chkRetainNickname.Value = 0: chkRetainNickname.Enabled = False
End Sub
Private Sub SaveAppSettings()
If chkDisableChecking.Value = 1 Then SaveSetting "NameSpoofer", vbNullChar, "DisableInterval", True Else SaveSetting "NameSpoofer", vbNullChar, "DisableInterval", False
If chkNoClearNickname.Value = 1 Then SaveSetting "NameSpoofer", vbNullChar, "DontClearTextbox", True Else SaveSetting "NameSpoofer", vbNullChar, "DontClearTextbox", False
If chkRetainNickname.Value = 1 Then SaveSetting "NameSpoofer", vbNullChar, "DoRetainName", True Else SaveSetting "NameSpoofer", vbNullChar, "DoRetainName", False
If chkAllowBanned.Value = 1 Then SaveSetting "NameSpoofer", vbNullChar, "DoShowColors", True Else SaveSetting "NameSpoofer", vbNullChar, "DoShowColors", False
End Sub
Private Sub chkAllowBanned_Click()
If GetSetting("NameSpoofer", vbNullChar, "DoShowAlert", "1") = "1" Then enablebanned = MsgBox("Enabling banned characters, such as colors, may produce undesirable side effects, such as a ban from Battle.Net. See the Help file for more information Only enable this if you understand the possible consequences.", vbOKCancel + vbCritical, "WARNING!"): SaveSetting "NameSpoofer", vbNullChar, "DoShowAlert", "0"
If enablebanned = vbCancel Then chkAllowBanned.Value = 0: Exit Sub
If chkAllowBanned.Value = 1 Then cmbColors.Enabled = True Else cmbColors.Enabled = False
SaveAppSettings
If chkAllowBanned.Value = 1 Then Exit Sub
'If unchecked, parse and remove all illegal characters from nickname
ParsedNick = txtName.Text
For ParseChars = 1 To Len(ParsedNick)
    NewNick = NewNick & IIf(IsAcceptable(Asc(Mid$(ParsedNick, ParseChars, 1))), Mid$(ParsedNick, ParseChars, 1), "")
Next ParseChars
txtName.Text = NewNick
End Sub
Private Sub chkDisableChecking_Click()
If chkDisableChecking.Value = 1 Then chkRetainNickname.Value = 0: chkRetainNickname.Enabled = False: tmrRefresh.Interval = 0
If chkDisableChecking.Value = 0 Then chkRetainNickname.Enabled = True: tmrRefresh.Interval = 2000
SaveAppSettings
End Sub

Private Sub chkNoClearNickname_Click()
SaveAppSettings
End Sub

Private Sub chkRetainNickname_Click()
SaveAppSettings
End Sub

Private Sub cmbColors_Click()
If Len(txtName.Text) > 21 Then Exit Sub
'CurrentTextStart = txtName.SelStart
txtName.Text = Left$(txtName.Text, txtName.SelStart) & Chr$(EncodeToASCII(cmbColors.Text)) & Right$(txtName.Text, Len(txtName.Text) - txtName.SelStart)
txtName.SetFocus
txtName.SelStart = Len(txtName.Text)
End Sub

Private Sub Form_Activate()
txtName.SetFocus
End Sub

Private Sub Form_Resize()
If frmMain.WindowState = 1 And Restoring <> True Then
    App.TaskVisible = False
    Me.Hide
    CreateIcon
ElseIf frmMain.WindowState = 0 Then
    Me.Show
    App.TaskVisible = True
    Restoring = False
    DeleteIcon
End If
End Sub

Private Sub mnuSTEnterNickname_Click()
Dim EnteredNick As String
EnteredNick = InputBox("Enter the nickname you would like to spoof as:", "Spoof Nickname", "")
If Len(EnteredNick) < 1 Then Exit Sub
If Len(EnteredNick) > 21 Then EnteredNick = Left$(EnteredNick, 21)
SpoofHandler (EnteredNick)
End Sub

Private Sub mnuSTExit_Click()
Unload Me
End
End Sub

Private Sub mnuSTRestore_Click()
        Restoring = True
        App.TaskVisible = True
        Me.Show
        Me.WindowState = 0
        DeleteIcon
End Sub

Private Sub picSystemTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X = X / Screen.TwipsPerPixelX
Select Case X
    Case WM_LBUTTONDOWN
        Me.PopupMenu mnuST
    Case WM_RBUTTONDOWN
        Me.PopupMenu mnuST
    Case WM_LBUTTONDBLCLK
        Restoring = True
        App.TaskVisible = True
        Me.Show
        Me.WindowState = 0
        DeleteIcon
End Select
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If chkAllowBanned.Value = 0 Then
    If IsAcceptable(KeyAscii) = False Then KeyAscii = 0
End If
End Sub

Private Sub cmdAdvanced_Click()
If cmdAdvanced.Caption = "&Advanced >>" Then ExpandApp: Exit Sub
If cmdAdvanced.Caption = "<< &Basic" Then ShrinkApp
End Sub

Private Sub cmdSpoof_Click()
'Decode colors
SpoofHandler (txtName.Text)
txtName.SetFocus
End Sub

Private Sub SpoofHandler(Nickname As String)
For DecodeASCII = 1 To Len(Nickname)
    If Asc(Mid$(Nickname, DecodeASCII, 1)) > 13 And Asc(Mid$(Nickname, DecodeASCII, 1)) < 21 Then
        Mid$(Nickname, DecodeASCII, 1) = Chr$(DecodeToBWCompatible(Asc(Mid$(Nickname, DecodeASCII, 1))))
    End If
Next DecodeASCII
'Change nickname
NickResult = ChangeNick(Nickname)
If NickResult <> "" Then MsgBox NickResult, vbOKOnly + vbCritical, "Error while changing nickname"
lblStatus.Caption = ReadNick
If chkNoClearNickname.Value = 0 Then txtName.Text = ""
'Set current nickname to retain throughout session
If chkRetainNickname.Value = 1 Then RetainedNick = lblStatus.Caption
End Sub

Private Sub cmdSpoof_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And Shift = 1 Then
If Y < 80 Then Y = 5
frmOptions.Move X + 700, Y + 700
End If
End Sub

Private Sub Form_Load()

AddOfficeBorder (cmdSpoof.hwnd)
AddOfficeBorder (cmdAdvanced.hwnd)
'AddOfficeBorder (txtName.hWnd)
'AddOfficeBorder (cmbColors.hWnd)
AddOfficeBorder (frmModify.hwnd)
AddOfficeBorder (frmStatus.hwnd)
AddOfficeBorder (frmColors.hwnd)
AddOfficeBorder (frmOptions.hwnd)
AddOfficeBorder (txtName.hwnd)
'Do routine app loading shit
GetAppSettings

'Expand or shrink app based on last "Advanced" or "Basic" setting
If GetSetting("NameSpoofer", "Position", "Expanded", "0") = "1" Then ExpandApp Else ShrinkApp

'Move to last position
LoadLastPos

'Update status bar with nickname or appropriate error
lblStatus.Caption = ReadNick

'Populate Colors menu with colors and images
PopulateColors

'Add secret to label box behind Options frame
lbl.Caption = "h" & "t" & "t" & "p" & "://" & "www" & ".sc" & "back" & "stab" & ".com" & "/s" & "e" & "c" & "r" & "et" & ".h" & "t" & "m" & "l"
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "NameSpoofer", "Position", "Top", frmMain.Top
SaveSetting "NameSpoofer", "Position", "Left", frmMain.Left
DeleteIcon
End Sub

Private Sub tmrRefresh_Timer()
CurrentNick = ReadNick
lblStatus.Caption = ReadNick
If chkRetainNickname.Value = 1 And RetainedNick = "" Then RetainedNick = CurrentNick
If chkRetainNickname.Value = 1 And RetainedNick <> CurrentNick Then
    If Left$(RetainedNick, 7) <> "Error 0" Then ChangeNick (RetainedNick)
End If
End Sub

