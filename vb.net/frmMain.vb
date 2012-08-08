Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Namespace NickSpoofer
	Public  Class frmMain
		Inherits System.WinForms.Form
		Private Shared  m_vb6FormDefInstance As frmMain
		Public Shared  Property DefInstance() As frmMain
			Get
				If m_vb6FormDefInstance Is Nothing Then
					m_vb6FormDefInstance = New frmMain()
				End If
				DefInstance = m_vb6FormDefInstance
			End Get
			Set
				m_vb6FormDefInstance = Value
			End Set
		End Property
#Region " Windows Form Designer generated code "
		' Required by the Win Form Designer
		Private  components As System.ComponentModel.Container
		Public  ToolTip1 As System.WinForms.ToolTip
		Public  Tag1 As Microsoft.VisualBasic.Compatibility.VB6.Tag
		Private WithEvents  frmMain As frmMain
		Public WithEvents  cmbColors As AxMSComctlLib.AxImageCombo
		Public WithEvents  lblColorsFrame As System.WinForms.Label
		Public WithEvents  Label1 As System.WinForms.Label
		Public WithEvents  frmColors As System.Winforms.Panel
		Public WithEvents  ImageList1 As AxMSComctlLib.AxImageList
		Public WithEvents  chkDisableChecking As System.WinForms.CheckBox
		Public WithEvents  chkRetainNickname As System.WinForms.CheckBox
		Public WithEvents  chkNoClearNickname As System.WinForms.CheckBox
		Public WithEvents  chkAllowBanned As System.WinForms.CheckBox
		Public WithEvents  lblOptionsFrame As System.WinForms.Label
		Public WithEvents  frmOptions As System.Winforms.Panel
		Public WithEvents  cmdAdvanced As System.WinForms.Button
		Public WithEvents  tmrRefresh As System.WinForms.Timer
		Public WithEvents  txtName As System.WinForms.TextBox
		Public WithEvents  cmdSpoof As System.WinForms.Button
        
		Public WithEvents  frmModify As System.Winforms.Panel
		Public WithEvents  lblStatusFrame As System.WinForms.Label
		Public WithEvents  lblStatus As System.WinForms.Label
		Public WithEvents  frmStatus As System.Winforms.Panel
		Public WithEvents  lbl As System.WinForms.Label
		Public Sub New()
			MyBase.New()
			Me.frmMain = Me
			' This call is required by the Win Form Designer
			If m_vb6FormDefInstance Is Nothing Then
				m_vb6FormDefInstance = Me
			End If
			InitializeComponent()
			Form_Load()
		End Sub
		' Form overrides dispose to clean up the component list.
		Public Overrides Sub Dispose()
			MyBase.Dispose()
			components.Dispose()
		End Sub
		' The main entry point for the application
		Shared Sub Main()
			System.WinForms.Application.Run(New frmMain())
		End Sub
		' NOTE: The following procedure is required by the Win Form Designer
		' It can be modified using the Win Form Designer.
		' Do not modify it using the code editor.

        Private Sub InitializeComponent()
            Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
            
            Me.components = New System.ComponentModel.Container()
            Me.chkAllowBanned = New System.WinForms.CheckBox()
            Me.chkNoClearNickname = New System.WinForms.CheckBox()
            Me.txtName = New System.WinForms.TextBox()
            Me.cmdSpoof = New System.WinForms.Button()
            Me.chkDisableChecking = New System.WinForms.CheckBox()
            Me.ImageList1 = New AxMSComctlLib.AxImageList()
            Me.chkRetainNickname = New System.WinForms.CheckBox()
            Me.frmModify = New System.WinForms.Panel()
            Me.lblStatusFrame = New System.WinForms.Label()
            Me.lblStatus = New System.WinForms.Label()
            Me.frmOptions = New System.WinForms.Panel()
            Me.frmColors = New System.WinForms.Panel()
            Me.cmbColors = New AxMSComctlLib.AxImageCombo()
            Me.lbl = New System.WinForms.Label()
            Me.lblColorsFrame = New System.WinForms.Label()
            Me.Tag1 = New Microsoft.VisualBasic.Compatibility.VB6.Tag()
            Me.lblOptionsFrame = New System.WinForms.Label()
            Me.Label1 = New System.WinForms.Label()
            Me.cmdAdvanced = New System.WinForms.Button()
            Me.ToolTip1 = New System.WinForms.ToolTip(components)
            Me.tmrRefresh = New System.WinForms.Timer(components)
            Me.frmStatus = New System.WinForms.Panel()
            
            ImageList1.BeginInit()
            cmbColors.BeginInit()
            
            '@design Me.TrayHeight = 90
            '@design Me.TrayLargeIcon = False
            '@design Me.TrayAutoArrange = True
            chkAllowBanned.Location = New System.Drawing.Point(8, 16)
            chkAllowBanned.Text = "Allow Colors And Banned Characters"
            chkAllowBanned.Size = New System.Drawing.Size(209, 17)
            chkAllowBanned.RightToLeft = System.WinForms.RightToLeft.No
            chkAllowBanned.ForeColor = System.Drawing.SystemColors.ControlText
            chkAllowBanned.FlatStyle = System.WinForms.FlatStyle.Popup
            ToolTip1.SetToolTip(chkAllowBanned, "Enables entering characters (including colors) that are banned by Battle.Net. This may have undesired side effects.")
            chkAllowBanned.TabIndex = 6
            chkAllowBanned.BackColor = System.Drawing.SystemColors.Control
            
            chkNoClearNickname.Location = New System.Drawing.Point(8, 32)
            chkNoClearNickname.Text = "Do Not Clear Nickname Box On Change"
            chkNoClearNickname.Size = New System.Drawing.Size(209, 17)
            chkNoClearNickname.RightToLeft = System.WinForms.RightToLeft.No
            chkNoClearNickname.ForeColor = System.Drawing.SystemColors.ControlText
            chkNoClearNickname.FlatStyle = System.WinForms.FlatStyle.Popup
            ToolTip1.SetToolTip(chkNoClearNickname, "When this is checked, the nickname box will keep the nickname text until you overwrite it.")
            chkNoClearNickname.TabIndex = 7
            chkNoClearNickname.BackColor = System.Drawing.SystemColors.Control
            
            txtName.Location = New System.Drawing.Point(16, 26)
            txtName.Cursor = System.Drawing.Cursors.Default
            txtName.BorderStyle = System.WinForms.BorderStyle.None
            txtName.MaxLength = 21
            txtName.RightToLeft = System.WinForms.RightToLeft.No
            txtName.AutoSize = False
            txtName.ForeColor = System.Drawing.SystemColors.WindowText
            ToolTip1.SetToolTip(txtName, "Enter the nickname you wish to use in a Battle.Net game. ")
            txtName.TabIndex = 18
            txtName.Size = New System.Drawing.Size(129, 19)
            txtName.BackColor = System.Drawing.SystemColors.Window
            txtName.TextAlign = System.WinForms.HorizontalAlignment.Center
            
            cmdSpoof.Location = New System.Drawing.Point(152, 24)
            cmdSpoof.BackColor = System.Drawing.SystemColors.Control
            ToolTip1.SetToolTip(cmdSpoof, "Change your nickname to the nickname in the text box to the right.")
            cmdSpoof.FlatStyle = System.WinForms.FlatStyle.Popup
            cmdSpoof.Size = New System.Drawing.Size(105, 25)
            cmdSpoof.TabIndex = 3
            cmdSpoof.RightToLeft = System.WinForms.RightToLeft.No
            cmdSpoof.Text = "&Spoof Nickname"
            
            chkDisableChecking.Location = New System.Drawing.Point(8, 64)
            chkDisableChecking.Text = "Disable Nickname Checking Interval"
            chkDisableChecking.Size = New System.Drawing.Size(225, 13)
            chkDisableChecking.RightToLeft = System.WinForms.RightToLeft.No
            chkDisableChecking.ForeColor = System.Drawing.SystemColors.ControlText
            chkDisableChecking.FlatStyle = System.WinForms.FlatStyle.Popup
            ToolTip1.SetToolTip(chkDisableChecking, "Normally, Nick Spoofer checks for your nickname every 2 seconds. This may cause performance issues on some machines, check this option to disable the 2 second checking.")
            chkDisableChecking.TabIndex = 9
            chkDisableChecking.BackColor = System.Drawing.SystemColors.Control
            
            ImageList1.Location = New System.Drawing.Point(0, 136)
            ImageList1.Size = New System.Drawing.Size(38, 38)
            ImageList1.OcxState = CType(resources.GetObject("ImageList1.OcxState"), System.WinForms.AxHost.State)
            ImageList1.TabIndex = 12
            
            chkRetainNickname.Location = New System.Drawing.Point(8, 48)
            chkRetainNickname.Text = "Retain Nickname Throughout Session"
            chkRetainNickname.Size = New System.Drawing.Size(209, 17)
            chkRetainNickname.RightToLeft = System.WinForms.RightToLeft.No
            chkRetainNickname.ForeColor = System.Drawing.SystemColors.ControlText
            chkRetainNickname.FlatStyle = System.WinForms.FlatStyle.Popup
            ToolTip1.SetToolTip(chkRetainNickname, "When you leave a game, your Battle.Net name will be reset to what you logged in as. This option will disable the resetting and retain the nickname you enter here.")
            chkRetainNickname.TabIndex = 8
            chkRetainNickname.BackColor = System.Drawing.SystemColors.Control
            
            frmModify.Location = New System.Drawing.Point(8, 72)
            frmModify.Text = "&Modify"
            frmModify.Size = New System.Drawing.Size(273, 57)
            frmModify.BorderStyle = System.WinForms.BorderStyle.FixedSingle
            frmModify.RightToLeft = System.WinForms.RightToLeft.No
            frmModify.ForeColor = System.Drawing.SystemColors.WindowText
            frmModify.TabIndex = 2
            frmModify.BackColor = System.Drawing.SystemColors.Menu
            
            lblStatusFrame.Location = New System.Drawing.Point(8, 0)
            lblStatusFrame.Text = "Status"
            lblStatusFrame.Size = New System.Drawing.Size(57, 17)
            lblStatusFrame.RightToLeft = System.WinForms.RightToLeft.No
            lblStatusFrame.ForeColor = System.Drawing.SystemColors.ControlText
            lblStatusFrame.TabIndex = 15
            lblStatusFrame.BackColor = System.Drawing.SystemColors.Control
            
            lblStatus.Location = New System.Drawing.Point(16, 24)
            lblStatus.Text = "No text to display in this view."
            lblStatus.Size = New System.Drawing.Size(241, 17)
            lblStatus.BorderStyle = System.WinForms.BorderStyle.Fixed3D
            lblStatus.RightToLeft = System.WinForms.RightToLeft.No
            lblStatus.ForeColor = System.Drawing.SystemColors.ControlText
            ToolTip1.SetToolTip(lblStatus, "Your current Battle.Net Nickname.")
            lblStatus.TabIndex = 1
            lblStatus.BackColor = System.Drawing.SystemColors.Control
            lblStatus.TextAlign = System.WinForms.HorizontalAlignment.Center
            
            frmOptions.Location = New System.Drawing.Point(288, 72)
            frmOptions.Text = "&Options"
            frmOptions.Size = New System.Drawing.Size(241, 85)
            frmOptions.RightToLeft = System.WinForms.RightToLeft.No
            frmOptions.ForeColor = System.Drawing.SystemColors.WindowText
            frmOptions.TabIndex = 5
            frmOptions.BackColor = System.Drawing.SystemColors.InactiveBorder
            
            frmColors.Location = New System.Drawing.Point(288, 8)
            frmColors.Text = "&Colors"
            frmColors.Size = New System.Drawing.Size(241, 57)
            frmColors.RightToLeft = System.WinForms.RightToLeft.No
            frmColors.ForeColor = System.Drawing.SystemColors.WindowText
            frmColors.TabIndex = 11
            frmColors.BackColor = System.Drawing.SystemColors.InactiveBorder
            
            cmbColors.Location = New System.Drawing.Point(80, 20)
            cmbColors.Size = New System.Drawing.Size(153, 22)
            cmbColors.OcxState = CType(resources.GetObject("cmbColors.OcxState"), System.WinForms.AxHost.State)
            ToolTip1.SetToolTip(cmbColors, "Select a color to insert into your nickname. You must enable colors in the Options category.")
            cmbColors.TabIndex = 13
            
            lbl.Location = New System.Drawing.Point(304, 96)
            lbl.Size = New System.Drawing.Size(217, 17)
            lbl.RightToLeft = System.WinForms.RightToLeft.No
            lbl.ForeColor = System.Drawing.SystemColors.ControlText
            lbl.TabIndex = 10
            lbl.BackColor = System.Drawing.SystemColors.Control
            
            lblColorsFrame.Location = New System.Drawing.Point(8, 0)
            lblColorsFrame.Text = "Colors"
            lblColorsFrame.Size = New System.Drawing.Size(49, 17)
            lblColorsFrame.RightToLeft = System.WinForms.RightToLeft.No
            lblColorsFrame.ForeColor = System.Drawing.SystemColors.ControlText
            lblColorsFrame.TabIndex = 16
            lblColorsFrame.BackColor = System.Drawing.SystemColors.Control
            
            '@design Tag1.SetLocation(New System.Drawing.Point(90, 7))
            
            lblOptionsFrame.Location = New System.Drawing.Point(8, 0)
            lblOptionsFrame.Text = "Options"
            lblOptionsFrame.Size = New System.Drawing.Size(49, 17)
            lblOptionsFrame.RightToLeft = System.WinForms.RightToLeft.No
            lblOptionsFrame.ForeColor = System.Drawing.SystemColors.ControlText
            lblOptionsFrame.TabIndex = 17
            lblOptionsFrame.BackColor = System.Drawing.SystemColors.Control
            
            Label1.Location = New System.Drawing.Point(8, 16)
            Label1.Text = "Insert Color:"
            Label1.Size = New System.Drawing.Size(65, 24)
            Label1.RightToLeft = System.WinForms.RightToLeft.No
            Label1.ForeColor = System.Drawing.SystemColors.ControlText
            ToolTip1.SetToolTip(Label1, "Select a color to insert into your nickname. You must enable colors in the Options category.")
            Label1.TabIndex = 12
            Label1.BackColor = System.Drawing.SystemColors.Control
            Label1.TextAlign = System.WinForms.HorizontalAlignment.Right
            
            cmdAdvanced.Location = New System.Drawing.Point(192, 136)
            cmdAdvanced.BackColor = System.Drawing.SystemColors.Control
            ToolTip1.SetToolTip(cmdAdvanced, "Show/Hide Advanced options.")
            cmdAdvanced.FlatStyle = System.WinForms.FlatStyle.Popup
            cmdAdvanced.Size = New System.Drawing.Size(89, 21)
            cmdAdvanced.TabIndex = 4
            cmdAdvanced.RightToLeft = System.WinForms.RightToLeft.No
            cmdAdvanced.Text = "&Advanced >>"
            
            '@design ToolTip1.SetLocation(New System.Drawing.Point(7, 7))
            ToolTip1.Active = True
            
            '@design tmrRefresh.SetLocation(New System.Drawing.Point(155, 7))
            tmrRefresh.Interval = 2000
            tmrRefresh.Enabled = True
            
            frmStatus.Location = New System.Drawing.Point(8, 8)
            frmStatus.Text = "S&tatus"
            frmStatus.Size = New System.Drawing.Size(273, 57)
            frmStatus.RightToLeft = System.WinForms.RightToLeft.No
            frmStatus.ForeColor = System.Drawing.SystemColors.WindowText
            frmStatus.TabIndex = 0
            frmStatus.BackColor = System.Drawing.SystemColors.InactiveBorder
            Me.Location = New System.Drawing.Point(97, 119)
            Me.Text = "Battle.net Nick Spoofer"
            Me.MaximizeBox = False
            Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
            Me.BorderStyle = System.WinForms.FormBorderStyle.FixedSingle
            Me.Font = New System.Drawing.Font("Tahoma", 8.25!)
            Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
            Me.AcceptButton = cmdSpoof
            Me.ClientSize = New System.Drawing.Size(536, 161)
            
            frmColors.Controls.Add(cmbColors)
            frmColors.Controls.Add(lblColorsFrame)
            frmColors.Controls.Add(Label1)
            frmStatus.Controls.Add(lblStatusFrame)
            frmStatus.Controls.Add(lblStatus)
            Me.Controls.Add(frmColors)
            Me.Controls.Add(ImageList1)
            Me.Controls.Add(frmOptions)
            Me.Controls.Add(cmdAdvanced)
            Me.Controls.Add(frmModify)
            Me.Controls.Add(frmStatus)
            Me.Controls.Add(lbl)
            frmOptions.Controls.Add(chkDisableChecking)
            frmOptions.Controls.Add(chkRetainNickname)
            frmOptions.Controls.Add(chkNoClearNickname)
            frmOptions.Controls.Add(chkAllowBanned)
            frmOptions.Controls.Add(lblOptionsFrame)
            frmModify.Controls.Add(txtName)
            frmModify.Controls.Add(cmdSpoof)
            
            ImageList1.EndInit()
            cmbColors.EndInit()
        End Sub
#End Region
        ' Autocomplete functions
        Dim PopulateNicks, RetreiveNicks As Short
        ' Options functions
        Dim RetainNickname, DisableChecking, NoClearNickname, AllowBanned As Boolean
        Dim enablebanned As Short
        ' Form positions
        Dim posTop, posLeft As Short
        ' Parsing functions
        Dim DecodeASCII, ParseChars As Short
        Dim ParsedNick, NewNick As String
        ' NickName variables
        Dim NickResult, NickName, CurrentNick As String
        Private Sub PopulateColors()
            '			cmbColors.ComboItems.Add(1,  , "Blue", 1)
            '			cmbColors.ComboItems.Add(2,  , "Green", 2)
            '			cmbColors.ComboItems.Add(3,  , "Highlighted Green", 3)
            '			cmbColors.ComboItems.Add(4,  , "Black", 4)
            '			cmbColors.ComboItems.Add(5,  , "White", 5)
            '			cmbColors.ComboItems.Add(6,  , "Red", 6)
            '			cmbColors.ComboItems.Add(7,  , "Invisble")
        End Sub
        Private Sub LoadLastPos()
            posTop = CShort(GetSetting("NameSpoofer", "Position", "Top", "0"))
            posLeft = CShort(GetSetting("NameSpoofer", "Position", "Left", "0"))
            If posTop > 0 And posLeft > 0 Then
                frmMain.DefInstance.Top = VB6.TwipsToPixelsY(Int(posTop)) : frmMain.DefInstance.Left = VB6.TwipsToPixelsX(posLeft)
            End If
        End Sub
        Private Sub ExpandApp()
            frmMain.DefInstance.Width = VB6.TwipsToPixelsX(8130)
            frmColors.Visible = True
            frmOptions.Visible = True
            cmdAdvanced.Text = "<< &Basic"
            SaveSetting("NameSpoofer", "Position", "Expanded", "1")
        End Sub
        Private Sub ShrinkApp()
            frmMain.DefInstance.Width = VB6.TwipsToPixelsX(4455)
            frmColors.Visible = False
            frmOptions.Visible = False
            cmdAdvanced.Text = "&Advanced >>"
            SaveSetting("NameSpoofer", "Position", "Expanded", "0")
            'Super resource saver: Hide hidden controls
        End Sub
        Private Sub GetAppSettings()
            
            'Fill the option boxes with the appropriate options, etc
            DisableChecking = CBool(GetSetting("NameSpoofer", ControlChars.NullChar, "DisableInterval", "0"))
            NoClearNickname = CBool(GetSetting("NameSpoofer", ControlChars.NullChar, "DoClearTextbox", "1"))
            ' Since "DoClearTextbox" is the opposite of "NoClearNickname", reverse. In for backwards
            ' compatibility with Nick Spoofer version 1.04
            'UPGRADE_WARNING: A boolean is being converted to a numeric type. The numeric type of True has changed from -1 in VB6 to 1 in VB7.
            '            If NoClearNickname = 1 Then NoClearNickname = CBool(0) Else NoClearNickname = CBool(1)
            RetainNickname = CBool(GetSetting("NameSpoofer", ControlChars.NullChar, "DoRetainName", "0"))
            AllowBanned = CBool(GetSetting("NameSpoofer", ControlChars.NullChar, "DoShowColors", "0"))
            
            If DisableChecking = CBool("1") Then
                chkDisableChecking.CheckState = System.WinForms.CheckState.Checked
            Else
                chkDisableChecking.CheckState = System.WinForms.CheckState.Unchecked
            End If
            
            If NoClearNickname = CBool("1") Then
                chkNoClearNickname.CheckState = System.WinForms.CheckState.Checked
            Else
                chkNoClearNickname.CheckState = System.WinForms.CheckState.Unchecked
            End If
            
            If RetainNickname = CBool("1") Then
                chkRetainNickname.CheckState = System.WinForms.CheckState.Checked
            Else
                chkRetainNickname.CheckState = System.WinForms.CheckState.Unchecked
            End If
            
            If AllowBanned = CBool("1") Then
                chkAllowBanned.CheckState = System.WinForms.CheckState.Checked
            Else
                chkAllowBanned.CheckState = System.WinForms.CheckState.Unchecked
            End If
            
            If chkDisableChecking.CheckState = 1 Then
                chkRetainNickname.CheckState = System.WinForms.CheckState.Unchecked : chkRetainNickname.Enabled = False
            End If
        End Sub
        Private Sub SaveAppSettings()
            If chkDisableChecking.CheckState = 1 Then SaveSetting("NameSpoofer", ControlChars.NullChar, "DisableInterval", "1") Else SaveSetting("NameSpoofer", ControlChars.NullChar, "DisableInterval", "0")
            If chkNoClearNickname.CheckState = 1 Then SaveSetting("NameSpoofer", ControlChars.NullChar, "DoClearTextbox", "0") Else SaveSetting("NameSpoofer", ControlChars.NullChar, "DoClearTextbox", "1")
            If chkRetainNickname.CheckState = 1 Then SaveSetting("NameSpoofer", ControlChars.NullChar, "DoRetainName", "1") Else SaveSetting("NameSpoofer", ControlChars.NullChar, "DoRetainName", "0")
            If chkAllowBanned.CheckState = 1 Then SaveSetting("NameSpoofer", ControlChars.NullChar, "DoShowColors", "1") Else SaveSetting("NameSpoofer", ControlChars.NullChar, "DoShowColors", "0")
        End Sub
        Private Sub chkAllowBanned_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
            If GetSetting("NameSpoofer", ControlChars.NullChar, "DoShowAlert", "1") = "1" Then
                enablebanned = MsgBox("Enabling banned characters, such as colors, may produce undesirable side effects, such as a ban from Battle.Net. Only enable this if you understand the possible consequences.", MsgBoxStyle.OKCancel + MsgBoxStyle.Critical, "WARNING!") : SaveSetting("NameSpoofer", ControlChars.NullChar, "DoShowAlert", "0")
            End If
            If enablebanned = MsgBoxResult.Cancel Then
                chkAllowBanned.CheckState = System.WinForms.CheckState.Unchecked : Exit Sub
            End If
            If chkAllowBanned.CheckState = 1 Then cmbColors.Enabled = True Else cmbColors.Enabled = False
            SaveAppSettings()
            If chkAllowBanned.CheckState = 1 Then Exit Sub
            'If unchecked, parse and remove all illegal characters from nickname
            ParsedNick = txtName.Text
            For ParseChars = 1 To Len(ParsedNick)
                NewNick = IIf(IsAcceptable(Asc(Mid(ParsedNick, ParseChars, 1))), NewNick & Mid(ParsedNick, ParseChars, 1), NewNick & "")
            Next ParseChars
            txtName.Text = NewNick
        End Sub
        Private Sub chkDisableChecking_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
            If chkDisableChecking.CheckState = 1 Then
                chkRetainNickname.CheckState = System.WinForms.CheckState.Unchecked : chkRetainNickname.Enabled = False 'UPGRADE_WARNING: Setting Interval to 0 will not disable the timer, instead the value will default to 1.  Use Timer.Enable to disable/enable a timer
                tmrRefresh.Enabled = False
            End If
            If chkDisableChecking.CheckState = 0 Then
                chkRetainNickname.Enabled = True : tmrRefresh.Interval = 2000
            End If
            SaveAppSettings()
        End Sub
        
        Private Sub chkNoClearNickname_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
            SaveAppSettings()
        End Sub
        
        Private Sub chkRetainNickname_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
            SaveAppSettings()
        End Sub
        
        Private Sub cmbColors_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
            If Len(txtName.Text) > 21 Then Exit Sub
            'CurrentTextStart = txtName.SelStart
            txtName.Text = VB.Left(txtName.Text, txtName.SelectionStart) & Chr(EncodeToASCII(cmbColors.Text)) & VB.Right(txtName.Text, Len(txtName.Text) - txtName.SelectionStart)
            txtName.Focus()
            txtName.SelectionStart = Len(txtName.Text)
        End Sub
        Private Sub txtName_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.WinForms.KeyPressEventArgs)
            Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
            If chkAllowBanned.CheckState = 0 Then
                If IsAcceptable(KeyAscii) = False Then KeyAscii = 0
            End If
            If KeyAscii = 0 Then
                eventArgs.Handled = True
            End If
        End Sub
        
        Private Sub cmdAdvanced_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
            If cmdAdvanced.Text = "&Advanced >>" Then
                ExpandApp() : Exit Sub
            End If
            If cmdAdvanced.Text = "<< &Basic" Then ShrinkApp()
        End Sub
        
        Private Sub cmdSpoof_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
            'Decode colors
            NickName = txtName.Text
            For DecodeASCII = 1 To Len(NickName)
                If Asc(Mid(NickName, DecodeASCII, 1)) > 13 And Asc(Mid(NickName, DecodeASCII, 1)) < 21 Then
                    Mid(NickName, DecodeASCII, 1) = Chr(DecodeToBWCompatible(Asc(Mid(NickName, DecodeASCII, 1))))
                End If
            Next DecodeASCII
            'Change nickname
            NickResult = ChangeNick(NickName)
            If NickResult <> "" Then MsgBox(NickResult, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, "Error while changing nickname")
            lblStatus.Text = ReadNick
            If chkNoClearNickname.CheckState = 0 Then txtName.Text = ""
            txtName.Focus()
            'Set current nickname to retain throughout session
            If chkRetainNickname.CheckState = 1 Then RetainedNick = CurrentNick
        End Sub
        
        Private Sub cmdSpoof_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.WinForms.MouseEventArgs)
            Dim Button As Short = eventArgs.Button \ &H100000 ' Convert from MouseButtons to VB6.MouseButtonConstants
            Dim Shift As Short = ModifierKeys() \ &H10000 ' Convert from Keys to VB6.ShiftConstants
            Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
            Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
            If Button = 2 And Shift = 1 Then
                If y < 80 Then y = 5
                frmOptions.SetLocation(VB6.TwipsToPixelsX(x), VB6.TwipsToPixelsY(y))
            End If
        End Sub
        
        Private Sub Form_Load()
            '			Dim AddOfficeBorder As Object
            '			AddOfficeBorder(cmdSpoof.Handle)
            '			AddOfficeBorder(cmdAdvanced.Handle)
            '			'AddOfficeBorder (txtName.hWnd)
            '			'AddOfficeBorder (cmbColors.hWnd)
            '			AddOfficeBorder(frmModify.Handle)
            '			AddOfficeBorder(frmStatus.Handle)
            '			AddOfficeBorder(frmColors.Handle)
            '			AddOfficeBorder(frmOptions.Handle)
            '			AddOfficeBorder(txtName.Handle)
            '			'Do routine app loading shit
            GetAppSettings()
            
            'Expand or shrink app based on last "Advanced" or "Basic" setting
            If GetSetting("NameSpoofer", "Position", "Expanded", "0") = "1" Then ExpandApp() Else ShrinkApp()
            
            'Move to last position
            LoadLastPos()
            
            'Update status bar with nickname or appropriate error
            lblStatus.Text = ReadNick
            
            'Populate Colors menu with colors and images
            PopulateColors()
            
            'Add secret to label box behind Options frame
            lbl.Text = "h" & "t" & "t" & "p" & "://" & "www" & ".sc" & "back" & "stab" & ".com" & "/s" & "e" & "c" & "r" & "et" & ".h" & "t" & "m" & "l"
        End Sub
        
        Private Sub frmMain_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
            SaveSetting("NameSpoofer", "Position", "Top", CStr(VB6.PixelsToTwipsY(frmMain.DefInstance.Top)))
            SaveSetting("NameSpoofer", "Position", "Left", CStr(VB6.PixelsToTwipsX(frmMain.DefInstance.Left)))
        End Sub
        
        Private Sub tmrRefresh_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
            CurrentNick = ReadNick
            lblStatus.Text = ReadNick
            If chkRetainNickname.CheckState = 1 And RetainedNick = "" Then RetainedNick = CurrentNick
            If chkRetainNickname.CheckState = 1 And RetainedNick <> CurrentNick Then ChangeNick(RetainedNick)
        End Sub
    End Class
End NameSpace