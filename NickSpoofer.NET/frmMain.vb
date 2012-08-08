Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMain
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.Container
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents picSystemTray As System.Windows.Forms.PictureBox
	Public WithEvents cmbColors As AxMSComctlLib.AxImageCombo
	Public WithEvents lblColorsFrame As System.Windows.Forms.Label
	Public WithEvents lblInsertColor As System.Windows.Forms.Label
	Public WithEvents frmColors As System.Windows.Forms.Panel
	Public WithEvents ImageList1 As AxMSComctlLib.AxImageList
	Public WithEvents chkDisableChecking As System.Windows.Forms.CheckBox
	Public WithEvents chkRetainNickname As System.Windows.Forms.CheckBox
	Public WithEvents chkNoClearNickname As System.Windows.Forms.CheckBox
	Public WithEvents chkAllowBanned As System.Windows.Forms.CheckBox
	Public WithEvents lblOptionsFrame As System.Windows.Forms.Label
	Public WithEvents frmOptions As System.Windows.Forms.Panel
	Public WithEvents cmdAdvanced As System.Windows.Forms.Button
	Public WithEvents tmrRefresh As System.Windows.Forms.Timer
	Public WithEvents txtName As System.Windows.Forms.TextBox
	Public WithEvents cmdSpoof As System.Windows.Forms.Button
	Public WithEvents lblModifyFrame As System.Windows.Forms.Label
	Public WithEvents frmModify As System.Windows.Forms.Panel
	Public WithEvents lblStatusFrame As System.Windows.Forms.Label
	Public WithEvents lblStatus As System.Windows.Forms.Label
	Public WithEvents frmStatus As System.Windows.Forms.Panel
	Public WithEvents lbl As System.Windows.Forms.Label
	Public WithEvents mnuSTRestore As System.Windows.Forms.MenuItem
	Public WithEvents mnuSTEnterNickname As System.Windows.Forms.MenuItem
	Public WithEvents mnuSTExit As System.Windows.Forms.MenuItem
	Public WithEvents mnuST As System.Windows.Forms.MenuItem
	Public MainMenu1 As System.Windows.Forms.MainMenu
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.picSystemTray = New System.Windows.Forms.PictureBox
		Me.frmColors = New System.Windows.Forms.Panel
		Me.cmbColors = New AxMSComctlLib.AxImageCombo
		Me.lblColorsFrame = New System.Windows.Forms.Label
		Me.lblInsertColor = New System.Windows.Forms.Label
		Me.ImageList1 = New AxMSComctlLib.AxImageList
		Me.frmOptions = New System.Windows.Forms.Panel
		Me.chkDisableChecking = New System.Windows.Forms.CheckBox
		Me.chkRetainNickname = New System.Windows.Forms.CheckBox
		Me.chkNoClearNickname = New System.Windows.Forms.CheckBox
		Me.chkAllowBanned = New System.Windows.Forms.CheckBox
		Me.lblOptionsFrame = New System.Windows.Forms.Label
		Me.cmdAdvanced = New System.Windows.Forms.Button
		Me.tmrRefresh = New System.Windows.Forms.Timer(components)
		Me.frmModify = New System.Windows.Forms.Panel
		Me.txtName = New System.Windows.Forms.TextBox
		Me.cmdSpoof = New System.Windows.Forms.Button
		Me.lblModifyFrame = New System.Windows.Forms.Label
		Me.frmStatus = New System.Windows.Forms.Panel
		Me.lblStatusFrame = New System.Windows.Forms.Label
		Me.lblStatus = New System.Windows.Forms.Label
		Me.lbl = New System.Windows.Forms.Label
		Me.MainMenu1 = New System.Windows.Forms.MainMenu
		Me.mnuST = New System.Windows.Forms.MenuItem
		Me.mnuSTRestore = New System.Windows.Forms.MenuItem
		Me.mnuSTEnterNickname = New System.Windows.Forms.MenuItem
		Me.mnuSTExit = New System.Windows.Forms.MenuItem
		CType(Me.cmbColors, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.ImageList1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Battle.net Nick Spoofer"
		Me.ClientSize = New System.Drawing.Size(536, 159)
		Me.Location = New System.Drawing.Point(97, 112)
		Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Icon = CType(resources.GetObject("frmMain.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmMain"
		Me.picSystemTray.Size = New System.Drawing.Size(33, 25)
		Me.picSystemTray.Location = New System.Drawing.Point(0, 136)
		Me.picSystemTray.Image = CType(resources.GetObject("picSystemTray.Image"), System.Drawing.Image)
		Me.picSystemTray.TabIndex = 19
		Me.picSystemTray.Visible = False
		Me.picSystemTray.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picSystemTray.Dock = System.Windows.Forms.DockStyle.None
		Me.picSystemTray.BackColor = System.Drawing.SystemColors.Control
		Me.picSystemTray.CausesValidation = True
		Me.picSystemTray.Enabled = True
		Me.picSystemTray.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picSystemTray.Cursor = System.Windows.Forms.Cursors.Default
		Me.picSystemTray.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picSystemTray.TabStop = True
		Me.picSystemTray.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.picSystemTray.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.picSystemTray.Name = "picSystemTray"
		Me.frmColors.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.frmColors.Text = "&Colors"
		Me.frmColors.ForeColor = System.Drawing.SystemColors.WindowText
		Me.frmColors.Size = New System.Drawing.Size(241, 57)
		Me.frmColors.Location = New System.Drawing.Point(288, 8)
		Me.frmColors.TabIndex = 11
		Me.frmColors.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmColors.BackColor = System.Drawing.SystemColors.Control
		Me.frmColors.Enabled = True
		Me.frmColors.Cursor = System.Windows.Forms.Cursors.Default
		Me.frmColors.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmColors.Visible = True
		Me.frmColors.Name = "frmColors"
		cmbColors.OcxState = CType(resources.GetObject("cmbColors.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmbColors.Size = New System.Drawing.Size(153, 22)
		Me.cmbColors.Location = New System.Drawing.Point(80, 20)
		Me.cmbColors.TabIndex = 13
		Me.cmbColors.Name = "cmbColors"
		Me.lblColorsFrame.Text = "Colors"
		Me.lblColorsFrame.Size = New System.Drawing.Size(49, 17)
		Me.lblColorsFrame.Location = New System.Drawing.Point(8, 0)
		Me.lblColorsFrame.TabIndex = 16
		Me.lblColorsFrame.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblColorsFrame.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblColorsFrame.BackColor = System.Drawing.SystemColors.Control
		Me.lblColorsFrame.Enabled = True
		Me.lblColorsFrame.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblColorsFrame.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblColorsFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblColorsFrame.UseMnemonic = True
		Me.lblColorsFrame.Visible = True
		Me.lblColorsFrame.AutoSize = False
		Me.lblColorsFrame.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblColorsFrame.Name = "lblColorsFrame"
		Me.lblInsertColor.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblInsertColor.Text = "Insert Color:"
		Me.lblInsertColor.Size = New System.Drawing.Size(65, 17)
		Me.lblInsertColor.Location = New System.Drawing.Point(8, 24)
		Me.lblInsertColor.TabIndex = 12
		Me.ToolTip1.SetToolTip(Me.lblInsertColor, "Select a color to insert into your nickname. You must enable colors in the Options category.")
		Me.lblInsertColor.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblInsertColor.BackColor = System.Drawing.SystemColors.Control
		Me.lblInsertColor.Enabled = True
		Me.lblInsertColor.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblInsertColor.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInsertColor.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInsertColor.UseMnemonic = True
		Me.lblInsertColor.Visible = True
		Me.lblInsertColor.AutoSize = False
		Me.lblInsertColor.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInsertColor.Name = "lblInsertColor"
		ImageList1.OcxState = CType(resources.GetObject("ImageList1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.ImageList1.Location = New System.Drawing.Point(0, 136)
		Me.ImageList1.Name = "ImageList1"
		Me.frmOptions.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.frmOptions.Text = "&Options"
		Me.frmOptions.ForeColor = System.Drawing.SystemColors.WindowText
		Me.frmOptions.Size = New System.Drawing.Size(241, 85)
		Me.frmOptions.Location = New System.Drawing.Point(288, 72)
		Me.frmOptions.TabIndex = 5
		Me.frmOptions.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmOptions.BackColor = System.Drawing.SystemColors.Control
		Me.frmOptions.Enabled = True
		Me.frmOptions.Cursor = System.Windows.Forms.Cursors.Default
		Me.frmOptions.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmOptions.Visible = True
		Me.frmOptions.Name = "frmOptions"
		Me.chkDisableChecking.Text = "Disable Nickname Checking Interval"
		Me.chkDisableChecking.Size = New System.Drawing.Size(225, 13)
		Me.chkDisableChecking.Location = New System.Drawing.Point(8, 64)
		Me.chkDisableChecking.TabIndex = 9
		Me.ToolTip1.SetToolTip(Me.chkDisableChecking, "Normally, Nick Spoofer checks for your nickname every 2 seconds. This may cause performance issues on some machines, check this option to disable the 2 second checking.")
		Me.chkDisableChecking.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkDisableChecking.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkDisableChecking.BackColor = System.Drawing.SystemColors.Control
		Me.chkDisableChecking.CausesValidation = True
		Me.chkDisableChecking.Enabled = True
		Me.chkDisableChecking.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkDisableChecking.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkDisableChecking.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkDisableChecking.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkDisableChecking.TabStop = True
		Me.chkDisableChecking.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkDisableChecking.Visible = True
		Me.chkDisableChecking.Name = "chkDisableChecking"
		Me.chkRetainNickname.Text = "Retain Nickname Throughout Session"
		Me.chkRetainNickname.Size = New System.Drawing.Size(209, 17)
		Me.chkRetainNickname.Location = New System.Drawing.Point(8, 48)
		Me.chkRetainNickname.TabIndex = 8
		Me.ToolTip1.SetToolTip(Me.chkRetainNickname, "When you leave a game, your Battle.Net name will be reset to what you logged in as. This option will disable the resetting and retain the nickname you enter here.")
		Me.chkRetainNickname.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkRetainNickname.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkRetainNickname.BackColor = System.Drawing.SystemColors.Control
		Me.chkRetainNickname.CausesValidation = True
		Me.chkRetainNickname.Enabled = True
		Me.chkRetainNickname.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkRetainNickname.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkRetainNickname.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkRetainNickname.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkRetainNickname.TabStop = True
		Me.chkRetainNickname.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkRetainNickname.Visible = True
		Me.chkRetainNickname.Name = "chkRetainNickname"
		Me.chkNoClearNickname.Text = "Do Not Clear Nickname Box On Change"
		Me.chkNoClearNickname.Size = New System.Drawing.Size(209, 17)
		Me.chkNoClearNickname.Location = New System.Drawing.Point(8, 32)
		Me.chkNoClearNickname.TabIndex = 7
		Me.ToolTip1.SetToolTip(Me.chkNoClearNickname, "When this is checked, the nickname box will keep the nickname text until you overwrite it.")
		Me.chkNoClearNickname.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkNoClearNickname.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkNoClearNickname.BackColor = System.Drawing.SystemColors.Control
		Me.chkNoClearNickname.CausesValidation = True
		Me.chkNoClearNickname.Enabled = True
		Me.chkNoClearNickname.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkNoClearNickname.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkNoClearNickname.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkNoClearNickname.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkNoClearNickname.TabStop = True
		Me.chkNoClearNickname.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkNoClearNickname.Visible = True
		Me.chkNoClearNickname.Name = "chkNoClearNickname"
		Me.chkAllowBanned.Text = "Allow Colors And Banned Characters"
		Me.chkAllowBanned.Size = New System.Drawing.Size(209, 17)
		Me.chkAllowBanned.Location = New System.Drawing.Point(8, 16)
		Me.chkAllowBanned.TabIndex = 6
		Me.ToolTip1.SetToolTip(Me.chkAllowBanned, "Enables entering characters (including colors) that are banned by Battle.Net. This may have undesired side effects.")
		Me.chkAllowBanned.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkAllowBanned.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkAllowBanned.BackColor = System.Drawing.SystemColors.Control
		Me.chkAllowBanned.CausesValidation = True
		Me.chkAllowBanned.Enabled = True
		Me.chkAllowBanned.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkAllowBanned.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkAllowBanned.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkAllowBanned.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkAllowBanned.TabStop = True
		Me.chkAllowBanned.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkAllowBanned.Visible = True
		Me.chkAllowBanned.Name = "chkAllowBanned"
		Me.lblOptionsFrame.Text = "Options"
		Me.lblOptionsFrame.Size = New System.Drawing.Size(49, 17)
		Me.lblOptionsFrame.Location = New System.Drawing.Point(8, 0)
		Me.lblOptionsFrame.TabIndex = 17
		Me.lblOptionsFrame.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblOptionsFrame.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblOptionsFrame.BackColor = System.Drawing.SystemColors.Control
		Me.lblOptionsFrame.Enabled = True
		Me.lblOptionsFrame.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblOptionsFrame.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblOptionsFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblOptionsFrame.UseMnemonic = True
		Me.lblOptionsFrame.Visible = True
		Me.lblOptionsFrame.AutoSize = False
		Me.lblOptionsFrame.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblOptionsFrame.Name = "lblOptionsFrame"
		Me.cmdAdvanced.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdAdvanced.Text = "&Advanced >>"
		Me.cmdAdvanced.Size = New System.Drawing.Size(89, 21)
		Me.cmdAdvanced.Location = New System.Drawing.Point(192, 136)
		Me.cmdAdvanced.TabIndex = 4
		Me.ToolTip1.SetToolTip(Me.cmdAdvanced, "Show/Hide Advanced options.")
		Me.cmdAdvanced.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAdvanced.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAdvanced.CausesValidation = True
		Me.cmdAdvanced.Enabled = True
		Me.cmdAdvanced.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAdvanced.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAdvanced.TabStop = True
		Me.cmdAdvanced.Name = "cmdAdvanced"
		Me.tmrRefresh.Interval = 2000
		Me.tmrRefresh.Enabled = True
		Me.frmModify.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.frmModify.Text = "&Modify"
		Me.frmModify.ForeColor = System.Drawing.SystemColors.WindowText
		Me.frmModify.Size = New System.Drawing.Size(273, 57)
		Me.frmModify.Location = New System.Drawing.Point(8, 72)
		Me.frmModify.TabIndex = 2
		Me.frmModify.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmModify.BackColor = System.Drawing.SystemColors.Control
		Me.frmModify.Enabled = True
		Me.frmModify.Cursor = System.Windows.Forms.Cursors.Default
		Me.frmModify.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmModify.Visible = True
		Me.frmModify.Name = "frmModify"
		Me.txtName.AutoSize = False
		Me.txtName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtName.Size = New System.Drawing.Size(129, 19)
		Me.txtName.Location = New System.Drawing.Point(16, 26)
		Me.txtName.Maxlength = 21
		Me.txtName.TabIndex = 18
		Me.ToolTip1.SetToolTip(Me.txtName, "Enter the nickname you wish to use in a Battle.Net game. ")
		Me.txtName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtName.AcceptsReturn = True
		Me.txtName.BackColor = System.Drawing.SystemColors.Window
		Me.txtName.CausesValidation = True
		Me.txtName.Enabled = True
		Me.txtName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtName.HideSelection = True
		Me.txtName.ReadOnly = False
		Me.txtName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtName.MultiLine = False
		Me.txtName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtName.TabStop = True
		Me.txtName.Visible = True
		Me.txtName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.txtName.Name = "txtName"
		Me.cmdSpoof.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSpoof.Text = "&Spoof Nickname"
		Me.AcceptButton = Me.cmdSpoof
		Me.cmdSpoof.Size = New System.Drawing.Size(105, 25)
		Me.cmdSpoof.Location = New System.Drawing.Point(152, 24)
		Me.cmdSpoof.TabIndex = 3
		Me.ToolTip1.SetToolTip(Me.cmdSpoof, "Change your nickname to the nickname in the text box to the left.")
		Me.cmdSpoof.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSpoof.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSpoof.CausesValidation = True
		Me.cmdSpoof.Enabled = True
		Me.cmdSpoof.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSpoof.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSpoof.TabStop = True
		Me.cmdSpoof.Name = "cmdSpoof"
		Me.lblModifyFrame.BackColor = System.Drawing.Color.Transparent
		Me.lblModifyFrame.Text = "Modify"
		Me.lblModifyFrame.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lblModifyFrame.Size = New System.Drawing.Size(36, 17)
		Me.lblModifyFrame.Location = New System.Drawing.Point(8, 0)
		Me.lblModifyFrame.TabIndex = 14
		Me.lblModifyFrame.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblModifyFrame.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblModifyFrame.Enabled = True
		Me.lblModifyFrame.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblModifyFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblModifyFrame.UseMnemonic = True
		Me.lblModifyFrame.Visible = True
		Me.lblModifyFrame.AutoSize = False
		Me.lblModifyFrame.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblModifyFrame.Name = "lblModifyFrame"
		Me.frmStatus.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.frmStatus.Text = "S&tatus"
		Me.frmStatus.ForeColor = System.Drawing.SystemColors.WindowText
		Me.frmStatus.Size = New System.Drawing.Size(273, 57)
		Me.frmStatus.Location = New System.Drawing.Point(8, 8)
		Me.frmStatus.TabIndex = 0
		Me.frmStatus.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmStatus.BackColor = System.Drawing.SystemColors.Control
		Me.frmStatus.Enabled = True
		Me.frmStatus.Cursor = System.Windows.Forms.Cursors.Default
		Me.frmStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmStatus.Visible = True
		Me.frmStatus.Name = "frmStatus"
		Me.lblStatusFrame.Text = "Status"
		Me.lblStatusFrame.Size = New System.Drawing.Size(57, 17)
		Me.lblStatusFrame.Location = New System.Drawing.Point(8, 0)
		Me.lblStatusFrame.TabIndex = 15
		Me.lblStatusFrame.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblStatusFrame.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblStatusFrame.BackColor = System.Drawing.SystemColors.Control
		Me.lblStatusFrame.Enabled = True
		Me.lblStatusFrame.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblStatusFrame.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblStatusFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblStatusFrame.UseMnemonic = True
		Me.lblStatusFrame.Visible = True
		Me.lblStatusFrame.AutoSize = False
		Me.lblStatusFrame.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblStatusFrame.Name = "lblStatusFrame"
		Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblStatus.Text = "No text to display in this view."
		Me.lblStatus.Size = New System.Drawing.Size(241, 17)
		Me.lblStatus.Location = New System.Drawing.Point(16, 24)
		Me.lblStatus.TabIndex = 1
		Me.ToolTip1.SetToolTip(Me.lblStatus, "Your current Battle.Net Nickname.")
		Me.lblStatus.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblStatus.BackColor = System.Drawing.SystemColors.Control
		Me.lblStatus.Enabled = True
		Me.lblStatus.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblStatus.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblStatus.UseMnemonic = True
		Me.lblStatus.Visible = True
		Me.lblStatus.AutoSize = False
		Me.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lblStatus.Name = "lblStatus"
		Me.lbl.Size = New System.Drawing.Size(217, 17)
		Me.lbl.Location = New System.Drawing.Point(304, 96)
		Me.lbl.TabIndex = 10
		Me.lbl.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lbl.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lbl.BackColor = System.Drawing.SystemColors.Control
		Me.lbl.Enabled = True
		Me.lbl.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lbl.Cursor = System.Windows.Forms.Cursors.Default
		Me.lbl.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lbl.UseMnemonic = True
		Me.lbl.Visible = True
		Me.lbl.AutoSize = False
		Me.lbl.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lbl.Name = "lbl"
		Me.mnuST.Text = "&SystemTray"
		Me.mnuST.Visible = False
		Me.mnuST.Checked = False
		Me.mnuST.Enabled = True
		Me.mnuST.MDIList = False
		Me.mnuSTRestore.Text = "&Restore"
		Me.mnuSTRestore.Checked = False
		Me.mnuSTRestore.Enabled = True
		Me.mnuSTRestore.Visible = True
		Me.mnuSTRestore.MDIList = False
		Me.mnuSTEnterNickname.Text = "E&nter Nickname..."
		Me.mnuSTEnterNickname.Checked = False
		Me.mnuSTEnterNickname.Enabled = True
		Me.mnuSTEnterNickname.Visible = True
		Me.mnuSTEnterNickname.MDIList = False
		Me.mnuSTExit.Text = "E&xit"
		Me.mnuSTExit.Checked = False
		Me.mnuSTExit.Enabled = True
		Me.mnuSTExit.Visible = True
		Me.mnuSTExit.MDIList = False
		Me.Controls.Add(picSystemTray)
		Me.Controls.Add(frmColors)
		Me.Controls.Add(ImageList1)
		Me.Controls.Add(frmOptions)
		Me.Controls.Add(cmdAdvanced)
		Me.Controls.Add(frmModify)
		Me.Controls.Add(frmStatus)
		Me.Controls.Add(lbl)
		Me.frmColors.Controls.Add(cmbColors)
		Me.frmColors.Controls.Add(lblColorsFrame)
		Me.frmColors.Controls.Add(lblInsertColor)
		Me.frmOptions.Controls.Add(chkDisableChecking)
		Me.frmOptions.Controls.Add(chkRetainNickname)
		Me.frmOptions.Controls.Add(chkNoClearNickname)
		Me.frmOptions.Controls.Add(chkAllowBanned)
		Me.frmOptions.Controls.Add(lblOptionsFrame)
		Me.frmModify.Controls.Add(txtName)
		Me.frmModify.Controls.Add(cmdSpoof)
		Me.frmModify.Controls.Add(lblModifyFrame)
		Me.frmStatus.Controls.Add(lblStatusFrame)
		Me.frmStatus.Controls.Add(lblStatus)
		CType(Me.cmbColors, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.ImageList1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.mnuST.Index = 0
		MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){mnuST})
		Me.mnuSTRestore.Index = 0
		Me.mnuSTEnterNickname.Index = 1
		Me.mnuSTExit.Index = 2
		mnuST.MenuItems.AddRange(New System.Windows.Forms.MenuItem(){mnuSTRestore, mnuSTEnterNickname, mnuSTExit})
		Me.Menu = MainMenu1
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmMain
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmMain
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmMain()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	' Autocomplete functions
	Dim PopulateNicks, RetreiveNicks As Short
	' Options functions
	Dim RetainNickname, DisableChecking, NoClearNickname, AllowBanned As Boolean
	Dim enablebanned As Short
	' Form positions
	Dim posTop, posLeft As Integer
	' Parsing functions
	Dim DecodeASCII, ParseChars As Short
	Dim ParsedNick, NewNick As String
	' NickName variables
	Dim NickResult, CurrentNick As String
	Dim Restoring As Boolean
	
	Public Sub CreateIcon()
		Dim Tic As NOTIFYICONDATA
		Tic.cbSize = Len(Tic)
		Tic.hwnd = picSystemTray.Handle.ToInt32
		Tic.uID = 1
		Tic.uFlags = NIF_DOALL
		Tic.uCallbackMessage = WM_MOUSEMOVE
		Tic.hIcon = CInt(CObj(picSystemTray.Image))
		Tic.szTip = "Nick Spoofer" & Chr(0)
		Shell_NotifyIcon(NIM_ADD, Tic)
	End Sub
	
	Public Sub DeleteIcon()
		Dim Tic As NOTIFYICONDATA
		Tic.cbSize = Len(Tic)
		Tic.hwnd = picSystemTray.Handle.ToInt32
		Tic.uID = 1
		Shell_NotifyIcon(NIM_DELETE, Tic)
	End Sub
	
	Private Sub PopulateColors()
		cmbColors.ComboItems.Add(1,  , "Blue", 1)
		cmbColors.ComboItems.Add(2,  , "Green", 2)
		cmbColors.ComboItems.Add(3,  , "Highlighted Green", 3)
		cmbColors.ComboItems.Add(4,  , "Black", 4)
		cmbColors.ComboItems.Add(5,  , "White", 5)
		cmbColors.ComboItems.Add(6,  , "Red", 6)
		cmbColors.ComboItems.Add(7,  , "Invisble")
	End Sub
	Private Sub LoadLastPos()
		posTop = CInt(GetSetting("NameSpoofer", "Position", "Top", "0"))
		posLeft = CInt(GetSetting("NameSpoofer", "Position", "Left", "0"))
		If posTop > VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) Or posLeft > VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) Then Exit Sub
		If posTop > 0 And posLeft > 0 Then frmMain.DefInstance.Top = VB6.TwipsToPixelsY(Int(posTop)) : frmMain.DefInstance.Left = VB6.TwipsToPixelsX(posLeft)
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
		DisableChecking = CBool(GetSetting("NameSpoofer", ControlChars.NullChar, "DisableInterval", CStr(False)))
		NoClearNickname = CBool(GetSetting("NameSpoofer", ControlChars.NullChar, "DontClearTextbox", CStr(False)))
		' Since "DoClearTextbox" is the opposite of "NoClearNickname", reverse. In for backwards
		' compatibility with Nick Spoofer version 1.04
		'If NoClearNickname = 1 Then NoClearNickname = 0 Else NoClearNickname = 1
		RetainNickname = CBool(GetSetting("NameSpoofer", ControlChars.NullChar, "DoRetainName", CStr(False)))
		AllowBanned = CBool(GetSetting("NameSpoofer", ControlChars.NullChar, "DoShowColors", CStr(False)))
		
		If DisableChecking Then
			chkDisableChecking.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkDisableChecking.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		If NoClearNickname Then
			chkNoClearNickname.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkNoClearNickname.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		If RetainNickname Then
			chkRetainNickname.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkRetainNickname.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		If AllowBanned Then
			cmbColors.Enabled = True
			chkAllowBanned.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			cmbColors.Enabled = False
			chkAllowBanned.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		If chkDisableChecking.CheckState = 1 Then chkRetainNickname.CheckState = System.Windows.Forms.CheckState.Unchecked : chkRetainNickname.Enabled = False
	End Sub
	Private Sub SaveAppSettings()
		If chkDisableChecking.CheckState = 1 Then SaveSetting("NameSpoofer", ControlChars.NullChar, "DisableInterval", CStr(True)) Else SaveSetting("NameSpoofer", ControlChars.NullChar, "DisableInterval", CStr(False))
		If chkNoClearNickname.CheckState = 1 Then SaveSetting("NameSpoofer", ControlChars.NullChar, "DontClearTextbox", CStr(True)) Else SaveSetting("NameSpoofer", ControlChars.NullChar, "DontClearTextbox", CStr(False))
		If chkRetainNickname.CheckState = 1 Then SaveSetting("NameSpoofer", ControlChars.NullChar, "DoRetainName", CStr(True)) Else SaveSetting("NameSpoofer", ControlChars.NullChar, "DoRetainName", CStr(False))
		If chkAllowBanned.CheckState = 1 Then SaveSetting("NameSpoofer", ControlChars.NullChar, "DoShowColors", CStr(True)) Else SaveSetting("NameSpoofer", ControlChars.NullChar, "DoShowColors", CStr(False))
	End Sub
	Private Sub chkAllowBanned_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkAllowBanned.CheckStateChanged
		If GetSetting("NameSpoofer", ControlChars.NullChar, "DoShowAlert", "1") = "1" Then enablebanned = MsgBox("Enabling banned characters, such as colors, may produce undesirable side effects, such as a ban from Battle.Net. See the Help file for more information Only enable this if you understand the possible consequences.", MsgBoxStyle.OKCancel + MsgBoxStyle.Critical, "WARNING!") : SaveSetting("NameSpoofer", ControlChars.NullChar, "DoShowAlert", "0")
		If enablebanned = MsgBoxResult.Cancel Then chkAllowBanned.CheckState = System.Windows.Forms.CheckState.Unchecked : Exit Sub
		If chkAllowBanned.CheckState = 1 Then cmbColors.Enabled = True Else cmbColors.Enabled = False
		SaveAppSettings()
		If chkAllowBanned.CheckState = 1 Then Exit Sub
		'If unchecked, parse and remove all illegal characters from nickname
		ParsedNick = txtName.Text
		For ParseChars = 1 To Len(ParsedNick)
			NewNick = NewNick & IIf(IsAcceptable(Asc(Mid(ParsedNick, ParseChars, 1))), Mid(ParsedNick, ParseChars, 1), "")
		Next ParseChars
		txtName.Text = NewNick
	End Sub
	Private Sub chkDisableChecking_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDisableChecking.CheckStateChanged
		If chkDisableChecking.CheckState = 1 Then
			chkRetainNickname.CheckState = System.Windows.Forms.CheckState.Unchecked : chkRetainNickname.Enabled = False
			'UPGRADE_WARNING: Timer property tmrRefresh.Interval cannot have a value of 0. Click for more: ms-help://MS.MSDNVS/vbcon/html/vbup2020.htm
			tmrRefresh.Interval = 0
		End If
		If chkDisableChecking.CheckState = 0 Then chkRetainNickname.Enabled = True : tmrRefresh.Interval = 2000
		SaveAppSettings()
	End Sub
	
	Private Sub chkNoClearNickname_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkNoClearNickname.CheckStateChanged
		SaveAppSettings()
	End Sub
	
	Private Sub chkRetainNickname_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkRetainNickname.CheckStateChanged
		SaveAppSettings()
	End Sub
	
	Private Sub cmbColors_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbColors.ClickEvent
		If Len(txtName.Text) > 21 Then Exit Sub
		'CurrentTextStart = txtName.SelStart
		txtName.Text = VB.Left(txtName.Text, txtName.SelectionStart) & Chr(EncodeToASCII((cmbColors.Text))) & VB.Right(txtName.Text, Len(txtName.Text) - txtName.SelectionStart)
		txtName.Focus()
		txtName.SelectionStart = Len(txtName.Text)
	End Sub
	
	'UPGRADE_WARNING: Form event frmMain.Activate has a new behavior. Click for more: ms-help://MS.MSDNVS/vbcon/html/vbup2065.htm
	Private Sub frmMain_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		txtName.Focus()
	End Sub
	
	Private Sub frmMain_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		If frmMain.DefInstance.WindowState = 1 And Restoring <> True Then
			'UPGRADE_ISSUE: App property App.TaskVisible was not upgraded. Click for more: ms-help://MS.MSDNVS/vbcon/html/vbup2069.htm
			App.TaskVisible = False
			Me.Hide()
			CreateIcon()
		ElseIf frmMain.DefInstance.WindowState = 0 Then 
			Me.Show()
			'UPGRADE_ISSUE: App property App.TaskVisible was not upgraded. Click for more: ms-help://MS.MSDNVS/vbcon/html/vbup2069.htm
			App.TaskVisible = True
			Restoring = False
			DeleteIcon()
		End If
	End Sub
	
	Public Sub mnuSTEnterNickname_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSTEnterNickname.Popup
		mnuSTEnterNickname_Click(eventSender, eventArgs)
	End Sub
	Public Sub mnuSTEnterNickname_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSTEnterNickname.Click
		Dim EnteredNick As String
		EnteredNick = InputBox("Enter the nickname you would like to spoof as:", "Spoof Nickname", "")
		If Len(EnteredNick) < 1 Then Exit Sub
		If Len(EnteredNick) > 21 Then EnteredNick = VB.Left(EnteredNick, 21)
		SpoofHandler((EnteredNick))
	End Sub
	
	Public Sub mnuSTExit_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSTExit.Popup
		mnuSTExit_Click(eventSender, eventArgs)
	End Sub
	Public Sub mnuSTExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSTExit.Click
		Me.Close()
		End
	End Sub
	
	Public Sub mnuSTRestore_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSTRestore.Popup
		mnuSTRestore_Click(eventSender, eventArgs)
	End Sub
	Public Sub mnuSTRestore_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSTRestore.Click
		Restoring = True
		'UPGRADE_ISSUE: App property App.TaskVisible was not upgraded. Click for more: ms-help://MS.MSDNVS/vbcon/html/vbup2069.htm
		App.TaskVisible = True
		Me.Show()
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		DeleteIcon()
	End Sub
	
	Private Sub picSystemTray_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picSystemTray.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		X = X / VB6.TwipsPerPixelX
		Select Case X
			Case WM_LBUTTONDOWN
				'UPGRADE_ISSUE: Form method frmMain.PopupMenu was not upgraded. Click for more: ms-help://MS.MSDNVS/vbcon/html/vbup2064.htm
				Me.PopupMenu(mnuST)
			Case WM_RBUTTONDOWN
				'UPGRADE_ISSUE: Form method frmMain.PopupMenu was not upgraded. Click for more: ms-help://MS.MSDNVS/vbcon/html/vbup2064.htm
				Me.PopupMenu(mnuST)
			Case WM_LBUTTONDBLCLK
				Restoring = True
				'UPGRADE_ISSUE: App property App.TaskVisible was not upgraded. Click for more: ms-help://MS.MSDNVS/vbcon/html/vbup2069.htm
				App.TaskVisible = True
				Me.Show()
				Me.WindowState = System.Windows.Forms.FormWindowState.Normal
				DeleteIcon()
		End Select
	End Sub
	
	Private Sub txtName_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtName.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If chkAllowBanned.CheckState = 0 Then
			If IsAcceptable(KeyAscii) = False Then KeyAscii = 0
		End If
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub cmdAdvanced_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdvanced.Click
		If cmdAdvanced.Text = "&Advanced >>" Then ExpandApp() : Exit Sub
		If cmdAdvanced.Text = "<< &Basic" Then ShrinkApp()
	End Sub
	
	Private Sub cmdSpoof_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSpoof.Click
		'Decode colors
		SpoofHandler((txtName.Text))
		txtName.Focus()
	End Sub
	
	Private Sub SpoofHandler(ByRef Nickname As String)
		For DecodeASCII = 1 To Len(Nickname)
			If Asc(Mid(Nickname, DecodeASCII, 1)) > 13 And Asc(Mid(Nickname, DecodeASCII, 1)) < 21 Then
				Mid(Nickname, DecodeASCII, 1) = Chr(DecodeToBWCompatible(Asc(Mid(Nickname, DecodeASCII, 1))))
			End If
		Next DecodeASCII
		'Change nickname
		NickResult = ChangeNick(Nickname)
		If NickResult <> "" Then MsgBox(NickResult, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, "Error while changing nickname")
		lblStatus.Text = ReadNick
		If chkNoClearNickname.CheckState = 0 Then txtName.Text = ""
		'Set current nickname to retain throughout session
		If chkRetainNickname.CheckState = 1 Then RetainedNick = lblStatus.Text
	End Sub
	
	Private Sub cmdSpoof_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmdSpoof.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		If Button = 2 And Shift = 1 Then
			If Y < 80 Then Y = 5
			frmOptions.SetBounds(VB6.TwipsToPixelsX(X + 700), VB6.TwipsToPixelsY(Y + 700), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		End If
	End Sub
	
	Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		AddOfficeBorder(cmdSpoof.Handle.ToInt32)
		AddOfficeBorder(cmdAdvanced.Handle.ToInt32)
		'AddOfficeBorder (txtName.hWnd)
		'AddOfficeBorder (cmbColors.hWnd)
		AddOfficeBorder(frmModify.Handle.ToInt32)
		AddOfficeBorder(frmStatus.Handle.ToInt32)
		AddOfficeBorder(frmColors.Handle.ToInt32)
		AddOfficeBorder(frmOptions.Handle.ToInt32)
		AddOfficeBorder(txtName.Handle.ToInt32)
		'Do routine app loading shit
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
	
	'UPGRADE_WARNING: Form event frmMain.Unload has a new behavior. Click for more: ms-help://MS.MSDNVS/vbcon/html/vbup2065.htm
	Private Sub frmMain_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
		SaveSetting("NameSpoofer", "Position", "Top", CStr(VB6.PixelsToTwipsY(frmMain.DefInstance.Top)))
		SaveSetting("NameSpoofer", "Position", "Left", CStr(VB6.PixelsToTwipsX(frmMain.DefInstance.Left)))
		DeleteIcon()
	End Sub
	
	Private Sub tmrRefresh_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrRefresh.Tick
		CurrentNick = ReadNick
		lblStatus.Text = ReadNick
		If chkRetainNickname.CheckState = 1 And RetainedNick = "" Then RetainedNick = CurrentNick
		If chkRetainNickname.CheckState = 1 And RetainedNick <> CurrentNick Then
			If VB.Left(RetainedNick, 7) <> "Error 0" Then ChangeNick(RetainedNick)
		End If
	End Sub
End Class