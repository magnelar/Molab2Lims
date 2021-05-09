<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FrmMain
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents mnuLog As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuOption As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents sep1 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuAbout As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents sep2 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuTools As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents txtErrCount As System.Windows.Forms.TextBox
	Public WithEvents txtQueCount As System.Windows.Forms.TextBox
	Public WithEvents txtMinBetwRej As System.Windows.Forms.TextBox
	Public WithEvents txtLastTimeStamp As System.Windows.Forms.TextBox
	Public WithEvents _sbStatusBar_Panel1 As System.Windows.Forms.ToolStripStatusLabel
	Public WithEvents _sbStatusBar_Panel2 As System.Windows.Forms.ToolStripStatusLabel
	Public WithEvents _sbStatusBar_Panel3 As System.Windows.Forms.ToolStripStatusLabel
	Public WithEvents sbStatusBar As System.Windows.Forms.StatusStrip
	Public WithEvents cmdRejFileStatus As System.Windows.Forms.Button
	Public WithEvents ImageList1 As System.Windows.Forms.ImageList
    Public WithEvents txtSecToTrans As System.Windows.Forms.TextBox
	Public WithEvents btnStopforDebug As System.Windows.Forms.Button
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents txtStatusDetail As System.Windows.Forms.TextBox
	Public WithEvents txtStatus As System.Windows.Forms.TextBox
	Public WithEvents btnStop As System.Windows.Forms.Button
	Public WithEvents btnStart As System.Windows.Forms.Button
	Public WithEvents btnRetry As System.Windows.Forms.Button
	Public WithEvents lstRejFiles As System.Windows.Forms.ListBox
	Public WithEvents lstLogFiles As System.Windows.Forms.ListBox
	Public WithEvents Timer2 As System.Windows.Forms.Timer
	Public WithEvents lblErrCount As System.Windows.Forms.Label
	Public WithEvents lblQueCount As System.Windows.Forms.Label
	Public WithEvents lblMinToNextRej As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label10 As System.Windows.Forms.Label
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmMain))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtLastTimeStamp = New System.Windows.Forms.TextBox()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtStatus = New System.Windows.Forms.TextBox()
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuTools = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLog = New System.Windows.Forms.ToolStripMenuItem()
        Me.LogToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FeilToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOption = New System.Windows.Forms.ToolStripMenuItem()
        Me.sep1 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuAbout = New System.Windows.Forms.ToolStripMenuItem()
        Me.sep2 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.txtErrCount = New System.Windows.Forms.TextBox()
        Me.txtQueCount = New System.Windows.Forms.TextBox()
        Me.txtMinBetwRej = New System.Windows.Forms.TextBox()
        Me.sbStatusBar = New System.Windows.Forms.StatusStrip()
        Me._sbStatusBar_Panel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me._sbStatusBar_Panel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me._sbStatusBar_Panel3 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.cmdRejFileStatus = New System.Windows.Forms.Button()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.txtSecToTrans = New System.Windows.Forms.TextBox()
        Me.btnStopforDebug = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.txtStatusDetail = New System.Windows.Forms.TextBox()
        Me.btnRetry = New System.Windows.Forms.Button()
        Me.lstRejFiles = New System.Windows.Forms.ListBox()
        Me.lstLogFiles = New System.Windows.Forms.ListBox()
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.lblErrCount = New System.Windows.Forms.Label()
        Me.lblQueCount = New System.Windows.Forms.Label()
        Me.lblMinToNextRej = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.MainMenu1.SuspendLayout()
        Me.sbStatusBar.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtLastTimeStamp
        '
        Me.txtLastTimeStamp.AcceptsReturn = True
        Me.txtLastTimeStamp.BackColor = System.Drawing.SystemColors.Window
        Me.txtLastTimeStamp.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLastTimeStamp.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLastTimeStamp.Location = New System.Drawing.Point(11, 424)
        Me.txtLastTimeStamp.MaxLength = 0
        Me.txtLastTimeStamp.Name = "txtLastTimeStamp"
        Me.txtLastTimeStamp.ReadOnly = True
        Me.txtLastTimeStamp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLastTimeStamp.Size = New System.Drawing.Size(113, 20)
        Me.txtLastTimeStamp.TabIndex = 17
        Me.txtLastTimeStamp.TabStop = False
        Me.ToolTip1.SetToolTip(Me.txtLastTimeStamp, "Lagringstidspunkt for loggefilen")
        Me.txtLastTimeStamp.Visible = False
        '
        'btnStop
        '
        Me.btnStop.BackColor = System.Drawing.SystemColors.Control
        Me.btnStop.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnStop.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnStop.Location = New System.Drawing.Point(208, 32)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnStop.Size = New System.Drawing.Size(130, 25)
        Me.btnStop.TabIndex = 6
        Me.btnStop.Text = "start"
        Me.ToolTip1.SetToolTip(Me.btnStop, "停止传输")
        Me.btnStop.UseVisualStyleBackColor = False
        '
        'btnStart
        '
        Me.btnStart.BackColor = System.Drawing.SystemColors.Control
        Me.btnStart.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnStart.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnStart.Location = New System.Drawing.Point(344, 32)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnStart.Size = New System.Drawing.Size(117, 25)
        Me.btnStart.TabIndex = 5
        Me.btnStart.Text = "stopp"
        Me.ToolTip1.SetToolTip(Me.btnStart, "开始传输")
        Me.btnStart.UseVisualStyleBackColor = False
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(22, 32)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(49, 17)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Status:"
        Me.ToolTip1.SetToolTip(Me.Label5, "状态")
        '
        'txtStatus
        '
        Me.txtStatus.AcceptsReturn = True
        Me.txtStatus.BackColor = System.Drawing.SystemColors.Window
        Me.txtStatus.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStatus.Enabled = False
        Me.txtStatus.ForeColor = System.Drawing.Color.Black
        Me.txtStatus.Location = New System.Drawing.Point(64, 32)
        Me.txtStatus.MaxLength = 0
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.ReadOnly = True
        Me.txtStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStatus.Size = New System.Drawing.Size(129, 20)
        Me.txtStatus.TabIndex = 7
        Me.txtStatus.TabStop = False
        Me.ToolTip1.SetToolTip(Me.txtStatus, "Her kommer status tekst")
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuTools})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(480, 24)
        Me.MainMenu1.TabIndex = 24
        '
        'mnuTools
        '
        Me.mnuTools.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuLog, Me.mnuOption, Me.sep1, Me.mnuAbout, Me.sep2, Me.mnuExit})
        Me.mnuTools.Name = "mnuTools"
        Me.mnuTools.Size = New System.Drawing.Size(59, 20)
        Me.mnuTools.Text = "Verktøy"
        Me.mnuTools.ToolTipText = "工具"
        '
        'mnuLog
        '
        Me.mnuLog.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.LogToolStripMenuItem, Me.FeilToolStripMenuItem})
        Me.mnuLog.Name = "mnuLog"
        Me.mnuLog.Size = New System.Drawing.Size(130, 22)
        Me.mnuLog.Text = "Log..."
        '
        'LogToolStripMenuItem
        '
        Me.LogToolStripMenuItem.Name = "LogToolStripMenuItem"
        Me.LogToolStripMenuItem.Size = New System.Drawing.Size(94, 22)
        Me.LogToolStripMenuItem.Text = "Log"
        '
        'FeilToolStripMenuItem
        '
        Me.FeilToolStripMenuItem.Name = "FeilToolStripMenuItem"
        Me.FeilToolStripMenuItem.Size = New System.Drawing.Size(94, 22)
        Me.FeilToolStripMenuItem.Text = "Feil"
        '
        'mnuOption
        '
        Me.mnuOption.Name = "mnuOption"
        Me.mnuOption.Size = New System.Drawing.Size(130, 22)
        Me.mnuOption.Text = "Optioner..."
        Me.mnuOption.ToolTipText = "选项"
        '
        'sep1
        '
        Me.sep1.Name = "sep1"
        Me.sep1.Size = New System.Drawing.Size(127, 6)
        '
        'mnuAbout
        '
        Me.mnuAbout.Name = "mnuAbout"
        Me.mnuAbout.Size = New System.Drawing.Size(130, 22)
        Me.mnuAbout.Text = "Om..."
        '
        'sep2
        '
        Me.sep2.Name = "sep2"
        Me.sep2.Size = New System.Drawing.Size(127, 6)
        '
        'mnuExit
        '
        Me.mnuExit.Name = "mnuExit"
        Me.mnuExit.Size = New System.Drawing.Size(130, 22)
        Me.mnuExit.Text = "E&xit"
        '
        'txtErrCount
        '
        Me.txtErrCount.AcceptsReturn = True
        Me.txtErrCount.BackColor = System.Drawing.SystemColors.Window
        Me.txtErrCount.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtErrCount.Enabled = False
        Me.txtErrCount.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtErrCount.Location = New System.Drawing.Point(425, 128)
        Me.txtErrCount.MaxLength = 0
        Me.txtErrCount.Name = "txtErrCount"
        Me.txtErrCount.ReadOnly = True
        Me.txtErrCount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtErrCount.Size = New System.Drawing.Size(32, 20)
        Me.txtErrCount.TabIndex = 22
        Me.txtErrCount.TabStop = False
        '
        'txtQueCount
        '
        Me.txtQueCount.AcceptsReturn = True
        Me.txtQueCount.BackColor = System.Drawing.SystemColors.Window
        Me.txtQueCount.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtQueCount.Enabled = False
        Me.txtQueCount.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtQueCount.Location = New System.Drawing.Point(208, 129)
        Me.txtQueCount.MaxLength = 0
        Me.txtQueCount.Name = "txtQueCount"
        Me.txtQueCount.ReadOnly = True
        Me.txtQueCount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtQueCount.Size = New System.Drawing.Size(27, 20)
        Me.txtQueCount.TabIndex = 21
        Me.txtQueCount.TabStop = False
        '
        'txtMinBetwRej
        '
        Me.txtMinBetwRej.AcceptsReturn = True
        Me.txtMinBetwRej.BackColor = System.Drawing.SystemColors.Window
        Me.txtMinBetwRej.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMinBetwRej.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMinBetwRej.Location = New System.Drawing.Point(208, 460)
        Me.txtMinBetwRej.MaxLength = 0
        Me.txtMinBetwRej.Name = "txtMinBetwRej"
        Me.txtMinBetwRej.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMinBetwRej.Size = New System.Drawing.Size(27, 20)
        Me.txtMinBetwRej.TabIndex = 18
        Me.txtMinBetwRej.Text = "Text1"
        '
        'sbStatusBar
        '
        Me.sbStatusBar.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me._sbStatusBar_Panel1, Me._sbStatusBar_Panel2, Me._sbStatusBar_Panel3})
        Me.sbStatusBar.Location = New System.Drawing.Point(0, 333)
        Me.sbStatusBar.Name = "sbStatusBar"
        Me.sbStatusBar.Size = New System.Drawing.Size(480, 22)
        Me.sbStatusBar.TabIndex = 16
        '
        '_sbStatusBar_Panel1
        '
        Me._sbStatusBar_Panel1.AccessibleName = ""
        Me._sbStatusBar_Panel1.AutoSize = False
        Me._sbStatusBar_Panel1.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me._sbStatusBar_Panel1.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me._sbStatusBar_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me._sbStatusBar_Panel1.Name = "_sbStatusBar_Panel1"
        Me._sbStatusBar_Panel1.Size = New System.Drawing.Size(345, 22)
        Me._sbStatusBar_Panel1.Spring = True
        Me._sbStatusBar_Panel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        '_sbStatusBar_Panel2
        '
        Me._sbStatusBar_Panel2.AutoSize = False
        Me._sbStatusBar_Panel2.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me._sbStatusBar_Panel2.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me._sbStatusBar_Panel2.Margin = New System.Windows.Forms.Padding(0)
        Me._sbStatusBar_Panel2.Name = "_sbStatusBar_Panel2"
        Me._sbStatusBar_Panel2.Size = New System.Drawing.Size(120, 22)
        Me._sbStatusBar_Panel2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        '_sbStatusBar_Panel3
        '
        Me._sbStatusBar_Panel3.AutoSize = False
        Me._sbStatusBar_Panel3.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me._sbStatusBar_Panel3.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me._sbStatusBar_Panel3.Margin = New System.Windows.Forms.Padding(0)
        Me._sbStatusBar_Panel3.Name = "_sbStatusBar_Panel3"
        Me._sbStatusBar_Panel3.Size = New System.Drawing.Size(96, 17)
        Me._sbStatusBar_Panel3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmdRejFileStatus
        '
        Me.cmdRejFileStatus.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRejFileStatus.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRejFileStatus.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRejFileStatus.Location = New System.Drawing.Point(169, 242)
        Me.cmdRejFileStatus.Name = "cmdRejFileStatus"
        Me.cmdRejFileStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRejFileStatus.Size = New System.Drawing.Size(116, 25)
        Me.cmdRejFileStatus.TabIndex = 15
        Me.cmdRejFileStatus.Text = "Rej.file Status"
        Me.cmdRejFileStatus.UseVisualStyleBackColor = False
        '
        'ImageList1
        '
        Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.TransparentColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        '
        'txtSecToTrans
        '
        Me.txtSecToTrans.AcceptsReturn = True
        Me.txtSecToTrans.BackColor = System.Drawing.SystemColors.Window
        Me.txtSecToTrans.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSecToTrans.Enabled = False
        Me.txtSecToTrans.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSecToTrans.Location = New System.Drawing.Point(208, 96)
        Me.txtSecToTrans.MaxLength = 0
        Me.txtSecToTrans.Name = "txtSecToTrans"
        Me.txtSecToTrans.ReadOnly = True
        Me.txtSecToTrans.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSecToTrans.Size = New System.Drawing.Size(27, 20)
        Me.txtSecToTrans.TabIndex = 10
        Me.txtSecToTrans.TabStop = False
        '
        'btnStopforDebug
        '
        Me.btnStopforDebug.BackColor = System.Drawing.SystemColors.Control
        Me.btnStopforDebug.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnStopforDebug.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnStopforDebug.Location = New System.Drawing.Point(343, 387)
        Me.btnStopforDebug.Name = "btnStopforDebug"
        Me.btnStopforDebug.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnStopforDebug.Size = New System.Drawing.Size(105, 25)
        Me.btnStopforDebug.TabIndex = 9
        Me.btnStopforDebug.Text = "Stop for debugging !"
        Me.btnStopforDebug.UseVisualStyleBackColor = False
        Me.btnStopforDebug.Visible = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'txtStatusDetail
        '
        Me.txtStatusDetail.AcceptsReturn = True
        Me.txtStatusDetail.BackColor = System.Drawing.SystemColors.Window
        Me.txtStatusDetail.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStatusDetail.Enabled = False
        Me.txtStatusDetail.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStatusDetail.Location = New System.Drawing.Point(8, 68)
        Me.txtStatusDetail.MaxLength = 0
        Me.txtStatusDetail.Name = "txtStatusDetail"
        Me.txtStatusDetail.ReadOnly = True
        Me.txtStatusDetail.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStatusDetail.Size = New System.Drawing.Size(449, 20)
        Me.txtStatusDetail.TabIndex = 8
        Me.txtStatusDetail.TabStop = False
        Me.txtStatusDetail.Text = "Detaljert status..."
        '
        'btnRetry
        '
        Me.btnRetry.BackColor = System.Drawing.SystemColors.Control
        Me.btnRetry.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnRetry.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnRetry.Location = New System.Drawing.Point(291, 242)
        Me.btnRetry.Name = "btnRetry"
        Me.btnRetry.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnRetry.Size = New System.Drawing.Size(170, 25)
        Me.btnRetry.TabIndex = 4
        Me.btnRetry.Text = "Nytt forsøk på avviste filer"
        Me.btnRetry.UseVisualStyleBackColor = False
        '
        'lstRejFiles
        '
        Me.lstRejFiles.BackColor = System.Drawing.SystemColors.Window
        Me.lstRejFiles.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstRejFiles.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstRejFiles.Location = New System.Drawing.Point(12, 273)
        Me.lstRejFiles.Name = "lstRejFiles"
        Me.lstRejFiles.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstRejFiles.Size = New System.Drawing.Size(453, 43)
        Me.lstRejFiles.TabIndex = 3
        Me.lstRejFiles.TabStop = False
        '
        'lstLogFiles
        '
        Me.lstLogFiles.BackColor = System.Drawing.SystemColors.Window
        Me.lstLogFiles.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstLogFiles.Enabled = False
        Me.lstLogFiles.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lstLogFiles.Location = New System.Drawing.Point(8, 158)
        Me.lstLogFiles.Name = "lstLogFiles"
        Me.lstLogFiles.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstLogFiles.Size = New System.Drawing.Size(453, 69)
        Me.lstLogFiles.Sorted = True
        Me.lstLogFiles.TabIndex = 0
        Me.lstLogFiles.TabStop = False
        '
        'Timer2
        '
        Me.Timer2.Enabled = True
        Me.Timer2.Interval = 5000
        '
        'lblErrCount
        '
        Me.lblErrCount.BackColor = System.Drawing.SystemColors.Control
        Me.lblErrCount.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblErrCount.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblErrCount.Location = New System.Drawing.Point(241, 132)
        Me.lblErrCount.Name = "lblErrCount"
        Me.lblErrCount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblErrCount.Size = New System.Drawing.Size(135, 16)
        Me.lblErrCount.TabIndex = 23
        Me.lblErrCount.Text = "Antall i feil katalog"
        '
        'lblQueCount
        '
        Me.lblQueCount.BackColor = System.Drawing.SystemColors.Control
        Me.lblQueCount.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblQueCount.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblQueCount.Location = New System.Drawing.Point(8, 125)
        Me.lblQueCount.Name = "lblQueCount"
        Me.lblQueCount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblQueCount.Size = New System.Drawing.Size(194, 23)
        Me.lblQueCount.TabIndex = 20
        Me.lblQueCount.Text = "Antall i behandlings katalog"
        '
        'lblMinToNextRej
        '
        Me.lblMinToNextRej.BackColor = System.Drawing.SystemColors.Control
        Me.lblMinToNextRej.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMinToNextRej.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMinToNextRej.Location = New System.Drawing.Point(12, 460)
        Me.lblMinToNextRej.Name = "lblMinToNextRej"
        Me.lblMinToNextRej.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMinToNextRej.Size = New System.Drawing.Size(177, 17)
        Me.lblMinToNextRej.TabIndex = 19
        Me.lblMinToNextRej.Text = "Min.til neste forsøk med rej.filer"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(12, 404)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(117, 17)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Siste oveførte filer"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(8, 100)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(194, 15)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Antall sekunder til neste overføring:"
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(12, 387)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(161, 17)
        Me.Label10.TabIndex = 11
        Me.Label10.Text = "Siste loggefil laget/oppdatert"
        Me.Label10.Visible = False
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(8, 242)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(155, 25)
        Me.Label9.TabIndex = 2
        Me.Label9.Text = "Siste avviste filer fra Oracle"
        '
        'txtServer
        '
        Me.txtServer.Enabled = False
        Me.txtServer.Location = New System.Drawing.Point(318, 100)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(139, 20)
        Me.txtServer.TabIndex = 25
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(241, 102)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 13)
        Me.Label2.TabIndex = 26
        Me.Label2.Text = "Server info"
        '
        'FrmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(480, 355)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtServer)
        Me.Controls.Add(Me.txtErrCount)
        Me.Controls.Add(Me.txtQueCount)
        Me.Controls.Add(Me.txtMinBetwRej)
        Me.Controls.Add(Me.txtLastTimeStamp)
        Me.Controls.Add(Me.sbStatusBar)
        Me.Controls.Add(Me.cmdRejFileStatus)
        Me.Controls.Add(Me.txtSecToTrans)
        Me.Controls.Add(Me.btnStopforDebug)
        Me.Controls.Add(Me.txtStatusDetail)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.btnRetry)
        Me.Controls.Add(Me.lstRejFiles)
        Me.Controls.Add(Me.lstLogFiles)
        Me.Controls.Add(Me.lblErrCount)
        Me.Controls.Add(Me.lblQueCount)
        Me.Controls.Add(Me.lblMinToNextRej)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(43, 65)
        Me.MaximizeBox = False
        Me.Name = "FrmMain"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Molab2Lims"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.sbStatusBar.ResumeLayout(False)
        Me.sbStatusBar.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LogToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FeilToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
#End Region
End Class