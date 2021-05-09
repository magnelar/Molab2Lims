<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmOptions
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
    Public WithEvents lblWilab0 As System.Windows.Forms.Label
	Public WithEvents lblWilab4 As System.Windows.Forms.Label
	Public WithEvents lblWilab1 As System.Windows.Forms.Label
	Public WithEvents lblWilab2 As System.Windows.Forms.Label
	Public WithEvents lblWilab3 As System.Windows.Forms.Label
	Public WithEvents WilabFRame As System.Windows.Forms.GroupBox
	Public WithEvents txtBBPath As System.Windows.Forms.TextBox
	Public WithEvents TxtCountRetry As System.Windows.Forms.TextBox
	Public WithEvents txtPath As System.Windows.Forms.TextBox
	Public WithEvents txtSecBetweenTrans As System.Windows.Forms.TextBox
	Public WithEvents chkAutoRetry As System.Windows.Forms.CheckBox
	Public WithEvents txtMinBetwRetry As System.Windows.Forms.TextBox
	Public WithEvents txtOldFiles As System.Windows.Forms.TextBox
	Public WithEvents lblBBPath As System.Windows.Forms.Label
	Public WithEvents lblNumRetry As System.Windows.Forms.Label
	Public WithEvents lblFilePath As System.Windows.Forms.Label
	Public WithEvents lblSecBetw As System.Windows.Forms.Label
	Public WithEvents lblMinBetwRetry As System.Windows.Forms.Label
	Public WithEvents lblOldFiles As System.Windows.Forms.Label
	Public WithEvents frmKonfig As System.Windows.Forms.GroupBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents OKButton As System.Windows.Forms.Button
	Public WithEvents ImageList1 As System.Windows.Forms.ImageList
    '2.0.0Public WithEvents Line2 As Microsoft.VisualBasic.PowerPacks.LineShape
    Public WithEvents cmdOnOff As System.Windows.Forms.Button
	Public WithEvents cmdServEdit As System.Windows.Forms.Button
    Public WithEvents txtPassW As System.Windows.Forms.TextBox
	Public WithEvents txtUser As System.Windows.Forms.TextBox
	Public WithEvents txtServ As System.Windows.Forms.TextBox
	Public WithEvents txtOrgUnit As System.Windows.Forms.TextBox
    '2.0.0 Public WithEvents Line1 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents lblPassW As System.Windows.Forms.Label
	Public WithEvents lblUser As System.Windows.Forms.Label
	Public WithEvents lblOrgUnit As System.Windows.Forms.Label
	Public WithEvents lblServer As System.Windows.Forms.Label
	Public WithEvents frmServer As System.Windows.Forms.GroupBox
    '   Public WithEvents chkWilab As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
    'Public WithEvents lstOrigin As Microsoft.VisualBasic.Compatibility.VB6.ListBoxArray
    'Public WithEvents lstType As Microsoft.VisualBasic.Compatibility.VB6.ListBoxArray
    'Public WithEvents txtWilabLen As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'Public WithEvents txtWilabPath As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'Public WithEvents txtWilabPos As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'Public WithEvents txtWilabText As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'Public WithEvents ShapeContainer2 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
    'Public WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOptions))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.WilabFRame = New System.Windows.Forms.GroupBox()
        Me.txtWPath = New System.Windows.Forms.TextBox()
        Me.txtWTekst = New System.Windows.Forms.TextBox()
        Me.txtWLen = New System.Windows.Forms.TextBox()
        Me.txtWPos = New System.Windows.Forms.TextBox()
        Me.chkWLab = New System.Windows.Forms.CheckBox()
        Me.lblWilab0 = New System.Windows.Forms.Label()
        Me.lblWilab4 = New System.Windows.Forms.Label()
        Me.lblWilab1 = New System.Windows.Forms.Label()
        Me.lblWilab2 = New System.Windows.Forms.Label()
        Me.lblWilab3 = New System.Windows.Forms.Label()
        Me.frmKonfig = New System.Windows.Forms.GroupBox()
        Me.TxtMetLen = New System.Windows.Forms.TextBox()
        Me.LblMetLen = New System.Windows.Forms.Label()
        Me.txtBBPath = New System.Windows.Forms.TextBox()
        Me.TxtCountRetry = New System.Windows.Forms.TextBox()
        Me.txtPath = New System.Windows.Forms.TextBox()
        Me.txtSecBetweenTrans = New System.Windows.Forms.TextBox()
        Me.chkAutoRetry = New System.Windows.Forms.CheckBox()
        Me.txtMinBetwRetry = New System.Windows.Forms.TextBox()
        Me.txtOldFiles = New System.Windows.Forms.TextBox()
        Me.lblBBPath = New System.Windows.Forms.Label()
        Me.lblNumRetry = New System.Windows.Forms.Label()
        Me.lblFilePath = New System.Windows.Forms.Label()
        Me.lblSecBetw = New System.Windows.Forms.Label()
        Me.lblMinBetwRetry = New System.Windows.Forms.Label()
        Me.lblOldFiles = New System.Windows.Forms.Label()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.frmServer = New System.Windows.Forms.GroupBox()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.cmdOnOff = New System.Windows.Forms.Button()
        Me.cmdServEdit = New System.Windows.Forms.Button()
        Me.txtPassW = New System.Windows.Forms.TextBox()
        Me.txtUser = New System.Windows.Forms.TextBox()
        Me.txtServ = New System.Windows.Forms.TextBox()
        Me.txtOrgUnit = New System.Windows.Forms.TextBox()
        Me.lblPassW = New System.Windows.Forms.Label()
        Me.lblUser = New System.Windows.Forms.Label()
        Me.lblOrgUnit = New System.Windows.Forms.Label()
        Me.lblServer = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RdBtnEng = New System.Windows.Forms.RadioButton()
        Me.RadBtnNorw = New System.Windows.Forms.RadioButton()
        Me.WilabFRame.SuspendLayout()
        Me.frmKonfig.SuspendLayout()
        Me.frmServer.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'WilabFRame
        '
        Me.WilabFRame.BackColor = System.Drawing.SystemColors.Control
        Me.WilabFRame.Controls.Add(Me.txtWPath)
        Me.WilabFRame.Controls.Add(Me.txtWTekst)
        Me.WilabFRame.Controls.Add(Me.txtWLen)
        Me.WilabFRame.Controls.Add(Me.txtWPos)
        Me.WilabFRame.Controls.Add(Me.chkWLab)
        Me.WilabFRame.Controls.Add(Me.lblWilab0)
        Me.WilabFRame.Controls.Add(Me.lblWilab4)
        Me.WilabFRame.Controls.Add(Me.lblWilab1)
        Me.WilabFRame.Controls.Add(Me.lblWilab2)
        Me.WilabFRame.Controls.Add(Me.lblWilab3)
        Me.WilabFRame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.WilabFRame.Location = New System.Drawing.Point(12, 426)
        Me.WilabFRame.Name = "WilabFRame"
        Me.WilabFRame.Padding = New System.Windows.Forms.Padding(0)
        Me.WilabFRame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.WilabFRame.Size = New System.Drawing.Size(402, 80)
        Me.WilabFRame.TabIndex = 49
        Me.WilabFRame.TabStop = False
        Me.WilabFRame.Text = "Gjennkjenning av ikke Elms data"
        Me.WilabFRame.Visible = False
        '
        'txtWPath
        '
        Me.txtWPath.Location = New System.Drawing.Point(226, 44)
        Me.txtWPath.Name = "txtWPath"
        Me.txtWPath.Size = New System.Drawing.Size(131, 20)
        Me.txtWPath.TabIndex = 69
        '
        'txtWTekst
        '
        Me.txtWTekst.Location = New System.Drawing.Point(147, 42)
        Me.txtWTekst.Name = "txtWTekst"
        Me.txtWTekst.Size = New System.Drawing.Size(73, 20)
        Me.txtWTekst.TabIndex = 68
        Me.txtWTekst.Text = "-"
        '
        'txtWLen
        '
        Me.txtWLen.Location = New System.Drawing.Point(106, 42)
        Me.txtWLen.Name = "txtWLen"
        Me.txtWLen.Size = New System.Drawing.Size(26, 20)
        Me.txtWLen.TabIndex = 67
        Me.txtWLen.Text = "1"
        '
        'txtWPos
        '
        Me.txtWPos.Location = New System.Drawing.Point(60, 42)
        Me.txtWPos.Name = "txtWPos"
        Me.txtWPos.Size = New System.Drawing.Size(37, 20)
        Me.txtWPos.TabIndex = 66
        Me.txtWPos.Text = "5"
        '
        'chkWLab
        '
        Me.chkWLab.AutoSize = True
        Me.chkWLab.Location = New System.Drawing.Point(10, 44)
        Me.chkWLab.Name = "chkWLab"
        Me.chkWLab.Size = New System.Drawing.Size(53, 17)
        Me.chkWLab.TabIndex = 65
        Me.chkWLab.Text = "Wilab"
        Me.chkWLab.UseVisualStyleBackColor = True
        '
        'lblWilab0
        '
        Me.lblWilab0.BackColor = System.Drawing.SystemColors.Control
        Me.lblWilab0.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWilab0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWilab0.Location = New System.Drawing.Point(8, 24)
        Me.lblWilab0.Name = "lblWilab0"
        Me.lblWilab0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWilab0.Size = New System.Drawing.Size(37, 17)
        Me.lblWilab0.TabIndex = 63
        Me.lblWilab0.Text = "Brukes"
        '
        'lblWilab4
        '
        Me.lblWilab4.BackColor = System.Drawing.SystemColors.Control
        Me.lblWilab4.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWilab4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWilab4.Location = New System.Drawing.Point(256, 24)
        Me.lblWilab4.Name = "lblWilab4"
        Me.lblWilab4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWilab4.Size = New System.Drawing.Size(85, 17)
        Me.lblWilab4.TabIndex = 62
        Me.lblWilab4.Text = "Filsti"
        '
        'lblWilab1
        '
        Me.lblWilab1.BackColor = System.Drawing.SystemColors.Control
        Me.lblWilab1.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWilab1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWilab1.Location = New System.Drawing.Point(56, 22)
        Me.lblWilab1.Name = "lblWilab1"
        Me.lblWilab1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWilab1.Size = New System.Drawing.Size(61, 17)
        Me.lblWilab1.TabIndex = 57
        Me.lblWilab1.Text = "Pos. for test"
        '
        'lblWilab2
        '
        Me.lblWilab2.BackColor = System.Drawing.SystemColors.Control
        Me.lblWilab2.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWilab2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWilab2.Location = New System.Drawing.Point(120, 24)
        Me.lblWilab2.Name = "lblWilab2"
        Me.lblWilab2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWilab2.Size = New System.Drawing.Size(41, 17)
        Me.lblWilab2.TabIndex = 56
        Me.lblWilab2.Text = "Lengde"
        '
        'lblWilab3
        '
        Me.lblWilab3.BackColor = System.Drawing.SystemColors.Control
        Me.lblWilab3.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWilab3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWilab3.Location = New System.Drawing.Point(168, 24)
        Me.lblWilab3.Name = "lblWilab3"
        Me.lblWilab3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWilab3.Size = New System.Drawing.Size(85, 17)
        Me.lblWilab3.TabIndex = 55
        Me.lblWilab3.Text = "Innhold i streng"
        '
        'frmKonfig
        '
        Me.frmKonfig.BackColor = System.Drawing.SystemColors.Control
        Me.frmKonfig.Controls.Add(Me.TxtMetLen)
        Me.frmKonfig.Controls.Add(Me.LblMetLen)
        Me.frmKonfig.Controls.Add(Me.txtBBPath)
        Me.frmKonfig.Controls.Add(Me.TxtCountRetry)
        Me.frmKonfig.Controls.Add(Me.txtPath)
        Me.frmKonfig.Controls.Add(Me.txtSecBetweenTrans)
        Me.frmKonfig.Controls.Add(Me.chkAutoRetry)
        Me.frmKonfig.Controls.Add(Me.txtMinBetwRetry)
        Me.frmKonfig.Controls.Add(Me.txtOldFiles)
        Me.frmKonfig.Controls.Add(Me.lblBBPath)
        Me.frmKonfig.Controls.Add(Me.lblNumRetry)
        Me.frmKonfig.Controls.Add(Me.lblFilePath)
        Me.frmKonfig.Controls.Add(Me.lblSecBetw)
        Me.frmKonfig.Controls.Add(Me.lblMinBetwRetry)
        Me.frmKonfig.Controls.Add(Me.lblOldFiles)
        Me.frmKonfig.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmKonfig.Location = New System.Drawing.Point(16, 16)
        Me.frmKonfig.Name = "frmKonfig"
        Me.frmKonfig.Padding = New System.Windows.Forms.Padding(0)
        Me.frmKonfig.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmKonfig.Size = New System.Drawing.Size(289, 171)
        Me.frmKonfig.TabIndex = 24
        Me.frmKonfig.TabStop = False
        Me.frmKonfig.Text = "Konfigurasjon"
        '
        'TxtMetLen
        '
        Me.TxtMetLen.AcceptsReturn = True
        Me.TxtMetLen.BackColor = System.Drawing.SystemColors.Window
        Me.TxtMetLen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtMetLen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtMetLen.Location = New System.Drawing.Point(110, 139)
        Me.TxtMetLen.MaxLength = 0
        Me.TxtMetLen.Name = "TxtMetLen"
        Me.TxtMetLen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtMetLen.Size = New System.Drawing.Size(35, 20)
        Me.TxtMetLen.TabIndex = 50
        Me.TxtMetLen.Text = "8"
        '
        'LblMetLen
        '
        Me.LblMetLen.BackColor = System.Drawing.SystemColors.Control
        Me.LblMetLen.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblMetLen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblMetLen.Location = New System.Drawing.Point(16, 139)
        Me.LblMetLen.Name = "LblMetLen"
        Me.LblMetLen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblMetLen.Size = New System.Drawing.Size(87, 17)
        Me.LblMetLen.TabIndex = 49
        Me.LblMetLen.Text = "Metode lengde"
        '
        'txtBBPath
        '
        Me.txtBBPath.AcceptsReturn = True
        Me.txtBBPath.BackColor = System.Drawing.SystemColors.Window
        Me.txtBBPath.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBBPath.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBBPath.Location = New System.Drawing.Point(64, 188)
        Me.txtBBPath.MaxLength = 0
        Me.txtBBPath.Name = "txtBBPath"
        Me.txtBBPath.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBBPath.Size = New System.Drawing.Size(185, 20)
        Me.txtBBPath.TabIndex = 47
        Me.txtBBPath.Visible = False
        '
        'TxtCountRetry
        '
        Me.TxtCountRetry.AcceptsReturn = True
        Me.TxtCountRetry.BackColor = System.Drawing.SystemColors.Window
        Me.TxtCountRetry.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtCountRetry.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtCountRetry.Location = New System.Drawing.Point(248, 94)
        Me.TxtCountRetry.MaxLength = 0
        Me.TxtCountRetry.Name = "TxtCountRetry"
        Me.TxtCountRetry.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtCountRetry.Size = New System.Drawing.Size(20, 20)
        Me.TxtCountRetry.TabIndex = 46
        Me.TxtCountRetry.Text = "2"
        '
        'txtPath
        '
        Me.txtPath.AcceptsReturn = True
        Me.txtPath.BackColor = System.Drawing.SystemColors.Window
        Me.txtPath.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPath.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPath.Location = New System.Drawing.Point(56, 20)
        Me.txtPath.MaxLength = 0
        Me.txtPath.Name = "txtPath"
        Me.txtPath.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPath.Size = New System.Drawing.Size(212, 20)
        Me.txtPath.TabIndex = 29
        '
        'txtSecBetweenTrans
        '
        Me.txtSecBetweenTrans.AcceptsReturn = True
        Me.txtSecBetweenTrans.BackColor = System.Drawing.SystemColors.Window
        Me.txtSecBetweenTrans.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSecBetweenTrans.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSecBetweenTrans.Location = New System.Drawing.Point(235, 51)
        Me.txtSecBetweenTrans.MaxLength = 0
        Me.txtSecBetweenTrans.Name = "txtSecBetweenTrans"
        Me.txtSecBetweenTrans.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSecBetweenTrans.Size = New System.Drawing.Size(33, 20)
        Me.txtSecBetweenTrans.TabIndex = 28
        Me.txtSecBetweenTrans.TabStop = False
        Me.txtSecBetweenTrans.Text = "10"
        '
        'chkAutoRetry
        '
        Me.chkAutoRetry.BackColor = System.Drawing.SystemColors.Control
        Me.chkAutoRetry.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkAutoRetry.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkAutoRetry.Location = New System.Drawing.Point(16, 74)
        Me.chkAutoRetry.Name = "chkAutoRetry"
        Me.chkAutoRetry.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkAutoRetry.Size = New System.Drawing.Size(221, 17)
        Me.chkAutoRetry.TabIndex = 27
        Me.chkAutoRetry.Text = "Auromatisk forsøk på avviste prøver"
        Me.chkAutoRetry.UseVisualStyleBackColor = False
        '
        'txtMinBetwRetry
        '
        Me.txtMinBetwRetry.AcceptsReturn = True
        Me.txtMinBetwRetry.BackColor = System.Drawing.SystemColors.Window
        Me.txtMinBetwRetry.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMinBetwRetry.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMinBetwRetry.Location = New System.Drawing.Point(132, 96)
        Me.txtMinBetwRetry.MaxLength = 0
        Me.txtMinBetwRetry.Name = "txtMinBetwRetry"
        Me.txtMinBetwRetry.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMinBetwRetry.Size = New System.Drawing.Size(20, 20)
        Me.txtMinBetwRetry.TabIndex = 26
        Me.txtMinBetwRetry.Text = "20"
        '
        'txtOldFiles
        '
        Me.txtOldFiles.AcceptsReturn = True
        Me.txtOldFiles.BackColor = System.Drawing.SystemColors.Window
        Me.txtOldFiles.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtOldFiles.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtOldFiles.Location = New System.Drawing.Point(245, 115)
        Me.txtOldFiles.MaxLength = 0
        Me.txtOldFiles.Name = "txtOldFiles"
        Me.txtOldFiles.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtOldFiles.Size = New System.Drawing.Size(20, 20)
        Me.txtOldFiles.TabIndex = 25
        Me.txtOldFiles.Text = "60"
        '
        'lblBBPath
        '
        Me.lblBBPath.BackColor = System.Drawing.SystemColors.Control
        Me.lblBBPath.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblBBPath.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBBPath.Location = New System.Drawing.Point(12, 191)
        Me.lblBBPath.Name = "lblBBPath"
        Me.lblBBPath.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblBBPath.Size = New System.Drawing.Size(41, 13)
        Me.lblBBPath.TabIndex = 48
        Me.lblBBPath.Text = "BB Path"
        Me.lblBBPath.Visible = False
        '
        'lblNumRetry
        '
        Me.lblNumRetry.BackColor = System.Drawing.SystemColors.Control
        Me.lblNumRetry.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNumRetry.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNumRetry.Location = New System.Drawing.Point(154, 94)
        Me.lblNumRetry.Name = "lblNumRetry"
        Me.lblNumRetry.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNumRetry.Size = New System.Drawing.Size(87, 17)
        Me.lblNumRetry.TabIndex = 45
        Me.lblNumRetry.Text = "Antall forsøk"
        '
        'lblFilePath
        '
        Me.lblFilePath.BackColor = System.Drawing.SystemColors.Control
        Me.lblFilePath.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFilePath.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFilePath.Location = New System.Drawing.Point(4, 23)
        Me.lblFilePath.Name = "lblFilePath"
        Me.lblFilePath.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFilePath.Size = New System.Drawing.Size(49, 17)
        Me.lblFilePath.TabIndex = 33
        Me.lblFilePath.Text = "Path"
        '
        'lblSecBetw
        '
        Me.lblSecBetw.BackColor = System.Drawing.SystemColors.Control
        Me.lblSecBetw.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSecBetw.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSecBetw.Location = New System.Drawing.Point(12, 51)
        Me.lblSecBetw.Name = "lblSecBetw"
        Me.lblSecBetw.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSecBetw.Size = New System.Drawing.Size(217, 17)
        Me.lblSecBetw.TabIndex = 32
        Me.lblSecBetw.Text = "Antall sekunder mellom hver overføring:"
        '
        'lblMinBetwRetry
        '
        Me.lblMinBetwRetry.BackColor = System.Drawing.SystemColors.Control
        Me.lblMinBetwRetry.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMinBetwRetry.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMinBetwRetry.Location = New System.Drawing.Point(16, 94)
        Me.lblMinBetwRetry.Name = "lblMinBetwRetry"
        Me.lblMinBetwRetry.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMinBetwRetry.Size = New System.Drawing.Size(113, 17)
        Me.lblMinBetwRetry.TabIndex = 31
        Me.lblMinBetwRetry.Text = "Minutter mellom forsøk:"
        '
        'lblOldFiles
        '
        Me.lblOldFiles.BackColor = System.Drawing.SystemColors.Control
        Me.lblOldFiles.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblOldFiles.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOldFiles.Location = New System.Drawing.Point(16, 115)
        Me.lblOldFiles.Name = "lblOldFiles"
        Me.lblOldFiles.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblOldFiles.Size = New System.Drawing.Size(204, 20)
        Me.lblOldFiles.TabIndex = 30
        Me.lblOldFiles.Text = "Fjerning av filer eldre enn dager:"
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(324, 59)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(81, 25)
        Me.cmdCancel.TabIndex = 20
        Me.cmdCancel.Text = "Avbryt"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'OKButton
        '
        Me.OKButton.BackColor = System.Drawing.SystemColors.Control
        Me.OKButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.OKButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.OKButton.Location = New System.Drawing.Point(324, 27)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.OKButton.Size = New System.Drawing.Size(81, 25)
        Me.OKButton.TabIndex = 15
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = False
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ImageList1.Images.SetKeyName(0, "On")
        Me.ImageList1.Images.SetKeyName(1, "Off")
        '
        'frmServer
        '
        Me.frmServer.BackColor = System.Drawing.SystemColors.Control
        Me.frmServer.Controls.Add(Me.txtServer)
        Me.frmServer.Controls.Add(Me.cmdOnOff)
        Me.frmServer.Controls.Add(Me.cmdServEdit)
        Me.frmServer.Controls.Add(Me.txtPassW)
        Me.frmServer.Controls.Add(Me.txtUser)
        Me.frmServer.Controls.Add(Me.txtServ)
        Me.frmServer.Controls.Add(Me.txtOrgUnit)
        Me.frmServer.Controls.Add(Me.lblPassW)
        Me.frmServer.Controls.Add(Me.lblUser)
        Me.frmServer.Controls.Add(Me.lblOrgUnit)
        Me.frmServer.Controls.Add(Me.lblServer)
        Me.frmServer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmServer.Location = New System.Drawing.Point(16, 193)
        Me.frmServer.Name = "frmServer"
        Me.frmServer.Padding = New System.Windows.Forms.Padding(0)
        Me.frmServer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmServer.Size = New System.Drawing.Size(289, 198)
        Me.frmServer.TabIndex = 0
        Me.frmServer.TabStop = False
        Me.frmServer.Text = "Server Administrasjon"
        '
        'txtServer
        '
        Me.txtServer.Enabled = False
        Me.txtServer.Location = New System.Drawing.Point(55, 131)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(202, 20)
        Me.txtServer.TabIndex = 23
        '
        'cmdOnOff
        '
        Me.cmdOnOff.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOnOff.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOnOff.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOnOff.Location = New System.Drawing.Point(148, 157)
        Me.cmdOnOff.Name = "cmdOnOff"
        Me.cmdOnOff.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOnOff.Size = New System.Drawing.Size(113, 25)
        Me.cmdOnOff.TabIndex = 21
        Me.cmdOnOff.Text = "Logg På/Av"
        Me.cmdOnOff.UseVisualStyleBackColor = False
        '
        'cmdServEdit
        '
        Me.cmdServEdit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdServEdit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdServEdit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdServEdit.Location = New System.Drawing.Point(18, 20)
        Me.cmdServEdit.Name = "cmdServEdit"
        Me.cmdServEdit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdServEdit.Size = New System.Drawing.Size(75, 24)
        Me.cmdServEdit.TabIndex = 19
        Me.cmdServEdit.Text = "Registrer"
        Me.cmdServEdit.UseVisualStyleBackColor = False
        '
        'txtPassW
        '
        Me.txtPassW.AcceptsReturn = True
        Me.txtPassW.BackColor = System.Drawing.SystemColors.Window
        Me.txtPassW.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPassW.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPassW.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtPassW.Location = New System.Drawing.Point(188, 96)
        Me.txtPassW.MaxLength = 0
        Me.txtPassW.Name = "txtPassW"
        Me.txtPassW.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassW.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPassW.Size = New System.Drawing.Size(73, 20)
        Me.txtPassW.TabIndex = 8
        '
        'txtUser
        '
        Me.txtUser.AcceptsReturn = True
        Me.txtUser.BackColor = System.Drawing.SystemColors.Window
        Me.txtUser.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUser.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUser.Location = New System.Drawing.Point(188, 70)
        Me.txtUser.MaxLength = 0
        Me.txtUser.Name = "txtUser"
        Me.txtUser.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUser.Size = New System.Drawing.Size(73, 20)
        Me.txtUser.TabIndex = 5
        '
        'txtServ
        '
        Me.txtServ.AcceptsReturn = True
        Me.txtServ.BackColor = System.Drawing.SystemColors.Window
        Me.txtServ.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtServ.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtServ.Location = New System.Drawing.Point(140, 20)
        Me.txtServ.MaxLength = 0
        Me.txtServ.Name = "txtServ"
        Me.txtServ.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtServ.Size = New System.Drawing.Size(121, 20)
        Me.txtServ.TabIndex = 2
        Me.txtServ.Text = "eits.nnn.elkem"
        '
        'txtOrgUnit
        '
        Me.txtOrgUnit.AcceptsReturn = True
        Me.txtOrgUnit.BackColor = System.Drawing.SystemColors.Window
        Me.txtOrgUnit.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtOrgUnit.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtOrgUnit.Location = New System.Drawing.Point(188, 46)
        Me.txtOrgUnit.MaxLength = 0
        Me.txtOrgUnit.Name = "txtOrgUnit"
        Me.txtOrgUnit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtOrgUnit.Size = New System.Drawing.Size(73, 20)
        Me.txtOrgUnit.TabIndex = 1
        '
        'lblPassW
        '
        Me.lblPassW.BackColor = System.Drawing.SystemColors.Control
        Me.lblPassW.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPassW.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPassW.Location = New System.Drawing.Point(92, 97)
        Me.lblPassW.Name = "lblPassW"
        Me.lblPassW.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPassW.Size = New System.Drawing.Size(49, 17)
        Me.lblPassW.TabIndex = 7
        Me.lblPassW.Text = "Passord"
        '
        'lblUser
        '
        Me.lblUser.BackColor = System.Drawing.SystemColors.Control
        Me.lblUser.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblUser.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblUser.Location = New System.Drawing.Point(92, 73)
        Me.lblUser.Name = "lblUser"
        Me.lblUser.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblUser.Size = New System.Drawing.Size(57, 17)
        Me.lblUser.TabIndex = 6
        Me.lblUser.Text = "Bruker"
        '
        'lblOrgUnit
        '
        Me.lblOrgUnit.BackColor = System.Drawing.SystemColors.Control
        Me.lblOrgUnit.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblOrgUnit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOrgUnit.Location = New System.Drawing.Point(92, 45)
        Me.lblOrgUnit.Name = "lblOrgUnit"
        Me.lblOrgUnit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblOrgUnit.Size = New System.Drawing.Size(90, 21)
        Me.lblOrgUnit.TabIndex = 4
        Me.lblOrgUnit.Text = "Org. enhet"
        '
        'lblServer
        '
        Me.lblServer.BackColor = System.Drawing.SystemColors.Control
        Me.lblServer.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblServer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblServer.Location = New System.Drawing.Point(92, 21)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblServer.Size = New System.Drawing.Size(57, 17)
        Me.lblServer.TabIndex = 3
        Me.lblServer.Text = "Server"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RdBtnEng)
        Me.GroupBox1.Controls.Add(Me.RadBtnNorw)
        Me.GroupBox1.Location = New System.Drawing.Point(325, 101)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(93, 71)
        Me.GroupBox1.TabIndex = 53
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Language"
        '
        'RdBtnEng
        '
        Me.RdBtnEng.AutoSize = True
        Me.RdBtnEng.Location = New System.Drawing.Point(6, 37)
        Me.RdBtnEng.Name = "RdBtnEng"
        Me.RdBtnEng.Size = New System.Drawing.Size(63, 17)
        Me.RdBtnEng.TabIndex = 55
        Me.RdBtnEng.TabStop = True
        Me.RdBtnEng.Text = "Engelsk"
        Me.RdBtnEng.UseVisualStyleBackColor = True
        '
        'RadBtnNorw
        '
        Me.RadBtnNorw.AutoSize = True
        Me.RadBtnNorw.Location = New System.Drawing.Point(6, 19)
        Me.RadBtnNorw.Name = "RadBtnNorw"
        Me.RadBtnNorw.Size = New System.Drawing.Size(53, 17)
        Me.RadBtnNorw.TabIndex = 54
        Me.RadBtnNorw.TabStop = True
        Me.RadBtnNorw.Text = "Norsk"
        Me.RadBtnNorw.UseVisualStyleBackColor = True
        '
        'frmOptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(460, 529)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.WilabFRame)
        Me.Controls.Add(Me.frmKonfig)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.frmServer)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(368, 165)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmOptions"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Optioner"
        Me.WilabFRame.ResumeLayout(False)
        Me.WilabFRame.PerformLayout()
        Me.frmKonfig.ResumeLayout(False)
        Me.frmKonfig.PerformLayout()
        Me.frmServer.ResumeLayout(False)
        Me.frmServer.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents chkWLab As System.Windows.Forms.CheckBox
    Friend WithEvents txtWPath As System.Windows.Forms.TextBox
    Friend WithEvents txtWTekst As System.Windows.Forms.TextBox
    Friend WithEvents txtWLen As System.Windows.Forms.TextBox
    Friend WithEvents txtWPos As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents RdBtnEng As RadioButton
    Friend WithEvents RadBtnNorw As RadioButton
    Public WithEvents TxtMetLen As TextBox
    Public WithEvents LblMetLen As Label
#End Region
End Class