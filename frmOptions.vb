Option Strict Off
Option Explicit On
'Molab2Lims bruk
Imports System.Data '1.0.3.2
Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types

'Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmOptions
	Inherits System.Windows.Forms.Form
	
	Dim mServCriteria As String
	Dim iFl As Short '1.6.7
    Private Sub LoadResourceStrings()
        On Error GoTo ErrorL
        '2.0.0 Me.Icon = Lib_GetFromImageList(101, VB6.LoadResConstants.ResIcon)
        ' cmdOnK1.Picture = Lib_GetFromImageList(101, vbResBitmap)
        If gLang = "Nor" Then '1.0.3.1
            RadBtnNorw.Checked = True
            Me.Text = "Optioner"
            frmKonfig.Text = "Konfigurasjon"
            lblFilePath.Text = "Fil sti"
            lblSecBetw.Text = "Antall sekunder mellom hver overføring:"
            chkAutoRetry.Text = "Automatisk forsøk på avviste prøver."
            lblMinBetwRetry.Text = "Min. mellom forsøk"
            lblOldFiles.Text = "Fjerning av filer eldre enn dager."
            RadBtnNorw.Text = "Norwegian"
            RdBtnEng.Text = "English"

            frmServer.Text = "Server"
            cmdServEdit.Text = "Endre"
            lblServer.Text = "Server"
            lblOrgUnit.Text = "Org. enhet"
            lblUser.Text = "Bruker"
            lblPassW.Text = "Passord"

            cmdOnOff.Text = "På/Av"
            OKButton.Text = "Ok"
            cmdCancel.Text = "Avbryt"
            GroupBox1.Text = "Språk"
            '1.0.4.5
            LblMetLen.Text = "Metode Lengde"
        ElseIf gLang = "Eng" Then
            RdBtnEng.Checked = True
            Me.Text = "Options"
            frmKonfig.Text = "Configuration"
            lblFilePath.Text = "File path"
            lblSecBetw.Text = "Seconds between transfering"
            chkAutoRetry.Text = "Automatic try with rejected files"
            lblMinBetwRetry.Text = "Minutes between try"
            lblOldFiles.Text = "Remove files older then number days."

            frmServer.Text = "Server"
            cmdServEdit.Text = "Change"
            lblServer.Text = "Server"
            lblOrgUnit.Text = "Org. unit"
            lblUser.Text = "User"
            lblPassW.Text = "Password"

            cmdOnOff.Text = "On/Off"
            OKButton.Text = "Ok"
            cmdCancel.Text = "Cancel"
            RadBtnNorw.Text = "Norwegian"
            RdBtnEng.Text = "English"
            GroupBox1.Text = "Language"
            '1.0.4.5
            LblMetLen.Text = "Methode Length"

        End If


        Exit Sub

ErrorL:
        Call GenericError("frmMain", "Feil i LoadResourceStrings...", Err.Number, Err.Description, Err.Source)
    End Sub
    Public Sub StartTool()

        If Not gNoErrorHandling Then On Error GoTo Err_StartTool
        'Dim itmX As System.Windows.Forms.ListViewItem
		Dim i As Short
        '1.6.4, 1.6.7
        '1.0.0.3 For iFl = 0 To 1
        '?
        Call FrmMain.ReadIniFil() '1.0.3.8
        'MessageBox.Show("1a " & gServer.ConStr & vbNewLine & gPath & vbNewLine & gOrgUnit & vbNewLine & gServer.User)
        If gNotElms.Use Then
            'chkWilab(iFl).CheckState = System.Windows.Forms.CheckState.Checked
            chkWLab.CheckState = System.Windows.Forms.CheckState.Checked
            txtWLen.Visible = True
            txtWPos.Visible = True
            txtWTekst.Visible = True
            txtWPath.Visible = True
        Else
            chkWLab.CheckState = System.Windows.Forms.CheckState.Unchecked
            txtWLen.Visible = False
            txtWPos.Visible = False
            txtWTekst.Visible = False
            txtWPath.Visible = False

            lblWilab1.Visible = False
            lblWilab2.Visible = False
            lblWilab3.Visible = False
            lblWilab4.Visible = False
        End If
        txtWPos.Text = CStr(gNotElms.Pos - 2) '1.0.7
        txtWLen.Text = CStr(gNotElms.len_Renamed)
        txtWTekst.Text = gNotElms.txt
        txtWPath.Text = gNotElms.Path
        'gServer.dbCon = New ADODB.Connection
        gServer.ServNo = 1
        With gServer
            'itmX = lsvServProp.Items.Add("")
            If .Connected Then
                'itmX.BackColor = Color.Green '2.0.0
                txtServer.BackColor = Color.GreenYellow
            Else
                txtServer.BackColor = Color.WhiteSmoke
                'itmX.BackColor = Color.WhiteSmoke '2.0.0
            End If
            'itmX.Text = CStr(.ServNo)
            txtServer.Text = .ConStr & " - " & .User & " - " & gOrgUnit

        End With
        Call LoadResourceStrings()
        txtPath.Text = gPath
        txtBBPath.Text = gBBPath
        txtSecBetweenTrans.Text = CStr(gSecBetweenTrans)
        txtMinBetwRetry.Text = CStr(gMinBetweenRejTrans)
        TxtCountRetry.Text = CStr(gCountRetry)
        txtOldFiles.Text = CStr(gOldFiles)
        If gAutoRetry Then
            chkAutoRetry.CheckState = System.Windows.Forms.CheckState.Checked
        Else
            chkAutoRetry.CheckState = System.Windows.Forms.CheckState.Unchecked
        End If
        Me.Show() 'vbModal, FrmMain
        'Call SaveParInOracle 1.3.3

Exit_StartTool:
        Exit Sub

Err_StartTool:
        Call GenericError("frmOptions", "Feil i StartTool...", Err.Number, Err.Description, Err.Source)
        Resume Exit_StartTool
    End Sub

    Private Sub CancelButton_Click()
		Me.Close()
		
	End Sub

    Private Sub chkAutoRetry_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkAutoRetry.CheckStateChanged
        If chkAutoRetry.CheckState Then
            txtMinBetwRetry.Enabled = True
        Else
            txtMinBetwRetry.Enabled = False
        End If
    End Sub


    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        'MessageBox.Show("2 " & gServer.ConStr & vbNewLine & gPath & vbNewLine & gOrgUnit & vbNewLine & gServer.User)

        Me.Close()

    End Sub


    Private Sub cmdOnOff_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOnOff.Click
        Try '1.0.3.0

            If gServer.Connected Then
                Dim conn As New OracleConnection(gConnect) '1.0.3.2
                If conn.State Then
                    conn.Close()
                End If
                gServer.Connected = False
                txtServer.BackColor = Color.WhiteSmoke '2.0.0
                'MessageBox.Show(gServer.ConStr)
            Else
                Dim pconnection As New OracleConnection(gConnect)
                pconnection.Open()
                txtServer.BackColor = Color.YellowGreen '2.0.0
                gServer.Connected = True
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            End If
        Catch Ex As Exception
            txtServer.BackColor = Color.LightYellow '1.1.1.1
            txtServer.Text = txtServ.Text & " - " & txtUser.Text & " - " & gOrgUnit & " : Feil!"
            Call GenericError("cmdOnOff_Click", "Feil i cmdOnOff_Click-kode..", Err.Number, Err.Description, Err.Source)
            gServer.Connected = False
            MessageBox.Show(Ex.Message & vbNewLine & gConnect)
        Finally

        End Try
        Exit Sub
    End Sub


    Private Sub cmdServEdit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdServEdit.Click
		If Not gNoErrorHandling Then On Error GoTo Err_cmdServEdit_Click
        'Dim iAct As Short
		Dim lPwd As String
        'iAct = ActServ((lsvServProp.FocusedItem))
        'If iAct > 0 Then
        gOrgUnit = txtOrgUnit.Text '1.0.0.2 flyttet et hakk opp.
        txtServer.Text = txtServ.Text & " - " & txtUser.Text & " - " & gOrgUnit
        'lsvServProp.SelectedItem.SubItems(3) = txtOrgUnit
        With gServer
            .ConStr = txtServ.Text
            .User = txtUser.Text
            lPwd = txtPassW.Text
            .Connected = False
            'Call Encrypt(lPwd, gSifrCode)
            .Pwd = lPwd
        End With
        '1.0.3.9
        gConnect = "DATA SOURCE=" & gServer.ConStr & ";PASSWORD=" & gServer.Pwd & ";PERSIST SECURITY INFO=True;USER ID=" & gServer.User

Exit_cmdServEdit_Click:
        Exit Sub

Err_cmdServEdit_Click:
        Call GenericError("cmdServEdit_Click", "Feil i cmdServEdit_Click-kode..", Err.Number, Err.Description, Err.Source)
        Resume Exit_cmdServEdit_Click
    End Sub



    Private Sub frmOptions_LostFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.LostFocus
		Me.Close()
	End Sub
	'
	'Private Sub lsvServProp_BeforeLabelEdit(Cancel As Integer)
	'Dim rad As Variant
	'rad = LsvServers.SelectedItem
	'MsgBox rad
	'
	'End Sub
	
	
	Private Sub Frame1_DragDrop(ByRef Source As System.Windows.Forms.Control, ByRef X As Single, ByRef y As Single)
		
	End Sub

    Public Function ActServ(ByRef iServNo As Object) As Short
        'Dim i As Short
        ActServ = 1
        'For i = 1 To gServCount
        'If gServer.ServNo = CDbl(lsvServProp.FocusedItem.Text) Then
        '    ActServ = i
        'End If
        'Next i

	End Function
    Private Sub SaveServInfo()
        Dim iMess As Short
        If Not gNoErrorHandling Then On Error GoTo Err_SaveServInfo
        Dim i As Short
        Dim lValid As Boolean
        Dim lPwd As String
        Dim rst As ADODB.Recordset
        rst = New ADODB.Recordset
        rst.Fields.Append("ServNo", ADODB.DataTypeEnum.adSmallInt)
        rst.Fields.Append("ServName", ADODB.DataTypeEnum.adVarChar, 25)
        rst.Fields.Append("User", ADODB.DataTypeEnum.adVarChar, 20)
        rst.Fields.Append("PassW", ADODB.DataTypeEnum.adVarChar, 20)
        rst.Fields.Append("OrgUnit", ADODB.DataTypeEnum.adVarChar, 20)
        rst.Open()
        'objRS.Close
        With gServer
            'krypter passord før lagring
            lPwd = .Pwd
            Call Encrypt(lPwd, gSifrCode)
            '.Pwd = lPwd
            rst.AddNew(New Object() {"ServNo", "ServName", "User", "PassW", "OrgUnit"}, New Object() { .ServNo, .ConStr, .User, lPwd, gOrgUnit})
        End With
        rst.UpdateBatch()
        If Dir(My.Application.Info.DirectoryPath & "\Server.dat") & "" <> "" Then
            lValid = True
        End If
        If Not gNoErrorHandling Then On Error GoTo Err_SaveServInfo
        If lValid Then
            Kill(My.Application.Info.DirectoryPath & "\Server.dat")
        End If
        rst.Save(My.Application.Info.DirectoryPath & "\Server.dat")

        rst.Close()
        rst = Nothing
Exit_SaveServInfo:
        Exit Sub

Err_SaveServInfo:
        Call GenericError("SaveServInfo", "Feil i SaveServInfo-kode..", Err.Number, Err.Description, Err.Source)
        Resume Exit_SaveServInfo
    End Sub


    Private Sub OKButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OKButton.Click
        Call LoadResourceStrings()
        Call SaveServInfo()
        gNotElms.Use = chkWLab.CheckState
        gNotElms.Path = txtWPath.Text
        gNotElms.Pos = CDbl(txtWPos.Text) + 2 ' rev 1.0.7
        gNotElms.len_Renamed = txtWLen.Text
        gNotElms.txt = txtWTekst.Text
        gOrgUnit = txtOrgUnit.Text
        gPath = txtPath.Text
        gBBPath = txtBBPath.Text
        '1.0.0.0 feb 2015
        If txtSecBetweenTrans.Text = "" Then
            gSecBetweenTrans = 10
        Else
            gSecBetweenTrans = CSng(txtSecBetweenTrans.Text)
        End If
        gAutoRetry = chkAutoRetry.CheckState
        If txtMinBetwRetry.Text = "" Then
            gMinBetweenRejTrans = 30
        Else
            gMinBetweenRejTrans = CSng(txtMinBetwRetry.Text)
        End If
        If TxtCountRetry.Text = "" Then
            gCountRetry = 2
        Else
            gCountRetry = CSng(TxtCountRetry.Text)
        End If
        If txtOldFiles.Text = "" Then
            gOldFiles = 60
        Else
            gOldFiles = CSng(txtOldFiles.Text)
        End If
        '1.0.4.5
        gMethLength = TxtMetLen.Text

        Call FrmMain.WriteIniFile()
        'MessageBox.Show("6 " & gServer.ConStr & vbNewLine & gPath & vbNewLine & gOrgUnit & vbNewLine & gServer.User)
        If gPath = "" Then
            gPath = "C:\TOELMS\Xrf"
        End If
        MakeCatStr()
        Me.Close()
        'Call getServInfo '1.3.5
        Call FrmMain.ServList()
    End Sub
    Private Sub CreTypeRecordSet()
        If Not gNoErrorHandling Then On Error GoTo Err_CreTypeRecordSet
        Dim rst As ADODB.Recordset

        ' Instantiate the recordset.
        rst = New ADODB.Recordset

        ' Append fields
        rst.Fields.Append("ServNo", ADODB.DataTypeEnum.adSmallInt)
        rst.Fields.Append("Furnace", ADODB.DataTypeEnum.adVarChar, 10)
        rst.Fields.Append("Type", ADODB.DataTypeEnum.adVarChar, 10)

        ' Put some data in the recordset
        rst.Open()
        rst.AddNew(New Object() {"ServNo", "Furnace", "Type"}, New Object() {1, "ALL", "ALL"})
        rst.AddNew(New Object() {"ServNo", "Furnace", "Type"}, New Object() {2, "FVO09", "TAPS"})
        rst.AddNew(New Object() {"ServNo", "Furnace", "Type"}, New Object() {2, "FVO09", "KL"})

        rst.UpdateBatch()
        Kill(My.Application.Info.DirectoryPath & "\SampType.dat")
        rst.Save(My.Application.Info.DirectoryPath & "\SampType.dat")

        '' Dump the recordset to the Immediate Window
        'rst.MoveFirst
        'Do Until rst.EOF
        '    Debug.Print rst("ServNo"), rst("ServName")
        '    rst.MoveNext
        'Loop
        rst.Close()

Exit_CreTypeRecordSet:
        Exit Sub

Err_CreTypeRecordSet:
        Call GenericError("CreTypeRecordSet", "Feil i CreFurnaceRecordSet-kode..", Err.Number, Err.Description, Err.Source)
        Resume Exit_CreTypeRecordSet
    End Sub

    Private Sub tulle_Click(ByRef Index As Short)
        MsgBox("tull")
    End Sub

    Private Sub chkWLab_CheckedChanged(sender As Object, e As EventArgs) Handles chkWLab.CheckedChanged
        'Dim Index As Short = chkWilab.GetIndex(eventSender)
        '	iFl = Index
        If chkWLab.CheckState Then
            txtWLen.Visible = True
            txtWPos.Visible = True
            txtWTekst.Visible = True
            txtWPath.Visible = True
        Else
            txtWLen.Visible = False
            txtWPos.Visible = False
            txtWTekst.Visible = False
            txtWPath.Visible = False

            lblWilab1.Visible = False
            lblWilab2.Visible = False
            lblWilab3.Visible = False
            lblWilab4.Visible = False
        End If

    End Sub

    Private Sub RadBtnNorw_CheckedChanged(sender As Object, e As EventArgs) Handles RadBtnNorw.CheckedChanged
        If RadBtnNorw.Checked Then
            gLang = "Nor"
        Else
            gLang = "Eng"
        End If
        LoadResourceStrings()
    End Sub

    Private Sub RdBtnEng_CheckedChanged(sender As Object, e As EventArgs) Handles RdBtnEng.CheckedChanged
        If RdBtnEng.Checked Then
            gLang = "Eng"
        Else
            gLang = "Nor"

        End If
        LoadResourceStrings()
    End Sub


End Class