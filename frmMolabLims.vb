'** ******************************************
'** Denne brukes ifm Molab2Lims
'** ***************************************

Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Management
'Imports System.Data '1.0.3.2
Imports Oracle.ManagedDataAccess.Client
'Imports Oracle.ManagedDataAccess.Types

Friend Class frmMain

    Inherits System.Windows.Forms.Form
    Friend Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Integer,
         ByVal nPriority As Integer) As Integer

    Friend Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Integer,
         ByVal dwPriorityClass As Integer) As Integer

    Friend Declare Function GetCurrentThread Lib "kernel32" () As Integer
    Friend Declare Function GetCurrentProcess Lib "kernel32" () As Integer

    Friend Const THREAD_PRIORITY_HIGHEST As Short = 2
    Friend Const HIGH_PRIORITY_CLASS As Integer = &H80
    Private lblTxt(18) As String '1.0.3.1,1.0.4.2
    '===========================================================================
    'FORM/TIMER EVENTS
    Structure tSampleId '1.0.0.0 Molab2Lims
        Dim Elms_no As String
        Dim UtSted As String
        Dim PrType As String
        Dim PrNr As String
        Dim SubNr As String
        Dim Tatt_Dato As String
        Dim Ankomst_Dato As String
        Dim Metode As String
        '1.0.0.5
        Dim Grade As String
        Dim Fraksjon As String
        Dim Kontainer As String
    End Structure

    Private Sub FrmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        'MsgBox "Kommandolinje: " & Command$
        Dim sCmd As String = LCase(VB.Command())
        '2.0.0.5 gNoErrorHandling = InStr(1, LCase(VB.Command()), "noerror") > 0
        gNoErrorHandling = sCmd.IndexOf("noerror") > 0
        If Not gNoErrorHandling Then On Error GoTo Err_Form_Load
        Dim i As Object
        'TBDJS
        '2.0.0.3
        'ProgressBar1.Value += 1
        SetThreadPriority(GetCurrentThread, THREAD_PRIORITY_HIGHEST)
        SetPriorityClass(GetCurrentProcess, HIGH_PRIORITY_CLASS)
        'ProgressBar1.Value += 1 '2.0.0.3

        Me.Text = APP_NAME
        Lib_Init()
        Call GetDesimalSeparator()
        System.Windows.Forms.Cursor.Current = Lib_GetArrow()
        gfrmMain = Me
        'Call Lib_ShowErrorBox(LOG_STATUS, "Form_Load", 2, "start", "")

        Call Lib_SetLanguage()
        '2.0.0Call LoadImages()

        '2.0.0 Me.Icon = Lib_GetFromImageList(101, VB6.LoadResConstants.ResIcon)
        ' cmdOnK1.Picture = Lib_GetFromImageList(101, vbResBitmap)
        '1.0.4.5 
        'forløpig må kanskje i bruk for å hente data Call GetConfig()
        '1.0.3.6
        Dim sFileName = My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".INI"

        '1.0.5.1 flyttet getServInfo for de som har server.dat får glede av den.
        If Not File.Exists(sFileName) Then
            Call GetServInfo()
            Call WriteIniFile()
        End If
        Call ReadIniFil()
        Call LoadResourceStrings()

        Me._sbStatusBar_Panel1.Text = lblTxt(0) 'venter på data
        Me._sbStatusBar_Panel2.Text = Now()

        'ChDir CurDir  ' ml 001107 used for test..
        '2.0.0.5 osv gNoConfirm = InStr(1, LCase(VB.Command()), "noconfirm") > 0
        'set i command line argument box ved behov :/quickstart/noconfirm/noerror
        gNoConfirm = sCmd.IndexOf("noconfirm") > 0
        gNoSave = sCmd.IndexOf("nosave") > 0
        Call SetStatusDetail(Lib_LoadResString(0, M_FILE_ID, "Starter..."))
        Me.Refresh()
        'Call Genericlog("Form_load", "====================================================")
        Call Genericlog("Form_load", "Molab2Lims startet")
        '1.0.5.2 gServer.dbCon = New ADODB.Connection
        'Next i
        lstLogFiles.Items.Clear()
        If sCmd.IndexOf("quickstart") > 0 Then
            gSecToTrans = 2
            gSecToRejTrans = 60
        Else
            gSecToTrans = 10
            gSecToRejTrans = 120
        End If
        gLastTimeStamp = System.DateTime.FromOADate(Now.ToOADate - 1) '010430 ml
        '--- (params)

        Call SetStatusDetail(Lib_LoadResString(0, M_FILE_ID, "Starter... Leser konfigurering..."))
        '1.0.3.1 Call GetConfig()
        Call ServList()
        Dim iQuent = QueCount(1)

        'ml 010214 logg til oracle, kan kommenteres bort under test
        Call UpdateScreen(True)
        gProcessing = False
        Call SetStatusDetail(lblTxt(7))
        'TEST:
        'Call CreRejFileRecordSet
        'start/ikke start logging ved oppstart
        Call SetStatus(True)
        Call StartUP() '1.0.3.2

        '1.0.0.0 Molab
        If gOrgUnit = "BRE" Then
            gUsedOraTransProg = "INSTR_ELMS2B"
        Else
            gUsedOraTransProg = "INSTR_ELMS2"
        End If

        '1.0.4.8
        Dim conn As New OracleConnection(gConnect)
        Call StoreMessage(conn, APP_NAME & " startet, " & gUsedOraTransProg & " brukes for overføring, Org unit : " & gOrgUnit)
        conn = Nothing
Exit_Form_Load:
        Exit Sub

Err_Form_Load:
        Call GenericError("Form_Load", "Feil i Form_Load-kode...", Err.Number, Err.Description, Err.Source)
        'Debug.Print(Err.Description)
        Resume Exit_Form_Load
    End Sub
    Private Sub btnRetry_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnRetry.Click
        '2.7.1 On Error GoTo ErrorL
        btnRetry.Enabled = False
        Timer1.Enabled = False '1.3.7
        Timer2.Enabled = True

        gBackUpFile = True
        Call InitRetrTrans()
        gBackUpFile = False
        'ml 010503 kjøre filene tilbake.
        Call RetRetFiles()
        Call RemoveOldFiles()
        Timer1.Enabled = True '1.3.7
        Timer2.Enabled = False
        gProcessing = False 'starte telling igjen

        btnRetry.Enabled = True
        Exit Sub
ErrorL:
        Call GenericError("frmMain", "Feil i btnRetry_Click...", Err.Number, Err.Description, Err.Source)
    End Sub

    Private Function Test_Mapping() As Short
        '1.0.1.5 6. mai 21
        'Test om mapping er aktiv
        Dim sDriver As String = String.Empty
        Dim sType As String = String.Empty
        Dim sPath As String = String.Empty
        Test_Mapping = 0
        For Each drive_info As DriveInfo In DriveInfo.GetDrives()
            sDriver = drive_info.Name
            sType = drive_info.DriveType().ToString
            sPath = gPath.Substring(0, 3)

            'If sPath = sDriver And sType = "Network" Then
            If sPath = sDriver Then
                Debug.Print("Yes " & sDriver)
                Debug.Print("sPath " & sPath)
                Debug.Print("S Type " & sType)
                Test_Mapping = 1
            End If
        Next drive_info

    End Function

    Private Function QueCount(ByRef iOption As Short) As Short
        '1.0.1.5 6. mai 21 Til function
        On Error GoTo ErrorL
        'ml 010503 for å spole tilbake mislykkede filer.
        Dim curr_file, lListItemText As String
        Dim lCurrTimeStamp As Date
        Dim i As Short
        Dim iSec As Object
        QueCount = 1
        '1.0.1.5 6. mai 21
        'Test om mapping er aktiv
        Dim m As Short
        m = Test_Mapping()
        If m = 0 Then
            'MsgBox("Main QueCount : Mapping er borte, finner ikke  " & gPath)
            QueCount = 0
        End If
        '
        curr_file = Dir(gPath & "\transfer\*.*", FileAttribute.Archive)
        i = 0
        Do While curr_file <> ""
            i = i + 1
            curr_file = Dir()
        Loop
        txtQueCount.Text = CStr(i)
        BBStat(3) = i
        curr_file = Dir(gPath & "\error\*.*", FileAttribute.Archive)
        i = 0
        Do While curr_file <> ""
            i = i + 1
            curr_file = Dir()
        Loop
        txtErrCount.Text = CStr(i)
        BBStat(2) = i

        iSec = Int(CDbl(Format(Now, "ss")))
        If iSec = 0 Or iSec = 5 Or iSec = 10 Or iSec = 15 Or iSec = 20 Or iSec = 25 Or iSec = 30 Or iSec = 35 Or iSec = 40 Or iSec = 45 Or iSec = 50 Or iSec = 55 Or iOption = 1 Then
            i = 0
            curr_file = Dir(gPath & "\rej\*.*", FileAttribute.Archive)
            lstRejFiles.Items.Clear()
            Do While curr_file <> ""
                i = i + 1
                lCurrTimeStamp = FileDateTime(gPath & "\rej\" & curr_file)
                lListItemText = (Format(Now, "dd.MM.yy hh.mm.ss") & "         Filnavn: " & curr_file & "         Fil lagret " & Format(lCurrTimeStamp, "dd.MM.yy hh:mm:ss"))
                lstRejFiles.Items.Insert(0, lListItemText)
                curr_file = Dir()
            Loop
            BBStat(4) = i
            i = 0
            curr_file = Dir(gPath & "\accept\*.*", FileAttribute.Archive)
            Do While curr_file <> ""
                i = i + 1
                curr_file = Dir()
            Loop
            BBStat(5) = i

        End If
        Exit Function
ErrorL:
        Call GenericError("FrmMain", "Feil i QueCount...", Err.Number, Err.Description, Err.Source)
    End Function
    Private Sub RetRetFiles()
        On Error GoTo ErrorL
        'ml 010503 for å spole tilbake mislykkede filer.
        Dim curr_file As String
        curr_file = Dir(gPath & "\re2\*.*", FileAttribute.Archive)
        Do While curr_file <> ""
            File.Move(gPath & "\re2\" & curr_file, gPath & "\rej\" & curr_file)
            curr_file = Dir()
        Loop
        Exit Sub
ErrorL:
        '1.7.1
        Call GenericError("FrmMain", "Feil i RetRetFiles...", Err.Number, Err.Description, Err.Source)
    End Sub
    Function GetRejStatus(ByRef pFile As String) As Short
        Dim rst As ADODB.Recordset
        Dim iFileCriteria As String
        On Error GoTo ErrorL
        rst = New ADODB.Recordset
        rst.Open(My.Application.Info.DirectoryPath & "\RejFilesStatus.dat")
        'MsgBox ("antall " & rst.RecordCount)
        iFileCriteria = "FileName='" & pFile & "'"
        rst.Filter = iFileCriteria
        Do Until rst.EOF
            GetRejStatus = rst.Fields("Retry").Value
            rst.MoveNext()
        Loop
        rst.Close()
        rst = Nothing
        Exit Function
ErrorL:
        Call GenericError("frmMain", "Feil i GetRejStatus-kode...", Err.Number, Err.Description, Err.Source)
    End Function
    Private Sub LoadResourceStrings()
        On Error GoTo ErrorL
        If gLang = "Nor" Then '1.0.3.1
            btnStart.Text = "Start ovrføring"
            btnStop.Text = "Stopp overføring"
            lblTxt(0) = "Venter på data"
            lblTxt(1) = "Er du sikker på at du vil AVSLUTTE loggeprogrammet?"
            lblTxt(2) = "Avslutte loggeprogram?"
            lblTxt(3) = "Overføring stoppet."
            lblTxt(4) = "Er du sikker på at du vil STOPPE overføring?"
            lblTxt(5) = "Stoppe overføring?"
            lblTxt(6) = "Stopp overføring før endring"
            lblTxt(7) = "Venter til neste overføring..."
            lblTxt(8) = "Ovrføring pågår"
            lblTxt(9) = "Overføring stanset"
            lblTxt(10) = "Åpner fil!"
            lblTxt(11) = "Leser data!"
            lblTxt(12) = "Read S messesage!"
            lblTxt(13) = "Prøve og verdi ferdig overført."
            lblTxt(14) = "Overføring startet."
            lblTxt(15) = "Stopp log og finn log filer i menyen : Verktøy!" '1.0.5.3
            lblTxt(16) = "Mangler metode!"
            lblTxt(17) = "Mapping borte!" '1.0.1.5 6. mai 21
            lblTxt(18) = "Venter på mapping.."
            Me.mnuTools.Text = "Verktøy"
            Me.mnuOption.Text = "Optioner..."
            lblErrCount.Text = "Antall i feil katalog"
            btnStopforDebug.Text = "Stopp for debug"
            btnRetry.Text = "Nytt forsøk på avviste filer"
            cmdRejFileStatus.Text = "Status avviste filer"
            Label9.Text = "Siste avviste filer fra Oracle"

            Label10.Text = "Siste logge filer laget/oppdatert"
            Label3.Text = "Antall sekunder til neste overføring"
            lblQueCount.Text = "Antall i behandlingskatalog"
            Label5.Text = "Status"
            Label9.Text = "Siste avviste filer til Oracle"
            Label1.Text = "Siste overførte filer"
            lblMinToNextRej.Text = "Min. til neste forsøk med avviste filer."
        ElseIf gLang = "Eng" Then '1.0.3.1
            btnStart.Text = "Start transfer"
            btnStop.Text = "Stop transfer"
            lblTxt(0) = "Waiting data"
            lblTxt(1) = "Sure you vil STOP the program?"
            lblTxt(2) = "Stop the program?"
            lblTxt(3) = "Transfering stopped."
            lblTxt(4) = "Sure you will STOP file transfering?"
            lblTxt(5) = "Stop file transfer?"
            lblTxt(6) = "Stop transfer before other options"
            lblTxt(7) = "Waiting for next file to transfer..."
            lblTxt(8) = "Transfering is on"
            lblTxt(9) = "Transfer stopped"
            lblTxt(10) = "Opens file!" '
            lblTxt(11) = "Read data!"
            lblTxt(12) = "Read S messesage!"
            lblTxt(13) = "Sample and values transfered."
            lblTxt(14) = "Transfered started."
            lblTxt(15) = "Stop log og find log files in the meny : Tools!" '1.0.5.3
            lblTxt(16) = "Methode missed!"

            lblTxt(17) = "Mapping missed!" '1.0.1.5 6. mai 21
            lblTxt(18) = "Wainting for reattached mapping.."

            Me.mnuTools.Text = "Tools"
            Me.mnuOption.Text = "Options..."
            Me.mnuAbout.Text = "About..."
            lblErrCount.Text = "Number in error catalog"
            btnStopforDebug.Text = "Stop for debug"
            btnRetry.Text = "New try with rejected files"
            cmdRejFileStatus.Text = "Status rejected files"
            Label9.Text = "Last rejected files from Oracle"
            Label10.Text = "Last log files saved/updated"
            Label3.Text = "Seconds to next transfer"
            lblQueCount.Text = "Number in transfer directory"
            Label5.Text = "Status"
            Label9.Text = "Last rejeced files to Oracle"
            Label1.Text = "Last transfered files"
            lblMinToNextRej.Text = "Min. to next try with rejected files"
        End If


        Exit Sub

ErrorL:
        Call GenericError("frmMain", "Feil i LoadResourceStrings-kode...", Err.Number, Err.Description, Err.Source)
    End Sub

    Private Sub btnStopforDebug_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnStopforDebug.Click
        Stop
    End Sub

    Private Sub StartUP()
        Dim DatoSjekk As Boolean
        Dim i As Short
        If Not gNoErrorHandling Then On Error GoTo Err_StartUP

        gConnect = "DATA SOURCE=" & gServer.ConStr & ";PASSWORD=" & gServer.Pwd & ";PERSIST SECURITY INFO=True;USER ID=" & gServer.User

Exit_StartUP:
        Exit Sub

Err_StartUP:
        Call GenericError("FrmMain", "Feil i StartUP kode...", Err.Number, Err.Description, Err.Source)
        gServer.Connected = False
        Resume Exit_StartUP
    End Sub



    Private Sub FrmMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        If gNoConfirm Then
            '1.0.4.5 Call SaveConfig()
            Call WriteIniFile() '1.0.3.6
            '1.0.4.8
            Dim conn As New OracleConnection(gConnect)
            Call StoreMessage(conn, APP_NAME & " stoppet, " & gUsedOraTransProg & " brukes for overføring, Org unit : " & gOrgUnit)
            conn = Nothing

            Call Genericlog("Form_unload", "Molab2Lims stoppet")
            Timer2.Enabled = False
            frmOptions.Close()
        Else
            Timer1.Enabled = False
            Timer2.Enabled = False
            'Call SaveConfig()
            Dim conn As New OracleConnection(gConnect)
            Call StoreMessage(conn, APP_NAME & " stoppet, " & gUsedOraTransProg & " brukes for overføring, Org unit : " & gOrgUnit)
            conn = Nothing

            Call WriteIniFile() '1.0.4.5
            Call Genericlog("Form_unload", "Molab2Lims stoppet", True)
            frmOptions.Close()
            'End If
        End If
    End Sub




    Public Sub mnuAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAbout.Click
        frmAbout.ShowDialog()

    End Sub

    Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click
        Me.Close()
    End Sub


    Public Sub mnuOption_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOption.Click

        If Not gActive Then
            Call ReadIniFil()
            frmOptions.StartTool()
        Else
            MsgBox(lblTxt(6))
        End If
    End Sub


    Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        Dim i As Short
        Dim iStr As String
        Dim TimeInter As Short
        Dim a As Object
        Dim iQuent As Short
        If Not gNoErrorHandling Then On Error GoTo Err_Timer1_Timer
        'i = Me.Left 1.4.2 dette årsak til timer kode går feil brukes ikke??
        'i = Me.Top
        iStr = Me.Name
        If gActive And Not gProcessing Then
            If gSecToTrans > 1 Then
                gSecToTrans = gSecToTrans - 1
                iQuent = QueCount(0)
                Me._sbStatusBar_Panel1.Text = lblTxt(0) '1.0.0.0c
                If iQuent = 0 Then '1.0.1.5 6. mai 21 - map sjekk 
                    gSecToTrans = gSecBetweenTrans
                    Me._sbStatusBar_Panel1.Text = lblTxt(17)
                    Call SetStatusDetail(lblTxt(18))

                End If
                Dim conn As New OracleConnection(gConnect) '1.0.3.2
                    If conn.State Then
                        conn.Close()
                    End If
                    gServer.Connected = False

                Else
                    'Tester baseforbindelsen
                    iQuent = QueCount(1)
                gSecToTrans = gSecBetweenTrans
                Call InitTrans()
            End If
        End If
        'Application.DoEvents() '1.0.4.1
        Call UpdateScreen(True)
        If gServer.Connected Then
            txtServer.BackColor = Color.GreenYellow

        Else
            txtServer.BackColor = Color.WhiteSmoke
        End If
        If gAutoRetry Then
            If gSecToRejTrans >= 1 Then
                gSecToRejTrans = gSecToRejTrans - 1
            Else
                gSecToRejTrans = gMinBetweenRejTrans * 60
                Call btnRetry_Click(btnRetry, New System.EventArgs())
            End If
        End If

Exit_Timer1_Timer:
        Exit Sub

Err_Timer1_Timer:
        Call GenericError("FrmMain", "Feil i timer-kode...", Err.Number, Err.Description, Err.Source)
        Resume Exit_Timer1_Timer
    End Sub

    '===========================================================================
    'BUTTON EVENTS

    Private Sub btnStart_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnStart.Click
        gProcessing = False 'ml 010503
        Timer1.Enabled = True '1.3.4
        SetStatus((True))
        '1.0.4.8
        Dim conn As New OracleConnection(gConnect)
        Me._sbStatusBar_Panel1.Text = lblTxt(14) '1.0.5.0
        Call StoreMessage(conn, APP_NAME & " logging startet.")
        conn = Nothing
        Call Genericlog("btnStart_Click", "Logging startet.", True)
        'Call Lib_AppLog(mLib.enuLOG_TYPE.LOG_TYPE_STATUS, mLib.enuPRIORITY.PRIORITY_LOW, mLib.enuSHOW_TYPE.SHOW_TYPE_NONE, "frmMain", "btnStart_Click", 0, "Logging startet", "Time enabled")
    End Sub

    Private Sub btnStop_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnStop.Click
        Try
            If gNoConfirm Then
                SetStatus((False))
                '1.0.3.2
                Dim conn As New OracleConnection(gConnect) '1.0.3.3
                If conn.State Then
                    conn.Close()
                End If
                conn = Nothing

                gServer.Connected = False
            Else
                If MsgBox(lblTxt(4), MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation + MsgBoxStyle.DefaultButton2, lblTxt(5)) = MsgBoxResult.Yes Then
                    SetStatus((False))
                    Dim conn As New OracleConnection(gConnect)
                    Call StoreMessage(conn, APP_NAME & " logging stoppet") '1.0.4.8
                    If conn.State Then
                        conn.Close()
                    End If
                    conn = Nothing
                    gServer.Connected = False
                    txtServer.BackColor = Color.White '1.1.1.1
                    '1.4.2 flyttet denne inn i løkke.
                    Me._sbStatusBar_Panel1.Text = lblTxt(3) '1.0.0.0e
                    Call Genericlog("btnStop_Click", "Logging stoppet", True)
                End If
            End If
        Catch Ex As Exception
            txtServer.BackColor = Color.LightYellow '1.1.1.1
            Call GenericError("cmdOnOff_Click", "Feil i cmdOnOff_Click-kode..", Err.Number, Err.Description, Err.Source)
            gServer.Connected = False
            MessageBox.Show(Ex.Message)
        End Try
    End Sub


    '===========================================================================
    'LOG PROC/FUNC

    Private Sub InitTrans()
        Dim lPath As String
        If Not gNoErrorHandling Then On Error GoTo Err_InitTrans
        lPath = gPath & "\transfer"
        gBackUpFile = False
        gProcessing = True
        gLastTimeStamp = Now
        Call ProcessNewFiles(lPath, gLastTimeStamp)
        'Call ProcessNewFiles(lPath)
        gProcessing = False

        Call UpdateScreen(True)

Exit_InitTrans:
        gProcessing = False
        Exit Sub

Err_InitTrans:
        Call GenericError("FrmMain", "Feil i InitTrans...", Err.Number, Err.Description, Err.Source)
        Resume Exit_InitTrans

    End Sub

    Sub InitRetrTrans()
        'ml 001109 new for handle rejected files.

        Dim lPath As String
        If Not gNoErrorHandling Then On Error GoTo Err_InitRetrTrans
        lPath = gPath & "\rej"
        gProcessing = True
        Call ProcessNewFiles(lPath, gLastTimeStamp) 'ml 010430
        gProcessing = False
        'gLastTimeStamp = ProcessNewFiles(lPath)
        Call UpdateScreen(True)

Exit_InitRetrTrans:
        gProcessing = False
        Exit Sub

Err_InitRetrTrans:
        Call GenericError("FrmMain", "Feil i InitRetrTrans...", Err.Number, Err.Description, Err.Source)
        Resume Exit_InitRetrTrans
    End Sub

    Public Sub ProcessNewFiles(ByRef pPath As String, Optional ByRef pLastTimeStamp As Date = #12:00:00 AM#)
        If Not gNoErrorHandling Then On Error GoTo Err_ProcessNewFiles
        Dim NoOfFiles As Short
        Dim curr_file As String
        Dim prev_file As String
        Dim lCurrTimeStamp As Date
        Dim lTempTimeStamp As Date
        Dim lLastTimeStamp As Date
        Dim lListItemText As String
        Dim FailedToConnect As Boolean
        Dim lConnected As Boolean
        Dim lElmsNo As Integer
        Dim lVisc As Single
        Dim i, X As Short
        Dim iMess As String
        Dim lsPath As String
        'Dim fs As Object '1.6.8

        '2.0.0.3
        SetThreadPriority(GetCurrentThread, THREAD_PRIORITY_HIGHEST)
        SetPriorityClass(GetCurrentProcess, HIGH_PRIORITY_CLASS)

        lLastTimeStamp = System.DateTime.FromOADate(pLastTimeStamp.ToOADate + System.DateTime.FromOADate(0.001).ToOADate)
        NoOfFiles = 0
        Call SetStatusDetail(Lib_LoadResString(0, M_FILE_ID, "Leter etter nye filer"))
        FailedToConnect = False
        lTempTimeStamp = Now
        curr_file = Dir(pPath & "\*.*", FileAttribute.Archive)
        Do While curr_file <> "" And lLastTimeStamp > lTempTimeStamp 'ml 010502
            lTempTimeStamp = Now
            If curr_file <> "." And curr_file <> ".." And Not IgnoreFile(curr_file) Then
                If (GetAttr(pPath & "\" & curr_file) And FileAttribute.Directory) <> FileAttribute.Directory Then
                    'MsgBox curr_file
                    lCurrTimeStamp = FileDateTime(pPath & "\" & curr_file)
                    '1.4.0 FOR Å HA KLAR STATUS PÅ FILENE
                    gServer.TransFlag = 0
                    NoOfFiles = NoOfFiles + 1
                    Call SetStatusDetail(Lib_LoadResString(0, M_FILE_ID, "Behandler datafil " & curr_file))
                    SetAttr(pPath & "\" & curr_file, FileAttribute.Normal)
                    lListItemText = (Format(Now, "dd.MM.yy hh.mm.ss") & "         Filnavn: " & curr_file & "         Fil lagret " & Format(lCurrTimeStamp, "dd.MM.yy hh:mm:ss"))
                    'Her starter jobben 
                    'bort med disse : 1.0.3.2 Application.DoEvents() '2.0.3
                    If ReadAndLogFile(pPath, curr_file, lCurrTimeStamp, lElmsNo) Then
                        '1.6.7 Her er allerede prøven overført.
                        If gNotElmsSample And gNotElms.FileMove Then
                            lsPath = gNotElms.Path '1.6.9
                            If Not Directory.Exists(lsPath) Then
                                lsPath = gPath & "\WilabError OBS!!!"
                            End If
                            lListItemText = (Format(Now, "dd.MM.yy hh.mm.ss") & " Ikke Elms fil : " & curr_file & " flyttet til " & lsPath)
                        ElseIf gNotElmsSample And Not gNotElms.Use Then
                            lListItemText = (Format(Now, "dd.MM.yy hh.mm.ss") & " Ikke Elms fil : " & curr_file & " strøket ")
                        End If
                        lstLogFiles.Items.Insert(0, lListItemText)
                        Call ProcessingLog(True, lListItemText) '1.3.1

                        'flytter filene
                        Call BackupFile(pPath, curr_file, True)
                        'Call Lib_LogAppError(LOG_TYPE_STATUS, "ProcessNewFiles", 1, curr_file, "lagret på accept") '1.3.3
                    Else
                        lstRejFiles.Items.Insert(0, lListItemText)
                        Call ProcessingLog(False, lListItemText) '1.3.1
                        Call BackupFile(pPath, curr_file, False)
                        System.Windows.Forms.Application.DoEvents()
                        System.Threading.Thread.Sleep(5000) '1.0.5.3 wait, pause 
                    End If
                End If
            End If
            curr_file = Dir(pPath & "\*.*", FileAttribute.Archive) '1.4.0
        Loop
        '1.4.5
        X = lstLogFiles.Items.Count - 1
        If X > 20 Then
            For i = X To 15 Step -1
                lstLogFiles.Items.RemoveAt(i)
            Next i
        End If
        X = lstRejFiles.Items.Count - 1
        If X > 15 Then
            For i = X To 10 Step -1
                lstRejFiles.Items.RemoveAt(i)
            Next i
        End If
        Call SetStatusDetail(lblTxt(7))
        'ProcessNewFiles = lTempTimeStamp

Exit_ProcessNewFiles:
        Exit Sub

Err_ProcessNewFiles:
        Call GenericError("FrmMain", "Feil i ProcessNewFiles...", Err.Number, Err.Description, Err.Source)
        Resume Exit_ProcessNewFiles

    End Sub
    'Private Function CheckFile(ByVal filename As String) As Boolean

    '    Try
    '        System.IO.File.Open(filename, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.None)
    '        FileClose(1)
    '        Return False
    '    Catch ex As Exception
    '        Return True
    '    End Try

    'End Function

    Public Function ReadAndLogFile(ByRef pPath As String, ByRef pFile As String, ByRef pTakenDate As Date, ByRef pElmsNo As Integer) As Boolean
        'gNoErrorHandling = True 'bugTest 1.0.5
        If Not gNoErrorHandling Then On Error GoTo Err_ReadAndLogFile
        'Returing true if readine and logging was ok
        Dim lLine As String
        Dim lBlanc As String = String.Empty
        Dim lBit(6) As String
        Dim X As Short
        Dim iPos(2) As Short
        Dim lMessType As String
        Dim lComment As String
        Dim lOk As Object
        Dim lLogOk As Boolean
        Dim lInsSampOk As Boolean
        Dim lInsResOk As Boolean
        Dim i, iLenSMess As Short
        Dim lDMess As String = String.Empty
        Dim lSMess As String
        Dim lCmess As String = String.Empty
        Dim lNCMess, lNRealCmess As Short
        Dim lArlMess As String
        Dim lsDate As String '1.0.0.0 Molab2Lims
        Dim lsubSMess As String
        Dim lStatusMess, lKanal As String
        Dim lMethCode As String = String.Empty
        Dim lCompound, lStatusMess2 As String
        Dim lValue As Double
        Dim lUnit, lNumString, lParInfo As String
        Dim iPosCount As Short '1.6.3
        Dim iFl As Short '1.6.7
        Dim sStatMess2 As String '2.0.1.5
        Dim conn As New OracleConnection(gConnect) '1.0.3.4
        Dim lFilenumber As Short '1.0.3.8
        Dim bArlElmsno As Boolean '1.0.5.4 elmsnr
        Dim sStep As String = String.Empty

        '-- GetTimeStamp, substract a certain (from registry) delay and asign lTakenDate
        '-- ml 010213 skrives om for å lese meldingsfiler.

        lMessType = ""
        lComment = "Molab2Lims: Filename=" & pFile & " "
        'lStatusMess = "Reading file " & pPath & "\" & pFile & " from " & Format(pTakenDate, "dd.mm.yyyy hh:mm:ss")
        Me._sbStatusBar_Panel1.Text = lblTxt(10)

        lFilenumber = FreeFile()
        FileOpen(lFilenumber, pPath & "\" & pFile, OpenMode.Input)
        lNCMess = 0
        lSMess = ""
        lNRealCmess = 0
        lMessType = ""
        gSource = ""
        lNCMess = 0
        gNotElmsSample = False
        'For i = 1 To gServCount
        gServer.NCMess = 0
        gServer.NRealCmess = 0
        bArlElmsno = False '1.0.5.4 elmsnr
        'Next i
        '** **********************************
        '** Denne sekvensen skrives om ifm Molab2Lims
        '** **************************
        Dim vSampleId As tSampleId
        vSampleId.Elms_no = "         "
        vSampleId.UtSted = ""
        vSampleId.PrType = ""
        vSampleId.SubNr = ""
        vSampleId.Tatt_Dato = ""
        vSampleId.PrNr = ""
        vSampleId.Metode = ""
        vSampleId.Ankomst_Dato = ""
        vSampleId.Grade = ""
        vSampleId.Kontainer = ""
        vSampleId.Fraksjon = ""
        Dim iStartPosId As Integer = 25
        lBlanc = "                                                                   " '1.0.6.0 flyttet opp
        sStep = "1"
        Me._sbStatusBar_Panel1.Text = lblTxt(11) ' "Leser data!"
        Application.DoEvents()
        Do While Not EOF(lFilenumber)
            lLine = LineInput(lFilenumber)
            'Debug.Print lLine
            If Len(lLine) < 2 Then
                On Error Resume Next '1.5.7
                lLine = LineInput(1)
                If Not gNoErrorHandling Then On Error GoTo Err_ReadAndLogFile
                'Exit Do
            End If
            '1.5.0 Arl
            sStep = "2"
            If VB.Left(lLine, 6) & "" = "SOURCE" Then
                Me._sbStatusBar_Panel1.Text = lblTxt(12)
                gSource = RTrim(Mid(lLine, iStartPosId, Len(lLine) - iStartPosId))
            ElseIf VB.Left(lLine, 7) & "" = "ELMS_NO" Then
                lMessType = "S"
                vSampleId.Elms_no = Mid(lLine, iStartPosId, 9)
                gNotElms.FileMove = False

                If gNotElms.Use Then ' 1.0.4.7
                    If Mid(lSMess, gNotElms.Pos, gNotElms.len_Renamed) = gNotElms.txt Then
                        gNotElms.FileMove = True
                    End If
                    'Next iFl
                    If gNotElms.Use And gNotElms.FileMove Then
                        gNotElmsSample = True
                    End If
                End If
            ElseIf VB.Left(lLine, 11) & "" = "ORIGIN_CODE" Then
                vSampleId.UtSted = RTrim(Mid(lLine, iStartPosId, Len(lLine) - iStartPosId + 1))
            ElseIf VB.Left(lLine, 12) & "" = "SAMP_NO_CODE" Then
                vSampleId.PrType = RTrim(Mid(lLine, iStartPosId, Len(lLine) - iStartPosId + 1))
            ElseIf VB.Left(lLine, 7) & "" = "SAMP_NO" Then
                vSampleId.PrNr = RTrim(Mid(lLine, iStartPosId, Len(lLine) - iStartPosId + 1))
            ElseIf VB.Left(lLine, 6) & "" = "SUB_NO" Then
                vSampleId.SubNr = RTrim(Mid(lLine, iStartPosId, Len(lLine) - iStartPosId + 1))
            ElseIf VB.Left(lLine, 10) & "" = "TAKEN_DATE" Then
                vSampleId.Tatt_Dato = RTrim(Mid(lLine, iStartPosId, Len(lLine) - iStartPosId + 1))
            ElseIf VB.Left(lLine, 11) & "" = "APPLICATION" Then
                vSampleId.Metode = RTrim(Mid(lLine, iStartPosId, Len(lLine) - iStartPosId + 1))
            ElseIf VB.Left(lLine, 16) & "" = "MEASUREMENT_TIME" Then
                'Kjøretidspunkt
                Dim lsDate2 As String = String.Empty
                lsDate2 = RTrim(Mid(lLine, iStartPosId, Len(lLine) - iStartPosId + 1))
                lsDate = Mid(lsDate2, 9, 2)
                lsDate = lsDate & Mid(lsDate2, 4, 2)
                lsDate = lsDate & Mid(lsDate2, 1, 2)
                lsDate = lsDate & Mid(lsDate2, 12, 2)
                lsDate = lsDate & Mid(lsDate2, 15, 2)
                vSampleId.Ankomst_Dato = lsDate
                lMessType = "D"
            ElseIf VB.Left(lLine, 5) & "" = "GRADE" Then
                vSampleId.Grade = RTrim(Mid(lLine, iStartPosId, Len(lLine) - iStartPosId + 1))
            ElseIf VB.Left(lLine, 8) & "" = "FRACTION" Then
                vSampleId.Fraksjon = RTrim(Mid(lLine, iStartPosId, Len(lLine) - iStartPosId + 1))
            ElseIf VB.Left(lLine, 12) & "" = "CONTAINER_NO" Then
                vSampleId.Kontainer = RTrim(Mid(lLine, iStartPosId, Len(lLine) - iStartPosId + 1))
                sStep = "2d"
            ElseIf lMessType = "C" Then
                If lLine.Length >= 50 Then
                    lCmess = lLine
                End If
                sStep = "2c"
            End If
            If lMessType = "D" Then
                'If Mid(lDMess, 3, 2) <> "  " Then '1.0.6.0
                'lNRealCmess = CShort(Mid(lDMess, 3, 2))
                'End If
                Dim IdentLen As Integer = 0
                Dim lMethLength As Integer = CInt(gMethLength)
                'lMethCode = lDMess
                iLenSMess = Len(lSMess)
                'Debug.Print lSMess
                Dim iLenPos As Short = 0
                Dim lsBlString As String = "                       "
                Dim lsMString As String = String.Empty
                iLenPos = 53
                lSMess = vSampleId.Elms_no
                If conn.State Then
                    gServer.Connected = True
                    txtServer.BackColor = Color.GreenYellow
                    '1.0.4.0 System.Windows.Forms.Application.DoEvents()
                Else
                    conn.Open()
                    gServer.Connected = True
                    txtServer.BackColor = Color.GreenYellow
                    '1.0.4.0 System.Windows.Forms.Application.DoEvents()
                End If
                '1.0.0.4c nødvendig initialisering
                giPrevToPos = 0
                giPrevDiff = 1

                sStep = "2d_1"
                '1.0.1.2 Sjekk om metoden finnes.
                If MethCodeExist(conn, vSampleId.Metode) Then
                    Me._sbStatusBar_Panel1.Text = "Fant metoden : " & vSampleId.Metode & "."
                    IdentLen = GetIdentLen(conn, vSampleId.Metode, "ORIGIN_CODE")
                    'If giPrevDiff = 2 Then lsMString = " "
                    lsMString = lsBlString.Substring(1, giPrevDiff - 1)
                    lSMess = lSMess & lsMString & vSampleId.UtSted & lBlanc.Substring(1, IdentLen - Len(vSampleId.UtSted))
                    IdentLen = GetIdentLen(conn, vSampleId.Metode, "SAMP_NO_CODE")
                    lsMString = lsBlString.Substring(1, giPrevDiff - 1)
                    lSMess = lSMess & lsMString & vSampleId.PrType & lBlanc.Substring(1, IdentLen - Len(vSampleId.PrType))
                    IdentLen = GetIdentLen(conn, vSampleId.Metode, "SAMP_NO")
                    lsMString = lsBlString.Substring(1, giPrevDiff - 1)
                    lSMess = lSMess & lsMString & vSampleId.PrNr & lBlanc.Substring(1, IdentLen - Len(vSampleId.PrNr))
                    IdentLen = GetIdentLen(conn, vSampleId.Metode, "SUB_NO")
                    lsMString = lsBlString.Substring(1, giPrevDiff - 1)
                    lSMess = lSMess & lsMString & vSampleId.SubNr & lBlanc.Substring(1, IdentLen - Len(vSampleId.SubNr))
                    IdentLen = GetIdentLen(conn, vSampleId.Metode, "TAKEN_DATE")
                    lsMString = lsBlString.Substring(1, giPrevDiff - 1)
                    lSMess = lSMess & lsMString & vSampleId.Tatt_Dato & lBlanc.Substring(1, IdentLen - Len(vSampleId.Tatt_Dato))
                    IdentLen = GetIdentLen(conn, vSampleId.Metode, "ARRIVED_DATE")
                    lsMString = lsBlString.Substring(1, giPrevDiff - 1)
                    lSMess = lSMess & lsMString & vSampleId.Ankomst_Dato & lBlanc.Substring(1, IdentLen - Len(vSampleId.Ankomst_Dato))
                    '******
                    '**    1.0.0.5 Nye kollonner: GRADE (2*3) ,FRACTION (16) og CONTAINER_NO (13) -> GRADE_CODE (52-57,FRACTION,CONTAINER_NO
                    '*****
                    IdentLen = GetIdentLen(conn, vSampleId.Metode, "FAMILY_CODE")
                    IdentLen = IdentLen + GetIdentLen(conn, vSampleId.Metode, "PRODUCT_CODE")
                    IdentLen = IdentLen + GetIdentLen(conn, vSampleId.Metode, "GRADE_CODE")
                    lsMString = lsBlString.Substring(1, giPrevDiff - 1)
                    lSMess = lSMess & lsMString & vSampleId.Grade & lBlanc.Substring(1, IdentLen - Len(vSampleId.Grade))
                    IdentLen = GetIdentLen(conn, vSampleId.Metode, "FRACTION")
                    lsMString = lsBlString.Substring(1, giPrevDiff - 1)
                    lSMess = lSMess & lsMString & vSampleId.Fraksjon & lBlanc.Substring(1, IdentLen - Len(vSampleId.Fraksjon))
                    IdentLen = GetIdentLen(conn, vSampleId.Metode, "CONTAINER_NO")
                    lsMString = lsBlString.Substring(1, giPrevDiff - 1)
                    lSMess = lSMess & lsMString & vSampleId.Kontainer & lBlanc.Substring(1, IdentLen - Len(vSampleId.Kontainer))
                    gRejServN = 0
                    'Lagring av prøve
                    sStep = "2d_2"
                    If Not gNotElmsSample Then
                        Call StartUP() '1.6.7
                        If conn.State Then
                            gServer.Connected = True
                            txtServer.BackColor = Color.GreenYellow
                            '1.0.4.0 System.Windows.Forms.Application.DoEvents()
                        Else
                            conn.Open()
                            gServer.Connected = True
                            txtServer.BackColor = Color.GreenYellow
                            '1.0.4.0 System.Windows.Forms.Application.DoEvents()
                        End If
                        'Debug.Print(lSMess & vbNewLine & vSampleId.Metode)
                        Me._sbStatusBar_Panel1.Text = "Save sample!"
                        lInsSampOk = StoreSamp(conn, vSampleId.Metode, lSMess, gOrgUnit)
                        If Not lInsSampOk Then
                            Me._sbStatusBar_Panel1.Text = lblTxt(15)
                        End If
                        '1.7.0 sørge for at fila flyttes til error
                        'Debug.Print("StoreSamp : " & vbNewLine & Err.Number & vbNewLine & Err.Description & Err.Source)
                    End If
                Else
                    gServer.TransFlag = 5
                    Call Genericlog("ReadAndLogFile", "Metoden : " & vSampleId.Metode & " finnes ikke.", True)
                    Me._sbStatusBar_Panel1.Text = lblTxt(15)

                End If
                sStep = "2d_3"
                '1.0.1.3  ElseIf lMessType = "C" And gServer.TransFlag < 5 Then
            ElseIf lMessType = "C" And gServer.TransFlag < 4 Then
                sStep = "2c_1"
                lNumString = ""
                lKanal = ""
                lParInfo = ""
                Dim iStartPosUnit As Integer = 50
                Dim iStartPosInfo As Integer = 45
                lKanal = RTrim(VB.Left(lCmess, iStartPosId - 1))
                'Debug.Print(Len(lCmess) & " og " & iStartPosUnit)
                lNumString = RTrim(Mid(lCmess, iStartPosId, iStartPosInfo - iStartPosId - 1))
                lParInfo = RTrim(Mid(lCmess, iStartPosInfo, iStartPosUnit - iStartPosInfo))
                lUnit = RTrim(Mid(lCmess, iStartPosUnit, Len(lCmess) + 1 - iStartPosUnit))
                lValue = 0
                '2.0.0.3
                'MsgBox("Hva skjer her : gDesimal= " & gDesimal & " for verdi " & lNumString & " på kanal " & lKanal & "?")
                If gDesimal = "," Then '2.0.1.0 2.0.1.1 ikke nødvendig?
                    lNumString = ModifyDecimalDelimiter(lNumString)
                End If
                sStep = "2c_2"
                If lNumString <> "" Then '2.0.1.2
                    lValue = CDbl(lNumString)

                    '2.0.0.3 lOk = Lib_ConvertStringToDouble(lNumString, lValue) '1.0.5
                    'If lOk And Not gNotElmsSample Then
                    If Not gNotElmsSample Then
                        'Debug.Print(Len(lCmess))
                        '1.0.1.1
                        Dim Encw1252 As System.Text.Encoding = System.Text.Encoding.GetEncoding("windows-1252")
                        Dim EncUTF8 As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")
                        Dim Str As String
                        Str = lKanal
                        Str = Encw1252.GetString(System.Text.Encoding.Convert(EncUTF8, Encw1252, System.Text.Encoding.Default.GetBytes(Str)))
                        If lKanal <> Str Then
                            lKanal = Str
                        End If
                        lCompound = lKanal
                        Me._sbStatusBar_Panel1.Text = "Save analysus value for :  " & lKanal
                        lInsResOk = StoreRes(conn, lKanal, lValue, lCompound, lUnit, lParInfo)
                    End If
                End If
                'End If
                sStep = "2c_3"
                If Err.Number <> 0 Then
                    Debug.Print("Evt.vb error : " & Err.Number & vbNewLine & Err.Description & Err.Source)
                    sStep = "2c_4"
                End If
            End If
            '1.7.2 If Not lInsSampOk Then
            '  GoTo Err_ReadAndLogFile:
            'End If
            If lMessType = "D" Then
                lMessType = "C"
            End If

        Loop
        System.Windows.Forms.Application.DoEvents()
        Me._sbStatusBar_Panel1.Text = lblTxt(13)
        sStep = "2e"

        '1.0.3.4
        System.Threading.Thread.Sleep(1000)
        System.Windows.Forms.Application.DoEvents()
        If conn.State Then
            conn.Close()
            gServer.Connected = False
            txtServer.BackColor = Color.White
        End If
        '***
        'Logging to Oracle
        '---
        sStep = "2f"
        lLogOk = True
        ReadAndLogFile = lLogOk '1.3.6
        lStatusMess = "Ferdig:  Par. på metode:" & lMethCode & ": " & lNRealCmess & ",par. lagret :" & lNCMess
        lStatusMess2 = ""
        sStatMess2 = ""
        'For i = 1 To gServCount
        ''0=skal ikke overføres,1=skal,2=prøve overført,3=verdi overført
        System.Windows.Forms.Application.DoEvents()
        If gServer.TransFlag = 0 Then '2.0.1.5
            sStatMess2 = "Skal ikke overføres" '
        ElseIf gServer.TransFlag = 1 Then
            sStatMess2 = "Skal ikke overføres, og ikke gjort." '
        ElseIf gServer.TransFlag = 2 Then
            sStatMess2 = "Prøve er overført, men ikke verdier." '
        ElseIf gServer.TransFlag = 3 Then
            sStatMess2 = "Prøve og verdier overført." '
        ElseIf gServer.TransFlag = 4 Then
            sStatMess2 = "Feil i overføring." '
            ReadAndLogFile = False '1.0.1.3
        ElseIf gServer.TransFlag = 5 Then '1.0.1.2 
            sStatMess2 = "Metode mangler." '
            Me._sbStatusBar_Panel1.Text = lblTxt(16)
            ReadAndLogFile = False
        End If
        sStep = "2g"
        lStatusMess2 = lStatusMess2 & " Server: " & i & " : " & sStatMess2
        'Next i
        lStatusMess2 = lSMess & lStatusMess2 & ",step :" & sStep
        Call Genericlog("ReadAndLogFile", lStatusMess2, True)


Exit_ReadAndLogFile:
        On Error Resume Next
        FileClose(lFilenumber) '1.0.3.8
        'gNoErrorHandling = False 'bugTest
        Exit Function

Err_ReadAndLogFile:
        Me._sbStatusBar_Panel1.Text = "Error. See log on " & My.Application.Info.DirectoryPath & "\error"
        'Application.DoEvents()
        'Debug.Print(lSMess & " " & lLine)
        '1.0.1.2 MessageBox.Show(Err.Description)
        Call GenericError("FrmMain", "Feil i ReadAndLogFile, step :" & sStep, Err.Number, Err.Description, Err.Source)
        Resume Exit_ReadAndLogFile
    End Function

    Private Sub BackupFile(ByRef pFromPath As String, ByRef pFileName As String, ByRef pOk As Boolean)
        If Not gNoErrorHandling Then On Error GoTo Err_BackupFile

        Dim fs As Object
        Dim iMess As String
        Dim i As Short
        Dim lPath As String = String.Empty
        Dim lFileName As String
        Dim lCurrTimeStamp As Date
        fs = CreateObject("Scripting.FileSystemObject")
        If pOk Then
            lPath = gPath & "\accept"
        ElseIf VB.Right(pFromPath, 3) <> "rej" Then
            lPath = gPath & "\rej"
        ElseIf VB.Right(pFromPath, 3) = "rej" Then
            If GetRejStatus(pFileName) < gCountRetry Then
                lPath = gPath & "\re2"
            Else
                lPath = gPath & "\error"
            End If
        End If

        If Not Directory.Exists(lPath) Then
            My.Computer.FileSystem.CreateDirectory(lPath)
        End If
        '1.3.6
        If InStr(1, pFileName, "rtgdat", CompareMethod.Text) > 0 Then
            lFileName = "R" & Format(MakeUId, "#0") & ".dat"
        Else
            lFileName = pFileName
        End If
        lCurrTimeStamp = FileDateTime(pFromPath & "\" & pFileName)
        '1.4.3 sjekke om fil finnes før flytting
        Call Genericlog("FrmMain", "BackupFile", True)
        If fs.fileexists(lPath & "\" & lFileName) Then
            fs.deletefile(lPath & "\" & lFileName)
        End If
        '1.6.4 1.6.7
        If Not gNotElmsSample Then
            fs.MoveFile(pFromPath & "\" & pFileName, lPath & "\" & lFileName)
        Else
            '1.6.7
            If gNotElmsSample And gNotElms.FileMove Then
                lPath = gNotElms.Path '1.6.8
                If Not Directory.Exists(lPath) Then
                    lPath = gPath & "\WilabError"
                End If
                If fs.fileexists(lPath & "\" & lFileName) Then
                    fs.deletefile(lPath & "\" & lFileName)
                End If
                fs.MoveFile(pFromPath & "\" & pFileName, lPath & "\" & lFileName)
            End If

        End If
        fs = Nothing


Exit_BackupFile:
        Exit Sub

Err_BackupFile:
        '1.5.9
        Call GenericError("FrmMain", "Feil i BackupFile...", Err.Number, Err.Description, Err.Source)
        Resume Exit_BackupFile
    End Sub



    '===========================================================================
    'MISC. PROC/FUNC

    '1.0.4.5 --- Legg på parameters...
    Private Sub GetConfig()
        '1.0.4.6 blir ikke brukt lenger.
        Dim lPwd As Object
        Dim sProgSetting As String '1.0.0.0

        If Not gNoErrorHandling Then On Error GoTo Err_GetConfig
        'sProgSetting = My.Application.Info.AssemblyName
        sProgSetting = "XrfTrans"
        'Call ReadIniFil() '1.0.4.5 forløpig 

        gNotElms.Use = CBool(GetSetting(sProgSetting, "Settings", "WilabUse", CStr(False)))
        gNotElms.Path = GetSetting(sProgSetting, "Settings", "WilabPath", "")
        gNotElms.Pos = CShort(GetSetting(sProgSetting, "Settings", "WilabPos", CStr(5)))
        gNotElms.len_Renamed = GetSetting(sProgSetting, "Settings", "WilabLen", CStr(1))
        gNotElms.txt = GetSetting(sProgSetting, "Settings", "WilabTekst", "-")

        gServer.User = GetSetting(sProgSetting, "Login", "User", "X40") 'ok
        gServer.ConStr = GetSetting(sProgSetting, "Login", "ElvisServer", "eits.nnn.elkem") 'ok
        frmOptions.txtServ.Text = gServer.ConStr
        frmOptions.txtUser.Text = gServer.User
        'End If
        gSecBetweenTrans = CSng(GetSetting(sProgSetting, "Settings", "SecBetweenTrans", CStr(10))) 'nei
        gMinBetweenRejTrans = CSng(GetSetting(sProgSetting, "Settings", "MinBetweenRejTrans", CStr(30))) 'nei
        gCountRetry = CSng(GetSetting(sProgSetting, "Settings", "CountRetry", CStr(2))) 'nei
        gOldFiles = CSng(GetSetting(sProgSetting, "Settings", "OldFiles", CStr(60))) 'nei
        gAutoRetry = GetSetting(sProgSetting, "Settings", "AutoRetry", "N") = "Y" 'nei
        gOrgUnit = GetSetting(sProgSetting, "Login", "OrgUnit", "ERR") 'ok
        frmOptions.txtOrgUnit.Text = gOrgUnit
        If (InStr(1, LCase(VB.Command()), "path") > 0) Then
            gPath = Mid(VB.Command(), InStr(1, LCase(VB.Command()), "path") + 5, InStr(InStr(1, LCase(VB.Command()), "path"), LCase(VB.Command() & " "), " ") - (InStr(1, LCase(VB.Command()), "path") + 5))
        Else
            gPath = GetSetting(sProgSetting, "Settings", "Path", "") 'nei
        End If
        If gPath = "" Then
            gPath = GetSetting("X40_Receive", "Settings", "Path", "C:\Program Files\XrfTrans\Data")
        End If
        frmOptions.txtPath.Text = gPath
        gBBPath = GetSetting(sProgSetting, "Settings", "BBPath", "") 'no ikke nødvenig.  For Big Brother
        frmOptions.txtBBPath.Text = gBBPath
        gLastTimeStamp = CDate(GetSetting(sProgSetting, "Settings", "LastTimeStamp", CStr(Now))) 'Default? no
        gFullSQLLog = GetSetting(sProgSetting, "Settings", "FullSqlLog", "N") = "Y" 'no
        gLogToFile = GetSetting(sProgSetting, "Settings", "LogToFile", "N") = "Y" 'no
        Call Genericlog("GetConfig", "SecBetweenTrans=" & gSecBetweenTrans, True)
        Call Genericlog("GetConfig", "Path=" & gPath, True)
        Call Genericlog("GetConfig", "FullSQLLog=" & gFullSQLLog, True)

        Call UpdateScreen(True)
        'Call GetDesimalSeparator
        If (InStr(1, LCase(VB.Command()), "user") > 0) And (InStr(1, LCase(VB.Command()), "pwd") > 0) And (InStr(1, LCase(VB.Command()), "server") > 0) Then
            gServer.User = Mid(VB.Command(), InStr(1, LCase(VB.Command()), "user") + 5, InStr(InStr(1, LCase(VB.Command()), "user"), LCase(VB.Command() & " "), " ") - (InStr(1, LCase(VB.Command()), "user") + 5))
            gServer.Pwd = Mid(VB.Command(), InStr(1, LCase(VB.Command()), "pwd") + 4, InStr(InStr(1, LCase(VB.Command()), "pwd"), LCase(VB.Command() & " "), " ") - (InStr(1, LCase(VB.Command()), "pwd") + 4))
            gServer.ConStr = Mid(VB.Command(), InStr(1, LCase(VB.Command()), "server") + 7, InStr(InStr(1, LCase(VB.Command()), "server"), LCase(VB.Command() & " "), " ") - (InStr(1, LCase(VB.Command()), "server") + 7))
        End If
        '1.0.3.1
        gLang = GetSetting(sProgSetting, "Settings", "Language") 'no
        If gLang = "" Then
            gLang = "Eng"
        End If

Exit_GetConfig:
        Exit Sub

Err_GetConfig:
        Call GenericError("FrmMain", "Feil i : GetConfig", Err.Number, Err.Description, Err.Source)
        Resume Exit_GetConfig
    End Sub
    Public Function ReadIniFil() As Short
        '1.0.3.6
        Dim sIniFil As String
        Dim sLine As String
        Dim iPos As Integer
        Dim sLeftLine As String
        Dim sRightLine As String
        Dim lFilenumber As Short '1.0.3.8
        gMethLength = 8 '
        Try
            ReadIniFil = 0
            sIniFil = My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".INI"

            lFilenumber = FreeFile()
            If File.Exists(sIniFil) Then
                FileOpen(lFilenumber, sIniFil, OpenMode.Input)
                Do While Not EOF(1)
                    Me._sbStatusBar_Panel1.Text = "Read ini file!"
                    sLine = LineInput(lFilenumber)
                    If Len(sLine) > 0 Then
                        iPos = InStr(1, LCase(sLine), "=")
                        sLeftLine = VB.Left(sLine, iPos - 1)
                        sRightLine = VB.Right(sLine, Len(sLine) - iPos)
                        Select Case sLeftLine
                            Case "ElvisServer"
                                gServer.ConStr = sRightLine
                            Case "PassWord"
                                gServer.Pwd = sRightLine
                                Call Encrypt(gServer.Pwd, gSifrCode)
                            Case "User"
                                gServer.User = sRightLine
                            Case "OrgUnit"
                                gOrgUnit = sRightLine
                            Case "MethLength" '1.0.4.3
                                gMethLength = sRightLine
                            Case "SecBetweenTrans" '1.0.4.5===>
                                gSecBetweenTrans = sRightLine
                            Case "MinBetweenRejTrans"
                                gMinBetweenRejTrans = sRightLine
                            Case "CountRetry"
                                gCountRetry = sRightLine
                            Case "OldFiles"
                                gOldFiles = sRightLine
                            Case "AutoRetry"
                                gAutoRetry = sRightLine
                            Case "Path" '1.0.4.6 fra path
                                gPath = sRightLine
                            Case "LastTimeStamp"
                                gLastTimeStamp = CDate(sRightLine)
                            Case "FullSqlLog"
                                gFullSQLLog = sRightLine
                            Case "LogToFile"
                                gLogToFile = sRightLine
                            Case "Language"
                                gLang = sRightLine
                        End Select
                        If gCountRetry = 0 Then '1.0.1.2
                            gCountRetry = 5
                        End If
                    End If
                Loop
                If (InStr(1, LCase(VB.Command()), "path") > 0) Then
                    gPath = Mid(VB.Command(), InStr(1, LCase(VB.Command()), "path") + 5, InStr(InStr(1, LCase(VB.Command()), "path"), LCase(VB.Command() & " "), " ") - (InStr(1, LCase(VB.Command()), "path") + 5))
                End If
            Else
                '1.0.5.1 flyttet ned, legge default verdi i writeinifile fordei den lages hvis den ikke finnes i startup. 
                Me._sbStatusBar_Panel1.Text = APP_NAME & " : Ini file missing."
                Call WriteIniFile()
            End If
            If gPath = "" Then
                gPath = GetSetting("X40_Receive", "Settings", "Path", "C:\Program Files\XrfTrans\Data")
            End If
            'gBBPath = GetSetting(sProgSetting, "Settings", "BBPath", "") 'no ikke nødvenig.  For Big Brother
            'frmOptions.txtBBPath.Text = gBBPath
            'gFullSQLLog = GetSetting(sProgSetting, "Settings", "FullSqlLog", "N") = "Y" 'no
            Call Genericlog("ReadIniFil", "SecBetweenTrans=" & gSecBetweenTrans, True)
            Call Genericlog("ReadIniFil", "Path=" & gPath, True)
            Call Genericlog("ReadIniFil", "FullSQLLog=" & gFullSQLLog, True)

            Call UpdateScreen(True)
            'Call GetDesimalSeparator
            If (InStr(1, LCase(VB.Command()), "user") > 0) And (InStr(1, LCase(VB.Command()), "pwd") > 0) And (InStr(1, LCase(VB.Command()), "server") > 0) Then
                gServer.User = Mid(VB.Command(), InStr(1, LCase(VB.Command()), "user") + 5, InStr(InStr(1, LCase(VB.Command()), "user"), LCase(VB.Command() & " "), " ") - (InStr(1, LCase(VB.Command()), "user") + 5))
                gServer.Pwd = Mid(VB.Command(), InStr(1, LCase(VB.Command()), "pwd") + 4, InStr(InStr(1, LCase(VB.Command()), "pwd"), LCase(VB.Command() & " "), " ") - (InStr(1, LCase(VB.Command()), "pwd") + 4))
                gServer.ConStr = Mid(VB.Command(), InStr(1, LCase(VB.Command()), "server") + 7, InStr(InStr(1, LCase(VB.Command()), "server"), LCase(VB.Command() & " "), " ") - (InStr(1, LCase(VB.Command()), "server") + 7))
            End If
            '1.0.3.1

            'MessageBox.Show("7 " & gServer.ConStr & vbNewLine & gPath & vbNewLine & gOrgUnit & vbNewLine & gServer.User)
            frmOptions.txtServ.Text = gServer.ConStr
            frmOptions.txtUser.Text = gServer.User
            frmOptions.txtPassW.Text = gServer.Pwd
            frmOptions.txtOrgUnit.Text = gOrgUnit
            frmOptions.TxtMetLen.Text = gMethLength

            '1.0.4.5 Wilab setting satt til false.  Hvis bruk, må noe fikses.
            frmOptions.txtPath.Text = gPath
            If gLang = "" Then gLang = "Eng"

            '1.0.4.7 Flere Wilab parameter settes default.  
            gNotElms.Pos = 5
            gNotElms.len_Renamed = 1
            gNotElms.txt = "-"


            'MessageBox.Show("Read ini fil : " & gServer.ConStr)

        Catch Ex As Exception
            MessageBox.Show(Ex.Message)
            'Call GenericError("FrmMain", "Feil i : GetConfig", Err.Number, Err.Description, Err.Source)
            ReadIniFil = -1
        Finally
            ReadIniFil = 1
        End Try
        FileClose(lFilenumber) '1.0.3.8

        Exit Function
ErrorExit:
        MessageBox.Show(sIniFil & " missing")
        ReadIniFil = -1
        Exit Function
    End Function
    Public Function WriteIniFile() As Short
        '1.0.3.6
        '1.0.5.1 lage default verdier.
        Dim lFilenumber As Short
        Dim sFileName As String
        Dim sCryptPassW As String = gServer.Pwd
        Try
            lFilenumber = FreeFile()
            WriteIniFile = 1
            sFileName = My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".INI"
            FileOpen(lFilenumber, sFileName, OpenMode.Output) '
            If gServer.ConStr = "" Then gServer.ConStr = "eits.nnn.elkem"
            PrintLine(lFilenumber, "ElvisServer=" & gServer.ConStr)
            If gServer.User = "" Then gServer.User = "X40"
            PrintLine(lFilenumber, "User=" & gServer.User)
            If sCryptPassW = "" Then sCryptPassW = "captain"
            Call Crypting(sCryptPassW, gSifrCode)
            PrintLine(lFilenumber, "PassWord=" & sCryptPassW)
            If gOrgUnit = "" Then gOrgUnit = "NNN"
            PrintLine(lFilenumber, "OrgUnit=" & gOrgUnit)
            If gPath = "" Then gPath = "C:\TOELMS"
            PrintLine(lFilenumber, "Path=" & gPath)
            If gMethLength = "" Then gMethLength = "8"
            PrintLine(lFilenumber, "MethLength=" & gMethLength) '1.0.4.4
            '1.0.4.5
            If gSecBetweenTrans = 0 Then gSecBetweenTrans = 10
            PrintLine(lFilenumber, "SecBetweenTrans=" & gSecBetweenTrans)
            PrintLine(lFilenumber, "MinBetweenRejTrans=" & gMinBetweenRejTrans)
            PrintLine(lFilenumber, "CountRetry=" & gCountRetry)
            PrintLine(lFilenumber, "OldFiles=" & gOldFiles)
            PrintLine(lFilenumber, "AutoRetry=" & gAutoRetry)
            PrintLine(lFilenumber, "LastTimeStamp=" & gLastTimeStamp)
            PrintLine(lFilenumber, "FullSqlLog=" & gFullSQLLog)
            PrintLine(lFilenumber, "LogToFile=" & gLogToFile)
            If gLang = "" Then gLang = "Nor"
            PrintLine(lFilenumber, "Language=" & gLang)

            FileClose(lFilenumber)
            'MessageBox.Show("3 " & gServer.ConStr & vbNewLine & gPath & vbNewLine & gOrgUnit & vbNewLine & gServer.User)
        Catch Ex As Exception
            MessageBox.Show(Ex.Message)
            Call GenericError("WriteIniFile", "Feil i UpdateScreen...", Err.Number, Err.Description, Err.Source)

        End Try
    End Function

    Private Sub UpdateScreen(ByRef pFull As Boolean)
        If Not gNoErrorHandling Then On Error GoTo Err_UpdateScreen
        If pFull Then
            '2.0.0.3
            'Me._sbStatusBar_Panel1.Text = "Venter på data!"
            Me._sbStatusBar_Panel2.Text = Now()
            Application.DoEvents()
            txtSecToTrans.Text = CStr(gSecToTrans)
            'txtMinBetwRej.Text = VB6.Format(gSecToRejTrans / 60, "##")
            txtMinBetwRej.Text = Format(gSecToRejTrans / 60, "##")
            txtLastTimeStamp.Text = Format(gLastTimeStamp, "dd.MM.yyyy hh:mm.ss")
            If gAutoRetry Then
                txtMinBetwRej.Visible = True
                lblMinToNextRej.Visible = True
            Else
                txtMinBetwRej.Visible = False
                lblMinToNextRej.Visible = False
            End If
        End If

Exit_UpdateScreen:
        Exit Sub

Err_UpdateScreen:
        Call GenericError("UpdateScreen", "Feil i UpdateScreen...", Err.Number, Err.Description, Err.Source)
        Resume Exit_UpdateScreen
    End Sub

    Public Sub SetStatus(ByRef pActive As Boolean)
        Dim i As Short
        If Not gNoErrorHandling Then On Error GoTo Err_SetStatus
        gActive = pActive
        btnStart.Enabled = Not gActive
        btnStop.Enabled = gActive
        If gActive Then
            txtStatus.Text = lblTxt(8)

            txtStatus.ForeColor = System.Drawing.ColorTranslator.FromOle(&H8000) 'Dark green
        Else
            txtStatus.ForeColor = System.Drawing.Color.Red
            txtStatus.Text = lblTxt(9)
            '1.0.5.2Call Close_db() '1.3.7
            'gActive = False
            Timer1.Enabled = False '1.3.4
            Timer2.Enabled = True '1.5.1
            Call Genericlog("frmMain", "SetStatus", True)
        End If
        frmOptions.txtSecBetweenTrans.Enabled = Not gActive
        frmOptions.txtPath.Enabled = Not gActive

Exit_SetStatus:
        Exit Sub

Err_SetStatus:
        Call GenericError("SetStatus", "Feil i SetStatus...", Err.Number, Err.Description, Err.Source)
        Resume Exit_SetStatus
    End Sub

    Public Sub SetStatusDetail(ByRef pMess As String, Optional ByRef pAddToExisting As Boolean = False)
        If Not gNoErrorHandling Then On Error GoTo Err_SetStatusDetail

        If Not pAddToExisting Then
            txtStatusDetail.Text = pMess
        Else
            txtStatusDetail.Text = txtStatusDetail.Text & " - " & pMess
        End If
        Me.Refresh()

Exit_SetStatusDetail:
        Exit Sub

Err_SetStatusDetail:
        Call GenericError("SetStatusDetail", "Feil i SetStatusDetail...", Err.Number, Err.Description, Err.Source)
        Resume Exit_SetStatusDetail
    End Sub


    Public Sub ServList()

        If Not gNoErrorHandling Then On Error GoTo Err_ServList

        txtServer.Text = ""
        With gServer
            '.ServNo = 1 nei, nei ?1.4.6
            txtServer.Text = .ConStr & " - " & .User & " - " & gOrgUnit
            txtServer.BackColor = Color.WhiteSmoke '5.1.0

        End With
        'Next X
        Timer1.Enabled = True
        Timer2.Enabled = False
        Call Genericlog("frmMain", "Skrevet serverliste", True)
        Call LoadResourceStrings() '1.0.4.5
Exit_ServList:
        Exit Sub

Err_ServList:
        Call GenericError("ServList", "Feil i ServList-kode...", Err.Number, Err.Description, Err.Source)
        Resume Exit_ServList
    End Sub


    Private Sub CreRejFileRecordSet()
        If Not gNoErrorHandling Then On Error GoTo Err_CreRejFileRecordSet
        Dim rst As ADODB.Recordset
        Dim MyName As Object
        Dim lValid As Boolean
        ' Instantiate the recordset.
        rst = New ADODB.Recordset

        ' Append fields
        rst.Fields.Append("FileName", ADODB.DataTypeEnum.adVarChar, 20)
        rst.Fields.Append("ServNo", ADODB.DataTypeEnum.adSmallInt)
        rst.Fields.Append("Status", ADODB.DataTypeEnum.adSmallInt)
        rst.Fields.Append("Retry", ADODB.DataTypeEnum.adSmallInt) '1.4.0
        ' Put some data in the recordset
        rst.Open()
        MyName = Dir(gPath & "\rej\*.*") ' Retrieve the first entry.
        Do While MyName <> "" ' Start the loop.
            rst.AddNew(New Object() {"FileName", "ServNo", "Status", "Retry"}, New Object() {MyName, 0, 0, 0})
            MyName = Dir() ' Get next entry.
        Loop
        rst.UpdateBatch()
        On Error GoTo 0
        If Dir(My.Application.Info.DirectoryPath & "\RejFilesStatus.dat") & "" <> "" Then
            lValid = True
        End If
        If Not gNoErrorHandling Then On Error GoTo Err_CreRejFileRecordSet
        If lValid Then
            Kill(My.Application.Info.DirectoryPath & "\RejFilesStatus.dat")
        End If
        rst.Save(My.Application.Info.DirectoryPath & "\RejFilesStatus.dat")

        '' Dump the recordset to the Immediate Window
        If Not rst.EOF Then
            rst.MoveFirst()
        End If
        Do Until rst.EOF
            'Debug.Print rst("FileName"), rst("ServNo")
            rst.MoveNext()
        Loop
        rst.Close()
        'UPGRADE_NOTE: Object rst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rst = Nothing
Exit_CreRejFileRecordSet:
        Exit Sub

Err_CreRejFileRecordSet:
        Call GenericError("frmRejFileStatus", "Feil i CreRejFileRecordSet...", Err.Number, Err.Description, Err.Source)
        Resume Exit_CreRejFileRecordSet
    End Sub

    Private Sub Timer2_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer2.Tick
        Call QueCount(1)
    End Sub

    Private Sub LogToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogToolStripMenuItem.Click

        'Dim fbd1 As FolderBrowserDialog
        Dim openFileDialog1 As New OpenFileDialog()
        Dim sMappe As Object
        Dim sFileName As String = String.Empty
        Dim lPath As String = My.Application.Info.DirectoryPath & "\log"

        sMappe = Nothing
        If Not gActive Then
            'Debug.Print(lPath)
            openFileDialog1.InitialDirectory = lPath
            openFileDialog1.Filter = "txt files (*.log)|*.txt|All files (*.*)|*.*"
            openFileDialog1.FilterIndex = 2
            openFileDialog1.RestoreDirectory = True
            If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Try
                    sMappe = openFileDialog1.OpenFile()
                    sFileName = openFileDialog1.FileName
                    'Debug.Print(sFileName)
                    '    If (sMappe IsNot Nothing) Then
                    '//do your loop here
                    '    End If
                Catch Ex As Exception
                    'MessageBox.Show(Ex.Message)
                    Call GenericError("UpdateScreen", "Feil i LogToolStripMenuItem_Click...", Err.Number, Ex.Message, Err.Source)
                Finally
                    If (sMappe IsNot Nothing) Then
                        sMappe.Close()
                    End If
                End Try
            End If
            System.Diagnostics.Process.Start("Notepad.Exe", sFileName)
            'TextBox1.Text = sMappe
        Else
            MsgBox(lblTxt(6))
        End If

    End Sub

    Private Sub FeilToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FeilToolStripMenuItem.Click

        'Dim fbd1 As FolderBrowserDialog
        Dim openFileDialog1 As New OpenFileDialog()
        Dim sMappe As Object
        Dim sFileName As String = String.Empty
        'Dim lPath As String = My.Application.Info.DirectoryPath & "\log"
        Dim lPath As String = My.Application.Info.DirectoryPath & "\error"

        sMappe = Nothing
        If Not gActive Then
            'Debug.Print(lPath)
            openFileDialog1.InitialDirectory = lPath
            openFileDialog1.Filter = "txt files (*.log)|*.txt|All files (*.*)|*.*"
            openFileDialog1.FilterIndex = 2
            openFileDialog1.RestoreDirectory = True
            If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Try
                    sMappe = openFileDialog1.OpenFile()
                    sFileName = openFileDialog1.FileName
                    Debug.Print(sFileName)
                    '    If (sMappe IsNot Nothing) Then
                    '//do your loop here
                    '    End If
                Catch Ex As Exception
                    MessageBox.Show(Ex.Message)
                Finally
                    If (sMappe IsNot Nothing) Then
                        sMappe.Close()
                    End If
                End Try
            End If
            System.Diagnostics.Process.Start("Notepad.Exe", sFileName)
            'TextBox1.Text = sMappe
        Else
            MsgBox(lblTxt(6))
        End If

    End Sub

    Private Sub cmdRejFileStatus_Click(sender As Object, e As EventArgs) Handles cmdRejFileStatus.Click
        '1.0.5.3
        MessageBox.Show("Stopp log og finn feil og log filer i menyen : Verktøy!")

    End Sub
End Class