Option Strict Off
Option Explicit On
Imports System
Imports System.IO

Module ModMolabGlobals
    'version history
    'Magne Larsen, iStone
    '1.0.0.0 08.03.17 Fra Xrf2Elms til Molab2Lims se kommentarer på tidligere endringer i ModGlobals.vb i Xrf2Elms prosjekt.
    '1.0.0.1 15.03.17 Har ikke data i fila for antall parameter, så overfører til fila er tom.
    '1.0.0.3    "     Ny deploy med rikti versjonsnr
    '1.0.0.4 21.03.17 Feilsøk 
    '1.0.0.5 05.04.17 Nye kollonner: GRADE,FRACTION og CONTAINER_NO -> GRADE_CODE,FRACTION,CONTAINER_NO
    '1.0.0.6 25.04.17 Finpuss: 
    '1.0.1.0 25.04.17 For å angi endring
    '1.0.1.1 12.05.17 Langt kanal navn og utf-8/ansi problem (å i kanalnavn)
    '1.0.1.2 06.06.17 Message box. Sjekk om metode finnes
    '1.0.1.3 25.10.17 Må ikke ha message box
    '1.0.1.4 08.01.18 Ingen ting.  Kun test Team Explorter  Se notatblokken til Magne : _Net Sync version Control
    '1.0.1.5 06.05.21 Mulighet for å sjekke at mapping er ok.
    '*************************************************
    '1.6.4 definition
    Structure NotElmsData
        Dim Use As Boolean
        Dim FileMove As Boolean
        Dim Path As String
        Dim Pos As Short
        Dim len_Renamed As Short
        Dim txt As String
    End Structure
    Public gNotElmsSample As Boolean
    '1.0.0.3 Public gNotElms(2) As NotElmsData ' 1.6.7
    Public gNotElms As NotElmsData ' 1.6.7

    Public Const APP_NAME As String = "Molab2Lims 1.0.1.5"
    Public gNoErrorHandling As Boolean
    Public gNoConfirm As Boolean
    Public gNoSave As Boolean
    Public gFullSQLLog As Boolean
    Public gLogToFile As Boolean '1.6.4
    Public gOrgUnit As String
    Public gMethLength As String '1.0.4.3
    'Public gOriginCode As String
    'Public gSampNoCode As String
    Public gPath As String
    Public gAutoRetry As Boolean
    Public gSecToRejTrans As Single
    Public gMinBetweenRejTrans As Single
    Public gCountRetry As Single
    Public gOldFiles As Single
    Public gSecToTrans As Single
    Public gSecBetweenTrans As Single
    'BigBrother stat.
    Public BBStat(6) As Double
    Public gBBPath As String
    '1:antall ok prøver
    '2:antall i feil katalog
    '3:antall i prossesskatalog
    '4:antall i rejectkatalog
    '5:antall i arkivkatalog
    '6:antall feilconnect
    Public gLastTimeStamp As Date
    Public gActive As Boolean
    Public gProcessing As Boolean
    'Public mConnected As Boolean
    'Public gMethCodePar As String
    'Public gMethCodeData As String
    'Public gMethCodeRes As String
    Public gBackUpFile As Boolean 'ml 001108
    Public Const M_FILE_ID As Short = 100
    'Public gServNO As Integer
    'Public gServCount As Short
    Public Const gSifrCode As String = "Magne"
    Public gLang As String '1.0.3.1
    Public gConnect As String '1.0.3.2
    Public gUsedOraTransProg '1.0.4.9
    Public giPrevToPos As Integer = 0
    Public giPrevDiff As Integer = 1

    Structure ServInfo
        Dim ServNo As Short
        '1.0.5.2 Dim dbCon As ADODB.Connection
        Dim User As String
        Dim Pwd As String
        Dim ConStr As String
        Dim Connected As Boolean
        Dim TransFlag As Short '0=skal ikke overføres,1=skal,2=prøve overført,3=verdi overført
        Dim Retry As Short 'antall forsøk på ny overføring
        Dim NCMess As Short 'antall parameter overført
        Dim NRealCmess As Short 'antall parameter målt
        Dim TransFile As String
    End Structure
    'Public gServers(3) As ServInfo
    Public gServer As ServInfo '2.0.2.0
    Public gRejServers() As Short
    Public gRejServN As Short
    '1.5.0 Arl
    Public gSource As String '1.0.0.0 Molab

    'gServCount
    'gServers.ServCount
    Public Sub GenericError(ByRef pProc As String, ByRef pErrHeading As String, ByRef pErrNo As Integer, ByRef pErrDescr As String, ByRef pErrSource As String)
        On Error Resume Next
        Dim lMess As String
        Dim lOraMess As String
        Dim lFileNumber As Short
        Dim lPath As String
        lPath = My.Application.Info.DirectoryPath & "\error"
        '1.0.2.0 Dim fs As Object
        'fs = CreateObject("Scripting.FileSystemObject")
        'If Not fs.FolderExists(lPath) Then

        If Directory.Exists(lPath) = False Then
            'MkDir(lPath)
            My.Computer.FileSystem.CreateDirectory(lPath)
        End If
        'fs = Nothing
        If gLang = "Eng" Then
            lMess = "There was an error in " & pProc & vbNewLine & vbNewLine & "Error number: " & pErrNo & vbNewLine & "Error message: " & pErrDescr '& |            '"Err.Source : " & pErrSource
        Else
            lMess = "Det oppstod en feil i " & pProc & vbNewLine & vbNewLine & "Feil nummer: " & pErrNo & vbNewLine & "Feilmelding: " & pErrDescr '& |            '"Err.Source : " & pErrSource
        End If
        If gLang = "Eng" Then
            FrmMain._sbStatusBar_Panel1.Text = "Error, see log!!!"
        Else
            FrmMain._sbStatusBar_Panel1.Text = "Oppstod feil, se log!!!"
        End If
        lFileNumber = FreeFile()
        '3.0.5 finn path for log fil
        'My.Application.Info
        FileOpen(lFileNumber, lPath & "\" & My.Application.Info.AssemblyName & FileNameIncrement() & ".err", OpenMode.Append) '
        PrintLine(lFileNumber, vbNewLine & "------------------------------------------" & vbNewLine & Format(Now, "dd.MM.yyyy hh:mm:ss") & vbNewLine & pErrHeading & vbNewLine & lMess)
        FileClose(lFileNumber)
        Call Genericlog(pProc, "ERROR : " & pErrHeading & vbNewLine & lMess)
    End Sub
    Public Sub Genericlog(ByRef pProc As String, ByRef pMess As String, Optional ByRef pNoHeader As Boolean = False)
        On Error Resume Next
        Dim lMess As String
        Dim lFileNumber As Short
        Dim lPath As String
        lPath = My.Application.Info.DirectoryPath & "\log"
        'Dim fs As Object
        lMess = ""
        If Directory.Exists(lPath) = False Then
            My.Computer.FileSystem.CreateDirectory(lPath)
        End If
        lFileNumber = FreeFile()
        FileOpen(lFileNumber, lPath & "\" & My.Application.Info.AssemblyName & FileNameIncrement() & ".log", OpenMode.Append)
        If Not pNoHeader Then
            lMess = vbNewLine & "------------------------------------------" & vbNewLine & Format(Now, "dd.MM.yyyy hh:mm:ss") & vbNewLine
        End If
        'Print #lFileNumber, lMess & pProc & ": " & pMess
        PrintLine(lFileNumber, lMess & " " & pMess)
        FileClose(lFileNumber)
    End Sub
    Public Sub getServInfo()
        Dim lPwd As String
        If Not gNoErrorHandling Then On Error GoTo Err_getServInfo
        Dim rst As ADODB.Recordset
        Dim i As Short
        Dim lValid As Boolean
        rst = New ADODB.Recordset
        'rst.Save App.Path & "\Server.dat"
        On Error GoTo 0
        If Dir(My.Application.Info.DirectoryPath & "\Server.dat") & "" <> "" Then
            lValid = True
        End If
        If Not gNoErrorHandling Then On Error GoTo Err_getServInfo
        If lValid Then
            rst.Open(My.Application.Info.DirectoryPath & "\Server.dat")
            'gServCount = rst.RecordCount
            'ReDim gServers(gServCount)
            'rst.MoveFirst
            With gServer
                .ServNo = rst.Fields("ServNo").Value
                .ConStr = rst.Fields("ServName").Value
                .User = rst.Fields("User").Value
                lPwd = rst.Fields(3).Value
                'dekrypter fra fil
                Call Encrypt(lPwd, gSifrCode)
                .Pwd = lPwd
                rst.MoveNext()
            End With
            rst.Close()
            rst = Nothing
        Else
            'ReDim gServers(gServCount)
            With gServer
                .ServNo = 1
            End With
        End If
Exit_getServInfo:
        Exit Sub

Err_getServInfo:
        Call GenericError("getServInfo", "Feil ved lesing av org_unit fra Oracle", Err.Number, Err.Description, Err.Source)
        Resume Exit_getServInfo
    End Sub

    Public Sub SQLLog(ByRef pMess As String)
        If Not gNoErrorHandling Then On Error GoTo Err_SQLLog
        Dim lMess As String = String.Empty
        Dim lFileNumber As Short

        lFileNumber = FreeFile()
        FileOpen(lFileNumber, gPath & "\log\TransSQL" & FileNameIncrement() & ".log", OpenMode.Append)
        If gNoSave Then
            lMess = "TEST: "
        End If
        PrintLine(lFileNumber, Format(Now, "dd.MM.yyyy hh:mm:ss") & " " & lMess & pMess)
        FileClose(lFileNumber)

Exit_SQLLog:
        Exit Sub

Err_SQLLog:
        Call GenericError("SQLLog", "Feil ved skriving til SQL-logg", Err.Number, Err.Description, Err.Source)
        Resume Exit_SQLLog
    End Sub

    Public Sub ProcessingLog(ByRef pOk As Boolean, ByRef pMess As String)
        If Not gNoErrorHandling Then On Error GoTo Err_ProcessingLog
        Dim lFileNumber As Short
        Dim lFileName As String
        If gLogToFile Then
            If pOk Then
                lFileName = gPath & "\log\XrfTransFile" & FileNameIncrement() & ".log"
            Else
                lFileName = gPath & "\log\XrfTransRej" & FileNameIncrement() & ".log"
            End If
            lFileNumber = FreeFile()
            FileOpen(lFileNumber, lFileName, OpenMode.Append)
            PrintLine(lFileNumber, pMess)
            FileClose(lFileNumber)
        End If
Exit_ProcessingLog:
        Exit Sub

Err_ProcessingLog:
        Call GenericError("ProcessingLog", "eil ved skriving til logg", Err.Number, Err.Description, Err.Source)
        Resume Exit_ProcessingLog
    End Sub

    Public Function FileNameIncrement() As String
        If Not gNoErrorHandling Then On Error GoTo Err_FileNameIncrement

        FileNameIncrement = Format(Now, "-yyyy-MM")

Exit_FileNameIncrement:
        Exit Function

Err_FileNameIncrement:
        'Do NOT log this error - might cause sircular error...
        FileNameIncrement = ""
        Resume Exit_FileNameIncrement
    End Function
    Sub Encrypt(ByRef secret As String, ByRef PassWord As String)
        ' secret$ = the string you wish to encrypt or decrypt.
        ' PassWord = the password with which to encrypt the string.
        If Not gNoErrorHandling Then On Error GoTo Err_Encrypt
        Dim X As Object
        Dim l As Short
        Dim i As Integer = Asc(Mid(secret, 1, 1))
        If i < 64 Or i > 127 Then
            Dim char_Renamed As String
            l = Len(PassWord)
            For X = 1 To Len(secret)
                char_Renamed = CStr(Asc(Mid(PassWord, (X Mod l) - l * CShort((X Mod l) = 0), 1)))
                Mid(secret, X, 1) = Chr(Asc(Mid(secret, X, 1)) Xor char_Renamed)
            Next
        End If
Exit_Encrypt:
        Exit Sub

Err_Encrypt:
        Call GenericError("Encrypt", "Feil i Encrypt", Err.Number, Err.Description, Err.Source)
        '2.0.0 erstatett : Call Lib_VBError("modGlobals", "Encrypt", Erl(), "Feil i Encrypt")
        Resume Exit_Encrypt
    End Sub
    Sub Crypting(ByRef secret As String, ByRef PassWord As String)
        ' secret$ = the string you wish to encrypt or decrypt.
        ' PassWord = the password with which to encrypt the string.
        If Not gNoErrorHandling Then On Error GoTo Err_Crypting
        Dim X As Object
        Dim l As Short
        Dim i As Integer = Asc(Mid(secret, 1, 1))
        If i >= 64 Or i <= 127 Then
            Dim char_Renamed As String
            l = Len(PassWord)
            For X = 1 To Len(secret)
                char_Renamed = CStr(Asc(Mid(PassWord, (X Mod l) - l * CShort((X Mod l) = 0), 1)))
                Mid(secret, X, 1) = Chr(Asc(Mid(secret, X, 1)) Xor char_Renamed)
            Next
        End If
Exit_Crypting:
        Exit Sub

Err_Crypting:
        Call GenericError("Crypting", "Feil i Crypting", Err.Number, Err.Description, Err.Source)
        Resume Exit_Crypting
    End Sub
End Module