Option Strict Off
Option Explicit On
Imports System
Imports System.IO
Module ModMolabFileSelection
    Dim mReject As Boolean
    'Molab2Lims
    'Add files here that are generated/updated from the Viscometer in the data directory
    ' (files that should not be logged into Oracle)

    Public Function IgnoreFile(ByRef pFile As String) As Boolean
        If LCase(pFile) = "dataread.txt" Or LCase(pFile) = "combo1.txt" Then
            IgnoreFile = True
        Else
            IgnoreFile = False
        End If
    End Function

    Public Sub MakeCatStr()
        mReject = False
        If Not gNoErrorHandling Then On Error GoTo Err_MakeCatStr
        'Dim fs As Object
        'fs = CreateObject("Scripting.FileSystemObject")
        If Directory.Exists(gPath) = False Then
            My.Computer.FileSystem.CreateDirectory(gPath)
            'If Not fs.FolderExists(gPath) Then
            MkDir(gPath)
        End If
        If Not Directory.Exists(gPath & "\temp\") Then
            My.Computer.FileSystem.CreateDirectory(gPath & "\temp\")
        End If
        If Not Directory.Exists(gPath & "\accept\") Then
            My.Computer.FileSystem.CreateDirectory(gPath & "\accept\")
        End If
        If Not Directory.Exists(gPath & "\transfer\") Then
            My.Computer.FileSystem.CreateDirectory(gPath & "\transfer\")
        End If
        If Not Directory.Exists(gPath & "\log\") Then
            My.Computer.FileSystem.CreateDirectory(gPath & "\log\")
        End If
        If Not Directory.Exists(gPath & "\rej\") Then
            My.Computer.FileSystem.CreateDirectory(gPath & "\rej\")
        End If
        If Not Directory.Exists(gPath & "\error\") Then
            My.Computer.FileSystem.CreateDirectory(gPath & "\error\")
        End If
        If Not Directory.Exists(gPath & "\WilabError") Then '1.6.9
            My.Computer.FileSystem.CreateDirectory(gPath & "\WilabError")
        End If
        If Not Directory.Exists(gBBPath) And gBBPath <> "" Then
            My.Computer.FileSystem.CreateDirectory(gBBPath)
        End If
        'fs = Nothing
Exit_MakeCatStr:
        Exit Sub

Err_MakeCatStr:
        Call GenericError("modFile", "Feil i MakeCatStr...", Err.Number, Err.Description, Err.Source)
        Resume Exit_MakeCatStr

    End Sub
    Public Sub PrintBBFile()
        'rev 1.5.0
        On Error Resume Next
        Dim lFileNumber, i As Short
        Dim sColor As String
        Dim sArcCat, sProsCat, sErrCat, sRejCat, sErrCon As String
        Dim sArcColor, sProsColor, sErrColor, sRejColor, sConColor As String
        sErrCat = " Files in error catalog : " & BBStat(2)
        sProsCat = " Files in process catalog : " & BBStat(3)
        sRejCat = " Files in reject catalog : " & BBStat(4)
        sArcCat = " Files in archive catalog : " & BBStat(5)
        sErrCon = " Count error connections : " & BBStat(6)
        sColor = "green"
        sErrColor = "green"
        sProsColor = "green"
        sRejColor = "green"
        sArcColor = "green"
        sConColor = "green"
        If BBStat(3) > 3 Then 'prosess
            sProsColor = "red"
            sColor = "red"
        End If
        If BBStat(6) > 2 Then 'base con
            sConColor = "red"
            sColor = "red"
        End If
        If BBStat(4) > 2 Then 'rej
            sRejColor = "red"
            sColor = "red"
        End If
        If BBStat(2) > 5 Then 'error
            sErrColor = "yellow" '1.5.4
            sColor = "yellow"
        End If
        lFileNumber = FreeFile()
        FileOpen(lFileNumber, gBBPath & "\XrfTrans", OpenMode.Output) '1.5.4
        PrintLine(lFileNumber, sColor & " XrfTrans Report " & Format(Now, "dd.MM.yyyy hh:mm:ss") & " [" & g_sComputerName & "]") '1.5.6
        PrintLine(lFileNumber, "&" & sErrColor & sErrCat)
        PrintLine(lFileNumber, "&" & sProsColor & sProsCat)
        PrintLine(lFileNumber, "&" & sRejColor & sRejCat)
        PrintLine(lFileNumber, "&" & sArcColor & sArcCat)
        PrintLine(lFileNumber, "&" & sConColor & sErrCon)
        FileClose(lFileNumber)
    End Sub
    Public Sub RemoveOldFiles()
        Dim lCurrTimeStamp As Date
        Dim pPath, curr_file As String
        'Dim fs As Object
        Dim iLastDelTime As Date
        Dim i As Short
        pPath = gPath & "\accept"

        curr_file = Dir(pPath & "\*.*", FileAttribute.Archive)
        iLastDelTime = System.DateTime.FromOADate(Now.ToOADate - gOldFiles)
        If curr_file <> "" Then
            lCurrTimeStamp = FileDateTime(pPath & "\" & curr_file)
        End If
        i = 0
        '1.4.3 i <5000
        Do While curr_file <> "" And i < 5000
            If iLastDelTime > lCurrTimeStamp Then
                File.Delete(pPath & "\" & curr_file)
            End If
            i = i + 1
            curr_file = Dir()
            If curr_file <> "" Then
                lCurrTimeStamp = FileDateTime(pPath & "\" & curr_file)
            End If
        Loop
    End Sub
    Public Function MakeUId() As Integer
        If Not gNoErrorHandling Then On Error GoTo Err_MakeUId
        Randomize()
        MakeUId = Int((1000000 - 10000000 + 1) * Rnd() + 10000000)
        Exit Function
Exit_MakeUId:
        Exit Function
Err_MakeUId:
        MakeUId = 0
        Resume Exit_MakeUId
    End Function
End Module