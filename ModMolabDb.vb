Option Strict Off
Option Explicit On
Imports System.Data '1.0.3.2
Imports Oracle.ManagedDataAccess.Client
'Imports Oracle.ManagedDataAccess.Types
'Molab2Lims
Module ModMolabDb

    Public Sub StoreMessage(ByRef conn As OracleConnection, ByRef pMess As String)
        '** ****************************************
        '1.0.4.8 Lagre melding i elp_mess
        '** *********************************************
        Dim sQuery As String = "insert into elp_mess (elms_no,ERROR_MESS,CORR_DATE) values (0,'" & pMess & "',SYSDATE)"
        Try
            If Not gNoSave Then
                If conn.State Then
                    gServer.Connected = True
                    FrmMain.txtServer.BackColor = Color.GreenYellow
                    '1.0.4.0 System.Windows.Forms.Application.DoEvents()
                Else
                    conn.Open()
                    gServer.Connected = True
                    FrmMain.txtServer.BackColor = Color.GreenYellow
                    'Debug.Print conn.ConnectionString
                    '1.0.4.0 System.Windows.Forms.Application.DoEvents()
                End If
                If gServer.Connected = True Then
                    gServer.TransFlag = 1
                    Dim cmd As New OracleCommand
                    cmd.Connection = conn
                    cmd.CommandText = sQuery
                    cmd.ExecuteNonQuery()
                    conn.Close()
                    cmd = Nothing
                    gServer.Connected = False
                    FrmMain.txtServer.BackColor = Color.White
                End If
            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message & " in StoreMessage")
            Call GenericError("modDb", "Feil i StoreMessage...", Err.Number, ex.Message, Err.Source)

        End Try
    End Sub

    Public Function StoreSamp(ByRef conn As OracleConnection, ByRef pMethcode As String, ByRef pSmess As String, ByRef gOrgUnit As String) As Boolean
        'ml 010314
        'INSTR_ELMS2.CREATE_INSTR_SAMPLE('SIMN96P','         BINGE1FLYT20000006ABC230201     0902251446','ITS');
        Dim iCurServer As Short
        Dim msgTxt, msgTxt2 As String
        Dim X As Short
        '1.0.3.2 System.Windows.Forms.Application.DoEvents()
        Try
            If Not gNoSave Then
                pMethcode = RemoveSpaces(pMethcode)
                'Debug.Print(pSmess & pMethcode & gOrgUnit)
                If gServer.Connected = True Then
                    gServer.TransFlag = 1
                    iCurServer = 1
                    '1.0.3.2 cmdNewSamp = New ADODB.Command
                    Dim cmdNewSamp As New OracleCommand
                    gUsedOraTransProg = ""
                    With cmdNewSamp
                        .Connection = conn '1.
                        If gOrgUnit = "BRE" Then '2.0.1.4
                            .CommandText = "ELMS.INSTR_ELMS2B.CREATE_INSTR_SAMPLE"
                        Else
                            .CommandText = "ELMS.INSTR_ELMS2.CREATE_INSTR_SAMPLE"
                        End If
                        '.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        .CommandType = CommandType.StoredProcedure
                        Dim p(2) As OracleParameter
                        p(0) = New OracleParameter()
                        p(0).OracleDbType = OracleDbType.Varchar2
                        p(0).ParameterName = "p_methcode"
                        p(0).Value = pMethcode
                        p(1) = New OracleParameter()
                        p(1).OracleDbType = OracleDbType.Varchar2
                        p(1).ParameterName = "p_ident"
                        p(1).Value = pSmess
                        p(2) = New OracleParameter()
                        p(2).OracleDbType = OracleDbType.Varchar2
                        p(2).ParameterName = "p_deforgunit"
                        p(2).Value = gOrgUnit

                        .Parameters.Add(p(0))
                        .Parameters.Add(p(1))
                        .Parameters.Add(p(2))

                        msgTxt = .CommandText
                        msgTxt2 = ""
                        For X = 0 To 2
                            msgTxt2 = msgTxt2 & .Parameters(X).Value & ","
                        Next X
                        msgTxt = msgTxt & " - " & msgTxt2
                        'Debug.Print(msgTxt)
                        'Debug.Print("'" & pSmess & "'")
                        If gFullSQLLog Then Call SQLLog(msgTxt)
                        .ExecuteNonQuery()
                        p = Nothing
                        gServer.TransFlag = 2
                        StoreSamp = True
                    End With
                    cmdNewSamp = Nothing
                ElseIf gServer.Connected = False Then
                    gServer.TransFlag = 4
                End If

            End If
            StoreSamp = True
            System.Windows.Forms.Application.DoEvents()
        Catch Ex As Exception
            StoreSamp = False
            gServer.TransFlag = 4
            'Debug.Print("StoreSamp : " & vbNewLine & Err.Number & vbNewLine & Err.Description & Err.Source)
            '1.0.1.3 Slik kan vi ikke ha det. MessageBox.Show(Ex.Message)
            Call GenericError("modDb", "Feil i StoreSamp...", Err.Number, Ex.Message, Err.Source)
            'Finally
            '   System.Windows.Forms.Application.DoEvents()
        End Try
        Exit Function

Err_StoreSamp:
        'UpdateInfo
        'StoreInfo
        StoreSamp = False
        gServer.TransFlag = 4
        System.Windows.Forms.Application.DoEvents()
        'Debug.Print("StoreSamp : " & vbNewLine & Err.Number & vbNewLine & Err.Description & Err.Source)
        Call GenericError("modDb", "Feil i StoreSamp...", Err.Number, Err.Description, Err.Source)
        'Resume Exit_StoreSamp

    End Function


    Public Function StoreRes(ByRef conn As OracleConnection, ByRef pKanal As String, ByRef pvalue As Double, ByRef pCompound As String, ByRef pUnit As String, ByRef pParInfo As String) As Boolean
        'Public Function StoreRes(ByRef pKanal As String, ByRef pvalue As String, ByRef pCompound As String, ByRef pUnit As String, ByRef pParInfo As String) As Boolean
        'INSTR_ELMS2.CREATE_INSTR_VALUE('As2',1.23123,'!');
        'INSTR_ELMS2.CREATE_INSTR_VALUE('Al',0.087,')');
        'LAB_INSTR.CREATE_INSTR_VALUE('As2',1.23123,'As2O3');
        'If Not gNoErrorHandling Then On Error GoTo Err_StoreRes

        'Dim cmdNewSamp As ADODB.Command
        Dim msgTxt, msgTxt2 As String
        Dim X, iCurServer As Short
        'System.Windows.Forms.Application.DoEvents()
        pKanal = RemoveSpaces(pKanal)
        pCompound = RemoveSpaces(pCompound)
        pUnit = RemoveSpaces(pUnit)
        'Debug.Print pvalue & pKanal & pCompound & pKanal
        'Get current elms_no
        Try
            If Not gNoSave Then
                If gServer.Connected Then
                    '1.0.0.1 gServer.TransFlag = 2 And
                    iCurServer = 1
                    Dim cmdNewSamp As New OracleCommand '1.0.3.2
                    With cmdNewSamp
                        .Connection = conn
                        If gOrgUnit = "BRE" Then '2.0.1.4
                            .CommandText = "ELMS.INSTR_ELMS2B.CREATE_INSTR_VALUE" '2.0.0.1
                        Else
                            .CommandText = "ELMS.INSTR_ELMS2.CREATE_INSTR_VALUE" '2.0.1.3
                        End If
                        .CommandType = CommandType.StoredProcedure
                        Dim p(4) As OracleParameter
                        p(0) = New OracleParameter()
                        p(0).OracleDbType = OracleDbType.Varchar2
                        'p(0).Direction = ParameterDirection.InputOutput
                        p(0).ParameterName = "p_kanal"
                        p(0).Value = pKanal

                        p(1) = New OracleParameter()
                        '1.0.3.3 
                        p(1).OracleDbType = OracleDbType.Double
                        p(1).ParameterName = "p_value"
                        p(1).Value = pvalue

                        p(2) = New OracleParameter()
                        p(2).OracleDbType = OracleDbType.Varchar2
                        p(2).ParameterName = "p_compound"
                        p(2).Value = pCompound

                        p(3) = New OracleParameter()
                        p(3).OracleDbType = OracleDbType.Varchar2
                        p(3).ParameterName = "p_parinfo"
                        p(3).Value = pParInfo

                        .Parameters.Add(p(0))
                        .Parameters.Add(p(1))
                        .Parameters.Add(p(2))
                        .Parameters.Add(p(3))

                        msgTxt = .CommandText
                        msgTxt2 = ""
                        For X = 0 To 3
                            msgTxt2 = msgTxt2 & .Parameters(X).Value & ","
                        Next X
                        msgTxt = msgTxt & " - " & msgTxt2
                        If gFullSQLLog Then Call SQLLog(msgTxt)
                        Debug.Print(msgTxt)
                        .ExecuteNonQuery()
                        gServer.TransFlag = 3
                        p = Nothing
                        gServer.NCMess = gServer.NCMess + 1
                    End With
                    cmdNewSamp = Nothing
                End If
            End If
            StoreRes = True
        Catch Ex As Exception
            StoreRes = False
            gServer.TransFlag = 4
            'MessageBox.Show(Ex.Message)
            Call GenericError("ModDb", "Feil i StoreRes...", Err.Number, Ex.Message, Err.Source)
        Finally
            System.Windows.Forms.Application.DoEvents()
        End Try
Exit_StoreRes:
        Exit Function

Err_StoreRes:
        'UpdateInfo
        'StoreInfo
        'Resume Exit_StoreRes

    End Function
    Public Function MethCodeExist(conn As OracleConnection, ByRef pProgName As String) As Boolean
        '** ******************************
        '** Ny ifm Molab2Lims 07.06.17 1.0.1.2
        '** ******************************
        Dim lSql As String = String.Empty
        Dim sMethCode As String = String.Empty
        MethCodeExist = False
        Try
            If gFullSQLLog Then Call SQLLog(lSql)
            lSql = "select METHCODE from ELP_METH where METH_ORG_UNIT='" & gOrgUnit & "' and METHCODE ='" & pProgName & "' AND nvl(delete_flag,'N')='N'"
            'Debug.Print(lSql)
            Dim cmd As New OracleCommand
            If conn.State Then '1.0.0.4
                gServer.Connected = True
            Else
                conn.Open()
                gServer.Connected = True
            End If
            cmd.Connection = conn
            cmd.CommandText = lSql
            Dim dr As OracleDataReader = cmd.ExecuteReader()
            dr.Read()
            'rstSampCode.Open(lSql, oradb, , , ADODB.CommandTypeEnum.adCmdUnknown)
            'Do Until rstSampCode.EOF
            If dr.HasRows Then
                sMethCode = dr.Item("METHCODE")
                MethCodeExist = True
            Else
                MethCodeExist = False
            End If
        Catch Ex As Exception
            Call GenericError("MethCodeExist", "Feil ved lesing av metode fra Oracle", Err.Number, Ex.Message, Err.Source)
        End Try

    End Function

    Public Function GetIdentLen(conn As OracleConnection, ByRef pMethName As String, pColName As String) As Integer
        '** ******************************
        '** Ny ifm Molab2Lims 09.03.17
        '** ******************************
        Dim lSql As String = String.Empty
        Dim liFrom As Integer = 0
        Dim liTo As Integer = 0
        GetIdentLen = 0
        Try
            If gFullSQLLog Then Call SQLLog(lSql)
            lSql = "SELECT from_pos,to_pos FROM ELP_IDENT P,ELP_METH M,ELP_INSTR I  WHERE M.INSTR_CODE = I.INSTR_CODE  AND I.IDENT_CODE = P.IDENT_CODE "
            lSql = lSql & "And m.methcode = '" & pMethName & "' AND  SAMP_COLUMN='" & pColName & "' order by from_pos"
            Dim cmd As New OracleCommand
            If conn.State Then '1.0.0.4
                gServer.Connected = True
            Else
                conn.Open()
                gServer.Connected = True
            End If
            cmd.Connection = conn
            cmd.CommandText = lSql

            'cmd.CommandType = CommandType.Text
            Dim dr As OracleDataReader = cmd.ExecuteReader()
            dr.Read()
            liFrom = dr.Item(0)
            liTo = dr.Item(1)
            If giPrevToPos > 0 Then
                giPrevDiff = liFrom - giPrevToPos
            Else
                giPrevDiff = 1
            End If
            GetIdentLen = liTo - liFrom + 1
            giPrevToPos = liTo
            '1.0.0.4conn.Close()
            cmd = Nothing
            gServer.Connected = False
            If GetIdentLen = 0 Then
                Call Genericlog("GetIdentLen", "Fant ikke ident For " & pMethName & vbNewLine & lSql)
            End If
        Catch Ex As Exception
            Call GenericError("GetIdentLen", "Feil ved lesing av ident fra Oracle", Err.Number, Ex.Message, Err.Source)
            GetIdentLen = 0
        End Try
    End Function
End Module