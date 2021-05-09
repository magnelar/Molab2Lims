Option Strict Off
Option Explicit On
Module mLib
    'Molab2Lims
    Public Const g_sNoError As String = ""
    Public Const SEP_CHR As String = "~"

    Private Const KEY_ICON As String = "ICON"
    Private Const KEY_BMP As String = "BMP"
    Private Const KEY_CUR As String = "CUR"

    Public Const cDateFormat As String = "yyyy/MM/dd"
    Public Const cTimeFormat As String = "hh:nn:ss"
    Public Const cDateTimeFormat As String = cDateFormat & " " & cTimeFormat
    Public Const LIB_NOTOK As Short = -1
    Public Const LIB_OK As Short = 0


    Public Const INI_APPDATA As String = "AppData"
    Public Const INIKEY_StationName As String = "StationName"
    Public Const INIKEY_WORKDIR As String = "WorkDir"
    Public Const INIKEY_LOGDIR As String = "LogDir"
    Public Const INIKEY_DBLOGDIR As String = "DbLogDir"
    Public Const INIKEY_LANGUAGE As String = "Language"
    Public Const INIKEY_WEBSERVER As String = "WebServer"

    ' Windows API function and type declarations
    '
    ' Arrange Method for MDI Forms
    Public Const CASCADE As Short = 0
    Public Const TILE_HORIZONTAL As Short = 1
    Public Const TILE_VERTICAL As Short = 2
    Public Const ARRANGE_ICONS As Short = 3

    '  SetWindowPos Flags
    Public Const SWP_NOSIZE As Integer = &H1
    Public Const SWP_NOMOVE As Integer = &H2
    Public Const SWP_NOZORDER As Integer = &H4
    Public Const SWP_NOREDRAW As Integer = &H8
    Public Const SWP_NOACTIVATE As Integer = &H10
    Public Const SWP_DRAWFRAME As Integer = &H20
    Public Const SWP_SHOWWINDOW As Integer = &H40
    Public Const SWP_HIDEWINDOW As Integer = &H80
    Public Const SWP_NOCOPYBITS As Integer = &H100
    Public Const SWP_NOREPOSITION As Integer = &H200
    Public Const HWND_TOP As Short = 0
    Public Const HWND_BOTTOM As Short = 1
    Public Const HWND_TOPMOST As Short = -1
    Public Const HWND_NOTOPMOST As Short = -2

    '  Window field offsets for GetWindowLong() and GetWindowWord()
    Public Const GWL_WNDPROC As Short = (-4)
    Public Const GWW_HINSTANCE As Short = (-6)
    Public Const GWW_HWNDPARENT As Short = (-8)
    Public Const GWW_ID As Short = (-12)
    Public Const GWL_STYLE As Short = (-16)
    Public Const GWL_EXSTYLE As Short = (-20)

    ' edit control styles
    Public Const ES_AUTOVSCROLL As Integer = &H40
    Public Const ES_AUTOHSCROLL As Integer = &H80
    Public Const ES_READONLY As Integer = &H800
    Public Const ES_CENTER As Integer = &H1

    'Window messages
    Public Const WM_USER As Integer = &H400
    Public Const WM_WINDOWPOSCHANGED As Integer = &H47
    Public Const WM_MOVE As Integer = &H3
    Public Const WM_NCLBUTTONDOWN As Integer = &HA1
    Public Const WM_NCLBUTTONUP As Integer = &HA2

    'WMTRAP.VBX stuff
    Public Const WMTRAP_PREPROCESS As Short = -1
    Public Const WMTRAP_EATMESSAGE As Short = 0
    Public Const WMTRAP_POSTPROCESS As Short = 1




    '/* WinHelp() commands */

    Public Const HELP_CONTEXT As Integer = &H1 'Display topic in ulTopic
    Public Const HELP_QUIT As Integer = &H2 'Terminate help
    Public Const HELP_INDEX As Integer = &H3 'Display index
    Public Const HELP_CONTENTS As Integer = &H3 'Display Contents for help
    Public Const HELP_HELPONHELP As Integer = &H4 'Display help on using help
    Public Const HELP_SETINDEX As Integer = &H5 'Set the current Index for multiindex help
    Public Const HELP_FORCEFILE As Integer = &H9 'Ensure that help displays correct file
    Public Const HELP_KEY As Integer = &H101 'Display topic for keyword in dwData
    Public Const HELP_COMMAND As Integer = &H102 'Execute a Winhelp macro
    Public Const HELP_PARTIALKEY As Integer = &H105 'Call the Windows help search engine
    Public Const HELP_MULTIKEY As Integer = &H201

    Public Enum enuLOG_TYPE
        LOG_TYPE_UNDEFINED = 0
        LOG_TYPE_STATUS = 1
        LOG_TYPE_ERROR = 2
        LOG_TYPE_WARNING = 3
        LOG_TYPE_EDIT = 4
        LOG_TYPE_VBERROR = 5
        LOG_TYPE_DBERROR = 6
    End Enum

    Public Enum enuPRIORITY
        PRIORITY_UNDEFINED = 0
        PRIORITY_LOW = 1
        PRIORITY_MEDIUM = 2
        PRIORITY_HIGH = 3
    End Enum

    Public Enum enuSHOW_TYPE
        SHOW_TYPE_NONE = 0
        SHOW_TYPE_SHOW = 1
    End Enum


    Structure RECT
        'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim Left_Renamed As Short
        Dim Top As Short
        'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim Right_Renamed As Short
        Dim bottom As Short
    End Structure

    Public Structure POINTAPI
        Dim x As Integer
        Dim y As Integer
    End Structure


    Structure OSVERSIONINFO
        Dim dwOSVersionInfoSize As Integer
        Dim dwMajorVersion As Integer
        Dim dwMinorVersion As Integer
        Dim dwBuildNumber As Integer
        Dim dwPlatformId As Integer
        <VBFixedString(128), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=128)> Public szCSDVersion() As Char '  Maintenance string for PSS usage
    End Structure

    Structure CRITICAL_SECTION
        Dim Reserved1 As Integer
        Dim Reserved2 As Integer
        Dim Reserved3 As Integer
        Dim Reserved4 As Integer
        Dim Reserved5 As Integer
        Dim Reserved6 As Integer
    End Structure

    Public Const LANG_NORWEGIAN As Integer = &H14
    Public Const LANG_GERMAN As Integer = &H7
    Public Const LANG_ENGLISH As Integer = &H9
    Public Const SUBLANG_NORWEGIAN_BOKMAL As Integer = &H1
    Public Const SUBLANG_GERMAN As Integer = &H1
    Public Const SUBLANG_ENGLISH_US As Integer = &H1

    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer

    Declare Function GetDesktopWindow Lib "user32" () As Integer
    Declare Function GetWindow Lib "user32" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
    Declare Function GetWindowLW Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
    Declare Function GetParent Lib "user32" (ByVal hwnd As Integer) As Integer
    Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
    Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Integer, ByVal lpString As String, ByVal cch As Integer) As Integer
    Public Declare Function GetCursorPosAPI Lib "user32" Alias "GetCursorPos" (ByRef lpPoint As POINTAPI) As Integer
    Declare Function GetWindowLongAPI Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
    Declare Function SetWindowLongAPI Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer

    Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, ByRef nSize As Integer) As Integer
    Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, ByRef nSize As Integer) As Integer
    Declare Function GetVersion Lib "Kernel32" () As Integer

    Declare Function SetThreadLocale Lib "Kernel32" (ByVal Locale As Integer) As Integer
    Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Integer)


    Public Const GWL_ID As Short = (-12)
    Public Const GW_HWNDNEXT As Short = 2
    Public Const GW_CHILD As Short = 5

    Private nMousePointerCounter As Short
    Private nSavedMousePointer As Short




    '3.0.5 Private cAppLog As ItsComponents.iAppLog
    Private CHR39 As String
    Private m_sAppLog As String
    Private g_sDbLogDir As String
    Private g_ilsImageList As System.Windows.Forms.ImageList


    Public gfrmMain As System.Windows.Forms.Form

    Public g_sNameOfUser As String
    Public g_sComputerName As String
    Public g_sAppName As String
    Public g_mdtAppStartTime As Date
    Public g_sStationName As String
    Public g_intAppStop As Short
    Public g_sAppIniFile As String



    'Sub Lib_Init(frmMain As Form)
    Public Sub Lib_Init()
        On Error GoTo ErrorL

        gfrmMain = Nothing
        g_mdtAppStartTime = Now

        CHR39 = Chr(39)
        g_ilsImageList = Nothing
        g_sAppIniFile = My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".INI"

        Call Lib_NameOfUser(g_sNameOfUser)
        If Len(g_sNameOfUser) > 10 Then
            g_sNameOfUser = Left(g_sNameOfUser, 10)
        End If
        Call Lib_NameOfPC(g_sComputerName)
        g_sAppName = My.Application.Info.AssemblyName
        Exit Sub
ErrorL:
        Call GenericError("mLib", "Lib_Init", Err.Number, Err.Description, Err.Source)
    End Sub

    '3.0.5 Public Sub Lib_Terminate()
    '  Set cAppLog = Nothing
    'End Sub


    '    Public Function Lib_MsgBox(ByRef Msg As String, ByRef Button As Short, ByRef Title As String) As Short
    '        On Error GoTo ErrorL
    '        System.Windows.Forms.Cursor.Current = Lib_ResetMousePointer()
    '        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    '        Lib_MsgBox = MsgBox(Msg, Button, Title)
    '        Exit Function

    'ErrorL:
    '        System.Windows.Forms.Cursor.Current = Lib_ResetMousePointer()
    '        Call GenericError("mLib", "Lib_MsgBox", Err.Number, Err.Description, Err.Source)
    '    End Function


    Public Sub Lib_SetHourglass()
        System.Windows.Forms.Application.DoEvents() 'cause forms to repaint before going on
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
    End Sub



    Public Function Lib_ConvertStringToDouble(ByRef sNumber As String, ByRef dOut As Object) As Boolean
        On Error GoTo ErrorL

        ' Function converts a string to a value of type double.
        ' Unlike VBs CDbl, this function can comnvert strings containing
        ' both "." and ",". Empty string is returned as 0.
        ' Input:    String to be converted.
        ' Returns:  True if ok

        Dim i As Short
        Dim sSep, sSearchSep As String

        On Error GoTo Lib_ConvertStringToDoubleError

        If sNumber = "" Then
            Lib_ConvertStringToDouble = False
        Else
            sSep = Format(0.1, ".")
            If InStr(sSep, ",") Then
                sSearchSep = ","
            Else
                sSearchSep = "."
            End If
            i = iInStrBackwards(sSearchSep, sNumber)
            If i > 0 Then
                Mid(sNumber, i, 1) = "."
            End If
            If IsNumeric(sNumber) = True Then
                dOut = CDbl(Val(sNumber))
                Lib_ConvertStringToDouble = True
            Else
                Lib_ConvertStringToDouble = False
            End If
        End If
        Exit Function

Lib_ConvertStringToDoubleError:
        Lib_ConvertStringToDouble = False
        Exit Function

ErrorL:
        '3.0.5Lib_VBError "mLib", "Lib_ConvertStringToDouble", Erl, ""
        Call GenericError("mLib", "FLib_ConvertStringToDouble", Err.Number, Err.Description, Err.Source)

    End Function


    Public Function DecimalDelimiter() As String
        On Error GoTo ErrorL
        DecimalDelimiter = Mid(Format(1.1, "General Number"), 2, 1)
        Exit Function

ErrorL:
        Call GenericError("mLib", "DecimalDelimiter", Err.Number, Err.Description, Err.Source)
        '3.0.5 Lib_VBError "mLib", "DecimalDelimiter", Erl, ""
    End Function

    Public Function iInStrBackwards(ByRef sMatch As String, ByRef s As String, Optional ByRef iStart As Object = Nothing) As Short

        ' Function searches for an occurence of a substring in a source string, backwards.
        ' Input:    sMatch      String to find.
        '           s           String to search.
        '           iStart      Start position in s.
        ' Returns:  Position in s if found, else 0
        ' Created:  04.11.96 SÅ
        ' Updated:  25.06.97 SÅ/AV

        Dim iLast, i, iBegin, iMatch As Short

        On Error GoTo Error_Renamed

        If IsNothing(iStart) Then
            iBegin = Len(s)
        Else
            iBegin = iStart
        End If
        iLast = iBegin
        For i = iBegin To 1 Step -1
            iMatch = InStr(Mid(s, i, iLast - i + 1), sMatch)
            If iMatch > 0 Then
                iInStrBackwards = i + iMatch - 1
                Exit Function
            End If
        Next i
        iInStrBackwards = 0
        Exit Function
Error_Renamed:
        iInStrBackwards = 0
        Call GenericError("mLib", "iInStrBackwards", Err.Number, Err.Description, Err.Source)
        'Lib_VBError "mLib", "iInStrBackwards", Erl, ""
    End Function


    Public Function ModifyDecimalDelimiter(ByVal InputString_Renamed As String) As String
        On Error GoTo ErrorL
        'Modify decimal delimiter to local settings in number expression given in InputString

        Dim SearchDelimit As String
        Dim ReplaceByDelimit As String
        Dim Pos As Short
        Dim temp As String

        temp = InputString_Renamed

        ReplaceByDelimit = DecimalDelimiter()

        If ReplaceByDelimit = "." Then SearchDelimit = "," Else SearchDelimit = "."

        Pos = InStr(1, temp, SearchDelimit)
        If Pos <> 0 Then
            Mid(temp, Pos, 1) = ReplaceByDelimit
            If Not IsNumeric(temp) Then
                'Inputstring is still not numeric, so reverse operation
                ModifyDecimalDelimiter = InputString_Renamed
            Else
                ModifyDecimalDelimiter = temp
            End If
        Else
            'Already have correct decimal delimiter, or does not have at all
            ModifyDecimalDelimiter = InputString_Renamed
        End If
        Exit Function

ErrorL:
        '3.0.5 Lib_VBError "mLib", "ModifyDecimalDelimiter", Erl, ""
        Call GenericError("mLib", "ModifyDecimalDelimiter", Err.Number, Err.Description, Err.Source)

    End Function


    Public Function Lib_MakeLangId(ByRef nLanguage As Short, ByRef nSubLuanguage As Short) As Short
        On Error GoTo ErrorL

        Lib_MakeLangId = ((nSubLuanguage * 1024) + nLanguage)

        Exit Function
ErrorL:
        Call GenericError("mLib", "SetLanguage", Err.Number, Err.Description, Err.Source)
    End Function


    Public Sub Lib_SetLanguage()
        On Error GoTo ErrorL
        Dim Ret As Integer
        Dim sLanguage As String = String.Empty '5.10

        '5.1.0 sLanguage = LCase(Trim(Lib_GetINIstr(INI_APPDATA, INIKEY_LANGUAGE, "norwegian", g_sAppIniFile)))
        'App.HelpFile = App.Path + "\" + App.EXEName + ".HLP"

        Select Case sLanguage
            Case "english", "e"
                Ret = SetThreadLocale(Lib_MakeLangId(LANG_ENGLISH, SUBLANG_ENGLISH_US))

            Case Else
                Ret = SetThreadLocale(Lib_MakeLangId(LANG_NORWEGIAN, SUBLANG_NORWEGIAN_BOKMAL))
        End Select
        Exit Sub

ErrorL:
        Call GenericError("mLib", "SetLanguage", Err.Number, Err.Description, Err.Source)
    End Sub


    Public Function Lib_LoadResString(ByRef Index As Short, Optional ByRef lngFileId As Integer = 0, Optional ByRef strDefault As String = "") As String
        On Error GoTo ErrorL
        If Index = 0 Then
            'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
            If IsNothing(strDefault) Then
                Lib_LoadResString = "Not Defined"
            Else
                Lib_LoadResString = strDefault
            End If
        Else
            Lib_LoadResString = My.Resources.ResourceManager.GetString("str" + CStr(lngFileId + Index))
        End If
        Exit Function
ErrorL:
        Call GenericError("mLib", "Lib_LoadResString", Err.Number, Err.Description, Err.Source)
        On Error Resume Next
        'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
        If IsNothing(strDefault) Then
            Lib_LoadResString = "Not Defined"
        Else
            Lib_LoadResString = strDefault
        End If
    End Function




    Public Sub Lib_SetImageList(ByRef ImageList As System.Windows.Forms.ImageList)
        On Error GoTo ErrorL
        g_ilsImageList = ImageList
        Exit Sub

ErrorL:
        Call GenericError("mLib", "Lib_SetImageList", Err.Number, Err.Description, Err.Source)
    End Sub

    Public Function Lib_Right(ByRef s As String, ByRef l As Short) As Object
        On Error GoTo ErrorL
        If Len(s) > l Then
            Lib_Right = Right(s, l)
        Else
            Lib_Right = s
        End If
        Exit Function

ErrorL:
        Call GenericError("mLib", "Lib_Right", Err.Number, Err.Description, Err.Source)
    End Function




    Public Function Lib_GetMainForm() As System.Windows.Forms.Form
        On Error GoTo ErrorL
        Lib_GetMainForm = gfrmMain
        Exit Function
ErrorL:
        Call GenericError("mLib", "Lib_GetMainForm", Err.Number, Err.Description, Err.Source)
    End Function


    Public Function Lib_StrBlankIfZero(ByRef Value As Object) As String
        On Error GoTo ErrorL
        If Value = 0 Then
            Lib_StrBlankIfZero = ""
        Else
            Lib_StrBlankIfZero = CStr(Value)
        End If
        Exit Function
ErrorL:
        Call GenericError("mLib", "Lib_StrBlankIfZero", Err.Number, Err.Description, Err.Source)
    End Function



    Public Function Lib_GetDiffSecond(ByRef StopDate As Object, ByRef StartDate As Object) As Integer
        On Error GoTo ErrorL
        Dim Daydiff As Object
        Dim HourDiff As Object
        Dim MinuteDiff As Object
        Dim SecondDiff As Object ' Declare variables
        Dim TotalMinDiff As Object
        Dim TotalSecDiff As Object
        Dim d1 As Date
        Dim d2 As Date

        If IsNothing(StopDate) Or Len(StopDate) = 0 Then
            StopDate = Now
        End If
        If IsNothing(StartDate) Or Len(StartDate) = 0 Then
            StartDate = Now
        End If
        If IsDate(StopDate) And IsDate(StartDate) Then
            Daydiff = System.DateTime.FromOADate(CDate(StopDate).ToOADate - CDate(StartDate).ToOADate)
            If CInt(Daydiff) < 10000 Then
                TotalSecDiff = CInt(Daydiff * 86400)
                Lib_GetDiffSecond = CInt(TotalSecDiff)
            Else
                Lib_GetDiffSecond = 999999
            End If
        Else
            Lib_GetDiffSecond = 999999
        End If
        Exit Function

ErrorL:
        Call GenericError("mLib", "Lib_GetDiffSecond", Err.Number, Err.Description, Err.Source)
    End Function



    Public Sub Lib_SplitFilename(ByRef FilePathName As String, ByRef Drive As String, ByRef Path As String, ByRef Filename As String)
        On Error GoTo ErrorL
        Dim i As Short
        Dim p As Short
        Dim d As Short

        i = 1
        p = 0
        d = InStr(i, FilePathName, ":")
        If d Then
            Drive = Left(FilePathName, d)
            i = d + 1
        End If
        While i <> 0
            i = InStr(i, FilePathName, "\")
            If i <> 0 Then
                p = i
                i = i + 1
            End If
        End While
        If p Then
            Filename = Right(FilePathName, Len(FilePathName) - p)
            Path = Mid(FilePathName, d + 1, p - d)
        Else
            Filename = FilePathName
            Path = ""
        End If
        Exit Sub
ErrorL:
        Call GenericError("mLib", "Lib_SplitFilename", Err.Number, Err.Description, Err.Source)
    End Sub


    Public Function Lib_InitMousePointer() As Boolean
        On Error GoTo ErrorL
        nMousePointerCounter = 0
        Lib_InitMousePointer = True
        Exit Function

ErrorL:
        Lib_InitMousePointer = False
        Call GenericError("mLib", "Lib_InitMousePointer", Err.Number, Err.Description, Err.Source)
    End Function


    Public Function Lib_ResetMousePointer() As Cursor '5.1.0
        On Error GoTo ErrorL
        nMousePointerCounter = 0
        Lib_ResetMousePointer = System.Windows.Forms.Cursors.Default
        Exit Function

ErrorL:
        Lib_ResetMousePointer = System.Windows.Forms.Cursors.Default
        Call GenericError("mLib", "Lib_ResetMousePointer", Err.Number, Err.Description, Err.Source)
    End Function


    Public Function Lib_GetArrow() As Cursor '5.1.0
        On Error GoTo ErrorL
        nMousePointerCounter = nMousePointerCounter - 1
        If nMousePointerCounter < 1 Then
            Lib_GetArrow = System.Windows.Forms.Cursors.Default
            nMousePointerCounter = 0
        Else
            Lib_GetArrow = System.Windows.Forms.Cursors.WaitCursor
        End If
        Exit Function

ErrorL:
        Lib_GetArrow = System.Windows.Forms.Cursors.Default
        Call GenericError("mLib", "Lib_GetArrow", Err.Number, Err.Description, Err.Source)
    End Function


    Public Function Lib_GetHourGlass() As Cursor '5.1.0
        On Error GoTo ErrorL
        nMousePointerCounter = nMousePointerCounter + 1
        Lib_GetHourGlass = System.Windows.Forms.Cursors.WaitCursor
        Exit Function

ErrorL:
        Lib_GetHourGlass = System.Windows.Forms.Cursors.WaitCursor
        Call GenericError("mLib", "Lib_GetHourGlass", Err.Number, Err.Description, Err.Source)
    End Function


    Public Sub Lib_SetWindowPosTop(ByRef hwnd As Integer)
        On Error GoTo ErrorL
        SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
        Exit Sub
ErrorL:
        Call GenericError("mLib", "Lib_SetWindowPosTop", Err.Number, Err.Description, Err.Source)
    End Sub


    Public Sub Lib_SetWindowPosTopMost(ByRef hwnd As Integer)
        On Error GoTo ErrorL
        SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
        Exit Sub
ErrorL:
        Call GenericError("mLib", "Lib_SetWindowPosTopMost", Err.Number, Err.Description, Err.Source)
    End Sub


    Public Sub Lib_SetWindowPosNotTopMost(ByRef hwnd As Integer)
        On Error GoTo ErrorL
        SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
        Exit Sub
ErrorL:
        Call GenericError("mLib", "Lib_SetWindowPosNotTopMost", Err.Number, Err.Description, Err.Source)
    End Sub


    Public Sub Lib_SetWindowPosBottom(ByRef hwnd As Integer)
        On Error GoTo ErrorL
        SetWindowPos(hwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

        Exit Sub
ErrorL:
        Call GenericError("mLib", "Lib_SetWindowPosBottom", Err.Number, Err.Description, Err.Source)
    End Sub


    Public Sub Lib_SetWindowPosAfter(ByRef hwnd As Integer, ByRef hWndInsertAfter As Integer)
        On Error GoTo ErrorL
        SetWindowPos(hwnd, hWndInsertAfter, 0, 0, 0, 0, &H1 + &H2)
        Exit Sub
ErrorL:
        Call GenericError("mLib", "Lib_SetWindowPosAfter", Err.Number, Err.Description, Err.Source)
    End Sub


    Public Function SetTopMostWindow(ByRef hwnd As Integer, ByRef Topmost As Boolean) As Integer
        On Error GoTo ErrorL

        If Topmost = True Then 'Make the window topmost
            SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        Else
            SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            SetTopMostWindow = False
        End If

        Exit Function
ErrorL:
        Call GenericError("mLib", "SetTopMostWindow", Err.Number, Err.Description, Err.Source)
    End Function



    '5.1.0	Public Function Lib_WriteINIstr(ByVal Section As String, ByVal Entry As String, ByVal sValue As String, ByVal lplFileName As String) As Boolean
    '		On Error GoTo ErrorL
    '		Dim Ret As Short

    '		If Entry = "" Then
    '			Ret = WritePrivateProfileStringAPI(Section, 0, 0, lplFileName)
    '		Else
    '			If sValue = "" Then
    '				Ret = WritePrivateProfileStringAPI(Section, Entry, 0, lplFileName)
    '			Else
    '				Ret = WritePrivateProfileStringAPI(Section, Entry, sValue, lplFileName)
    '			End If
    '		End If
    '		Lib_WriteINIstr = (Ret <> 0)
    '		Exit Function
    'ErrorL: 
    '		Call GenericError("mLib", "Lib_WriteINIstr", Err.Number, Err.Description, Err.Source)
    '	End Function


    'UPGRADE_NOTE: Default was upgraded to Default_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    '    Public Function Lib_GetINIstr(ByVal Section As String, ByVal Entry As String, ByVal Default_Renamed As String, ByVal lplFileName As String) As String
    '        On Error GoTo ErrorL
    '        Dim Ret As Short
    '        Dim s As String
    '        Dim Pos As Short
    '        s = New String(Chr(0), 4096)

    '        If Entry = "" Then
    '            Ret = GetPrivateProfileStringAPI(Section, 0, Default_Renamed, s, Len(s), lplFileName)
    '        Else
    '            Ret = GetPrivateProfileStringAPI(Section, Entry, Default_Renamed, s, Len(s), lplFileName)
    '        End If

    '        If Ret > 0 Then
    '            If Entry <> "" Then
    '                Pos = InStr(s, Chr(0))
    '                If Pos > 1 Then
    '                    Lib_GetINIstr = Trim(Left(s, Pos - 1))
    '                Else
    '                    Lib_GetINIstr = ""
    '                End If
    '            Else
    '                Lib_GetINIstr = Trim(s)
    '            End If
    '        End If
    '        Exit Function
    'ErrorL:
    '        Call GenericError("mLib", "Lib_GetINIstr", Err.Number, Err.Description, Err.Source)
    '    End Function


    Public Function Lib_NameOfUser(ByRef UserName As String) As Integer
        On Error GoTo ErrorL
        Dim NameSize As Integer
        Dim x As Integer
        Dim verinfo As OSVERSIONINFO

        UserName = Space(16)
        NameSize = Len(UserName)
        x = GetUserName(UserName, NameSize)
        If x Then
            UserName = Lib_StringC2VB(UserName)
        Else
            UserName = vbNullString
        End If

        Exit Function
ErrorL:
        Call GenericError("mLib", "Lib_NameOfUser", Err.Number, Err.Description, Err.Source)
    End Function

    '  FindWindowLike
    '    '- Finds the window handles of the windows matching the specified
    '      Parameters
    '
    '   hwndArray()
    '   ' - An integer array used to return the window handles
    '
    '   hWndStart
    '    - The handle of the window to search under.
    '    - The routine searches through all of this window's children and their
    '      children recursively.
    '    - If hWndStart = 0 then the routine searches through all windows.
    '
    '   WindowText
    '    - The pattern used with the Like operator to compare window's text.
    '
    '   Classname
    '    - The pattern used with the Like operator to compare window's class
    '      name.
    '
    '   ID
    '    - A child ID number used to identify a window.
    '    - Can be a decimal number or a hex string.
    '    - Prefix hex strings with "&H" or an error will occur.
    '    - To ignore the ID pass the Visual Basic Null function.
    '
    '   Returns
    '    - The number of windows that matched the parameters.
    '    - Also returns the window handles in hWndArray()

    '----------------------------------------------------------------------
    'Remove this next line to use the strong-typed declarations
#Const WinVar = True
    Public Function Lib_FindWindowLike(ByRef hWndArray() As Integer, ByVal hWndStart As Integer, ByRef WindowText As String, ByRef Classname As String, ByRef ID As Object) As Integer
        On Error GoTo ErrorL

        Dim hwnd As Integer
        Dim R As Integer
        ' Hold the level of recursion:
        Static level As Integer
        ' Hold the number of matching windows:
        Static iFound As Integer
        Dim sWindowText As String
        Dim sClassname As String
        Dim sID As Object
        ' Initialize if necessary:
        If level = 0 Then
            iFound = 0
            ReDim hWndArray(0)
            If hWndStart = 0 Then hWndStart = GetDesktopWindow()
        End If
        ' Increase recursion counter:
        level = level + 1
        ' Get first child window:
        hwnd = GetWindow(hWndStart, GW_CHILD)
        Do Until hwnd = 0
            System.Windows.Forms.Application.DoEvents() ' Not necessary
            ' Search children by recursion:
            R = Lib_FindWindowLike(hWndArray, hwnd, WindowText, Classname, ID)
            ' Get the window text and class name:
            sWindowText = Space(255)
            R = GetWindowText(hwnd, sWindowText, 255)
            sWindowText = Left(sWindowText, R)
            sClassname = Space(255)
            R = GetClassName(hwnd, sClassname, 255)
            sClassname = Left(sClassname, R)
            ' If window is a child get the ID:
            If GetParent(hwnd) <> 0 Then
                R = GetWindowLW(hwnd, GWL_ID)
                sID = CInt("&H" & Hex(R))
            Else
                sID = System.DBNull.Value
            End If
            ' Check that window matches the search parameters:
            If sWindowText Like WindowText And sClassname Like Classname Then
                If IsDBNull(ID) Then
                    ' If find a match, increment counter and
                    '  add handle to array:
                    iFound = iFound + 1
                    ReDim Preserve hWndArray(iFound)
                    hWndArray(iFound) = hwnd
                ElseIf Not IsDBNull(sID) Then
                    If CInt(sID) = CInt(ID) Then
                        ' If find a match increment counter and
                        '  add handle to array:
                        iFound = iFound + 1
                        ReDim Preserve hWndArray(iFound)
                        hWndArray(iFound) = hwnd
                    End If
                End If
                Debug.Print("Window Found: ")
                Debug.Print("  Window Text  : " & sWindowText)
                Debug.Print("  Window Class : " & sClassname)
                Debug.Print("  Window Handle: " & CStr(hwnd))
            End If
            ' Get next child window:
            hwnd = GetWindow(hwnd, GW_HWNDNEXT)
        Loop
        ' Decrement recursion counter:
        level = level - 1
        ' Return the number of windows found:
        Lib_FindWindowLike = iFound
        Exit Function
ErrorL:
        Call GenericError("mLib", "Lib_FindWindowLike", Err.Number, Err.Description, Err.Source)
    End Function

    Public Function Lib_StringC2VB(ByVal Cstring As String) As String
        On Error GoTo ErrorL
        Dim Pos As Short

        Pos = InStr(Cstring, Chr(0))
        If Pos > 1 Then
            Lib_StringC2VB = Left(Cstring, Pos - 1)
        Else
            Lib_StringC2VB = ""
        End If
        Exit Function

ErrorL:
        Call GenericError("mLib", "Lib_StringC2VB", Err.Number, Err.Description, Err.Source)
    End Function


    Public Function Lib_NameOfPC(ByRef MachineName As String) As Integer
        On Error GoTo ErrorL

        Dim NameSize As Integer
        Dim x As Integer

        MachineName = Space(16)
        NameSize = Len(MachineName)
        x = GetComputerName(MachineName, NameSize)
        If x Then
            MachineName = Lib_StringC2VB(MachineName)
        Else
            MachineName = vbNullString
        End If
        Exit Function
ErrorL:
        Call GenericError("mLib", "Lib_NameOfPC", Err.Number, Err.Description, Err.Source)
    End Function


    Public Function Lib_Crypt(ByRef strPassword As String, ByRef strCryptKey As String) As String
        On Error GoTo ErrorL
        Dim intLength As Short
        Dim intCnt As Short
        Dim strChar As String
        Dim strEncrypted As String

        intLength = Len(strCryptKey)
        strEncrypted = strPassword
        For intCnt = 1 To Len(strEncrypted)
            strChar = CStr(Asc(Mid(strCryptKey, (intCnt Mod intLength) - intLength * CShort((intCnt Mod intLength) = 0), 1)))
            Mid(strEncrypted, intCnt, 1) = Chr(Asc(Mid(strEncrypted, intCnt, 1)) Xor strChar)
        Next
        Lib_Crypt = strEncrypted
        Exit Function
ErrorL:
        Call GenericError("mLib", "Lib_Crypt", Err.Number, Err.Description, Err.Source)
    End Function
End Module