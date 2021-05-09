Option Strict Off
Option Explicit On
Module mLib
	Public Const g_sNoError As String = ""
	Public Const SEP_CHR As String = "~"
	
	Private Const KEY_ICON As String = "ICON"
	Private Const KEY_BMP As String = "BMP"
	Private Const KEY_CUR As String = "CUR"
	
	Public Const cDateFormat As String = "yyyy/mm/dd"
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
		Dim X As Integer
		Dim y As Integer
	End Structure
	
	
	Structure OSVERSIONINFO
		Dim dwOSVersionInfoSize As Integer
		Dim dwMajorVersion As Integer
		Dim dwMinorVersion As Integer
		Dim dwBuildNumber As Integer
		Dim dwPlatformId As Integer
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(128),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=128)> Public szCSDVersion() As Char '  Maintenance string for PSS usage
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
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Function WritePrivateProfileStringAPI Lib "Kernel32"  Alias "WritePrivateProfileStringA"(ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Short
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Function GetPrivateProfileStringAPI Lib "Kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Short, ByVal lpFileName As String) As Short
	Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
	
	Declare Function GetDesktopWindow Lib "user32" () As Integer
	Declare Function GetWindow Lib "user32" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
	Declare Function GetWindowLW Lib "user32"  Alias "GetWindowLongA"(ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
	Declare Function GetParent Lib "user32" (ByVal hwnd As Integer) As Integer
	Declare Function GetClassName Lib "user32"  Alias "GetClassNameA"(ByVal hwnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
	Declare Function GetWindowText Lib "user32"  Alias "GetWindowTextA"(ByVal hwnd As Integer, ByVal lpString As String, ByVal cch As Integer) As Integer
	'UPGRADE_WARNING: Structure POINTAPI may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function GetCursorPosAPI Lib "user32"  Alias "GetCursorPos"(ByRef lpPoint As POINTAPI) As Integer
	Declare Function GetWindowLongAPI Lib "user32"  Alias "GetWindowLongA"(ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
	Declare Function SetWindowLongAPI Lib "user32"  Alias "SetWindowLongA"(ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
	Declare Function SendMessage Lib "user32"  Alias "SendMessageA"(ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
	
	Declare Function GetComputerName Lib "Kernel32"  Alias "GetComputerNameA"(ByVal lpbuffer As String, ByRef nSize As Integer) As Integer
	Declare Function GetUserName Lib "advapi32.dll"  Alias "GetUserNameA"(ByVal lpbuffer As String, ByRef nSize As Integer) As Integer
	Declare Function GetVersion Lib "Kernel32" () As Integer
	
	Declare Function SetThreadLocale Lib "Kernel32" (ByVal Locale As Integer) As Integer
	Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Integer)
	'Public Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
	'Public Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
	'Public Declare Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
	'Public Declare Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
	
	
	Public Const GWL_ID As Short = (-12)
	Public Const GW_HWNDNEXT As Short = 2
	Public Const GW_CHILD As Short = 5
	
	Private nMousePointerCounter As Short
	Private nSavedMousePointer As Short
	
	
	
	
	Private cAppLog As ItsComponents.iAppLog
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
		
		'UPGRADE_NOTE: Object gfrmMain may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		gfrmMain = Nothing
		g_mdtAppStartTime = Now
		
		CHR39 = Chr(39)
		'UPGRADE_NOTE: Object g_ilsImageList may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		g_ilsImageList = Nothing
		'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		g_sAppIniFile = My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".INI"
		g_sDbLogDir = Trim(Lib_GetINIstr(INI_APPDATA, INIKEY_DBLOGDIR, "c:\AppLog", g_sAppIniFile))
		
		Call Lib_NameOfUser(g_sNameOfUser)
		If Len(g_sNameOfUser) > 10 Then
			g_sNameOfUser = Left(g_sNameOfUser, 10)
		End If
		Call Lib_NameOfPC(g_sComputerName)
		'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		g_sAppName = My.Application.Info.AssemblyName
		cAppLog = New ItsApplog.cApplog
		Call cAppLog.AppLogInit(g_sDbLogDir, g_sComputerName, g_sAppName, My.Application.Info.DirectoryPath)
		Exit Sub
ErrorL: 
		Lib_VBError("mLib", "Lib_Init", Erl(), "")
	End Sub
	
	Public Sub Lib_Terminate()
		'UPGRADE_NOTE: Object cAppLog may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cAppLog = Nothing
	End Sub
	
	
	Public Function Lib_MsgBox(ByRef Msg As String, ByRef Button As Short, ByRef Title As String) As Short
		On Error GoTo ErrorL
		'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = Lib_ResetMousePointer()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Lib_MsgBox = MsgBox(Msg, Button, Title)
		Exit Function
		
ErrorL: 
		'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = Lib_ResetMousePointer()
		Lib_VBError("mLib", "Lib_MsgBox", Erl(), Msg, "")
	End Function
	
	
	'UPGRADE_NOTE: minute was upgraded to minute_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function Lib_Add2_Time(ByRef dtTime As Date, ByRef minute_Renamed As Short) As Date
		On Error GoTo ErrorL
		Lib_Add2_Time = DateAdd(Microsoft.VisualBasic.DateInterval.Minute, minute_Renamed, dtTime)
		Exit Function
ErrorL: 
		Lib_VBError("mLib", "Lib_Add2_Time", Erl(), "")
	End Function
	
	Public Sub lib_ShowAppLog()
		On Error GoTo ErrorL
		If Not cAppLog Is Nothing Then
			Call cAppLog.ShowAppLog()
		End If
		Exit Sub
ErrorL: 
		Lib_VBError("mLib", "mnuFileExit_Click", Erl(), "mdiRaffEdit")
	End Sub
	
	Public Sub Lib_DdeLog(ByVal JobName As String, ByVal GroupName As String, ByVal AliasTopic As String, ByVal AliasItem As String, ByVal TopicName As String, ByVal ItemName As String, ByVal ItemValue As String, ByVal Value_in As String, ByVal Qualitiy As Integer, ByVal GroupTime As Date, ByVal TimeStamp As Date)
		On Error GoTo ErrorL
		'Call Timer_Start
		If Not cAppLog Is Nothing Then
			Call cAppLog.DdeLog(JobName, GroupName, AliasTopic, AliasItem, TopicName, ItemName, ItemValue, Value_in, Qualitiy, GroupTime, TimeStamp)
		End If
		'Call Timer_Stop_ms("Lib_DdeLog")
		Exit Sub
ErrorL: 
		Lib_VBError("mLib", "mnuFileExit_Click", Erl(), "mdiRaffEdit")
	End Sub
	
	
	Public Sub Lib_AppLog(ByVal lngLogType As enuLOG_TYPE, ByVal lngPriority As enuPRIORITY, ByVal lngShowType As enuSHOW_TYPE, ByVal ModuleName As String, ByVal FunctionName As String, ByVal LineNo As Integer, ByVal User_Text1 As String, ByVal User_Text2 As String, Optional ByVal Err_ErrorNo As Integer = 0, Optional ByVal Err_Description As String = "", Optional ByVal Err_Source As String = "")
		On Error GoTo ErrorL
		Dim lReturn As Integer
		
		'Debug.Print "Lib_AppLog - Module: "; ModuleName & ", enuLOG_TYPE: " & Format(lngLogType) & " Function: " & FunctionName & "Info =" & User_Text1 & "  " & User_Text2
		If Not cAppLog Is Nothing Then
			lReturn = cAppLog.AppLog(VB6.Format(lngLogType), lngPriority, lngShowType, ModuleName, FunctionName, LineNo, User_Text1, User_Text2, Err_ErrorNo, Err_Description, Err_Source)
		End If
		Exit Sub
ErrorL: 
		Err.Clear()
	End Sub
	
	
	' TBD:  This is kept in for bakcward compatibility.  Will be removed as soon as all
	' client have converted to new Applog API.
	Public Sub Lib_LogAppError(ByRef LogType As String, ByRef FuncName As String, ByRef LineNo As Integer, ByRef Code As String, ByRef info As String, Optional ByRef ErrorCode As Integer = 0, Optional ByRef Errorinfo As String = "", Optional ByRef Err_Source As String = "")
		On Error GoTo ErrorL
		Call Lib_AppLog(CShort(LogType), enuPRIORITY.PRIORITY_UNDEFINED, enuSHOW_TYPE.SHOW_TYPE_NONE, "", FuncName, LineNo, Code, info, ErrorCode, Errorinfo, "")
		Exit Sub
ErrorL: 
	End Sub
	
	' TBD:  This is kept in for bakcward compatibility.  Will be removed as soon as all
	' client have converted to new Applog API.
	Public Sub Lib_ShowError(ByRef strFuncName As String, ByRef lngLineNo As Integer, ByRef strUser_Text1 As String, Optional ByRef strUser_Text2 As String = "", Optional ByRef blnShowMsgBox As Boolean = True)
		On Error GoTo ErrorL
		Call Lib_VBError(strUser_Text1, strFuncName, lngLineNo, strUser_Text1, strUser_Text2, blnShowMsgBox)
		Exit Sub
ErrorL: 
	End Sub
	
	
	Public Sub Lib_VBError(ByVal strModuleName As String, ByRef strFuncName As String, ByRef lngLineNo As Integer, ByRef strUser_Text1 As String, Optional ByRef strUser_Text2 As String = "", Optional ByRef blnShowMsgBox As Boolean = True)
		Dim lngErrNo As Integer
		Dim strFirstError As String
		Dim strFuncNum As Short
		Dim strErrStr As String
		Dim strErrSource As String
		Dim lngReturn As Integer
		Dim lngShowType As enuSHOW_TYPE
		
		
		lngErrNo = Err.Number
		strErrStr = Err.Description
		strErrSource = Err.Source
		If blnShowMsgBox Then
			lngShowType = enuSHOW_TYPE.SHOW_TYPE_SHOW
		Else
			lngShowType = enuSHOW_TYPE.SHOW_TYPE_NONE
		End If
		
		On Error Resume Next
		Debug.Print("Lib_VBError - Module: " & strModuleName & ", Function: " & strFuncName & ", Description: " & ErrorToString(lngErrNo)) ' Print error to Debug window.
		If lngErrNo = 0 Then
			Exit Sub
		End If
		If lngErrNo = 3021 Then
			lngShowType = enuSHOW_TYPE.SHOW_TYPE_NONE
		End If
		Call Lib_AppLog(enuLOG_TYPE.LOG_TYPE_VBERROR, enuPRIORITY.PRIORITY_HIGH, lngShowType, strModuleName, strFuncName, lngLineNo, strUser_Text1, strUser_Text2, lngErrNo, strErrStr, strErrSource)
		Err.Number = lngErrNo
	End Sub
	
	'------------------------------------------------------------
	'this sub sets the HourGlass icon for the mouse
	'------------------------------------------------------------
	Public Sub Lib_SetHourglass()
		System.Windows.Forms.Application.DoEvents() 'cause forms to repaint before going on
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
	End Sub
	
	
	
	Public Function Lib_ConvertStringToDouble(ByRef sNumber As String, ByRef dOut As Object) As Boolean
		On Error GoTo ErrorL
		
		' Function converts a string to a value of type double.
		' Unlike VBs CDbl, this function can comnvert strings containing
		' both "." and ",". Empty string is returned as 0.
		' Input:    String to be converted.
		' Returns:  True if ok
		'1.7.5 bør vel fungere med komma som desimal streng.
		Dim i As Short
		Dim sSep, sSearchSep As String
		
		On Error GoTo Lib_ConvertStringToDoubleError
		
		If sNumber = "" Then
			Lib_ConvertStringToDouble = False
		Else
			sSep = VB6.Format(0.1, ".")
			If InStr(sSep, ",") Then
				sSearchSep = ","
			Else
				sSearchSep = ","
			End If
			i = iInStrBackwards(sSearchSep, sNumber)
			If i > 0 Then
				Mid(sNumber, i, 1) = "."
			End If
			If IsNumeric(sNumber) = True Then
				'UPGRADE_WARNING: Couldn't resolve default property of object dOut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		Lib_VBError("mLib", "Lib_ConvertStringToDouble", Erl(), "")
	End Function
	
	
	Public Function DecimalDelimiter() As String
		On Error GoTo ErrorL
		DecimalDelimiter = Mid(VB6.Format(1.1, "General Number"), 2, 1)
		Exit Function
		
ErrorL: 
		Lib_VBError("mLib", "DecimalDelimiter", Erl(), "")
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
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(iStart) Then
			iBegin = Len(s)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object iStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		Lib_VBError("mLib", "iInStrBackwards", Erl(), "")
	End Function
	
	
	'UPGRADE_NOTE: InputString was upgraded to InputString_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function ModifyDecimalDelimiter(ByVal InputString_Renamed As String) As String
		On Error GoTo ErrorL
		'Modify decimal delimiter to local settings in number expression given in InputString
		
		Dim SearchDelimit As String
		Dim ReplaceByDelimit As String
		Dim Pos As Short
		Dim temp As String
		
		temp = InputString_Renamed
		
		ReplaceByDelimit = DecimalDelimiter
		
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
		Lib_VBError("mLib", "ModifyDecimalDelimiter", Erl(), "")
	End Function
	
	
	Public Function Lib_GetField(ByRef s As String, ByRef Pos As Object) As Object
		On Error GoTo ErrorL
		Dim i As Short
		Dim x1 As Short
		Dim x2 As Short
		x2 = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object Pos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For i = 1 To Pos
			x1 = x2 + 1
			x2 = InStr(x1, s, SEP_CHR)
		Next 
		If x1 > 0 And x2 > 0 Then
			Lib_GetField = Trim(Mid(s, x1, x2 - x1))
		ElseIf x1 > 0 Then 
			Lib_GetField = Trim(Right(s, Len(s) - x1 + 1))
		End If
		Exit Function
		
ErrorL: 
		Lib_VBError("mLib", "Lib_GetField", Erl(), "")
	End Function
	
	
	Public Function Lib_WriteVar(ByVal Section As String, ByVal Entry As String, ByVal lplFileName As String, ByRef nParam As Short, ByRef v1 As Object, Optional ByRef v2 As Object = 0, Optional ByRef v3 As Object = Nothing, Optional ByRef v4 As Object = Nothing) As Boolean
		On Error GoTo ErrorL
		Lib_WriteVar = False
		Select Case nParam
			Case 1
				'UPGRADE_WARNING: Couldn't resolve default property of object v1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Lib_WriteVar = Lib_WriteINIstr(Section, Entry, VB6.Format(v1), lplFileName)
			Case 2
				'UPGRADE_WARNING: Couldn't resolve default property of object v2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object v1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Lib_WriteVar = Lib_WriteINIstr(Section, Entry, VB6.Format(v1) & SEP_CHR & VB6.Format(v2), lplFileName)
			Case 3
				'UPGRADE_WARNING: Couldn't resolve default property of object v3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object v2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object v1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Lib_WriteVar = Lib_WriteINIstr(Section, Entry, VB6.Format(v1) & SEP_CHR & VB6.Format(v2) & SEP_CHR & VB6.Format(v3), lplFileName)
			Case 4
				'UPGRADE_WARNING: Couldn't resolve default property of object v4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object v3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object v2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object v1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Lib_WriteVar = Lib_WriteINIstr(Section, Entry, VB6.Format(v1) & SEP_CHR & VB6.Format(v2) & SEP_CHR & VB6.Format(v3) & SEP_CHR & VB6.Format(v4), lplFileName)
		End Select
		Exit Function
		
ErrorL: 
		Lib_VBError("mLib", "Lib_WriteVar", Erl(), "")
	End Function
	
	
	
	Public Function Lib_GetVar(ByVal Section As String, ByVal Entry As String, ByVal lplFileName As String, ByRef nParam As Short, ByRef v1 As Object, ByRef d1 As Object, Optional ByRef v2 As Object = Nothing, Optional ByRef d2 As Object = Nothing, Optional ByRef v3 As Object = Nothing, Optional ByRef d3 As Object = Nothing, Optional ByRef v4 As Object = Nothing, Optional ByRef d4 As Object = Nothing) As Boolean
		On Error GoTo ErrorL
		Dim s As String
		Dim sV(4) As String
		Dim v(4) As Object
		Dim i As Short
		Dim Vt(4) As VariantType
		Dim ok As Boolean
		
		Lib_GetVar = False
		If nParam < 1 Or nParam > 4 Then Exit Function
		
		Lib_GetVar = True
		Select Case nParam
			Case 1
				'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				Vt(1) = VarType(v1)
			Case 2
				'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				Vt(1) = VarType(v1)
				'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				Vt(2) = VarType(v2)
			Case 3
				'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				Vt(1) = VarType(v1)
				'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				Vt(2) = VarType(v2)
				'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				Vt(3) = VarType(v3)
			Case 4
				'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				Vt(1) = VarType(v1)
				'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				Vt(2) = VarType(v2)
				'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				Vt(3) = VarType(v3)
				'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				Vt(4) = VarType(v4)
		End Select
		
		s = Trim(Lib_GetINIstr(Section, Entry, "", lplFileName))
		For i = 1 To nParam
			'UPGRADE_WARNING: Couldn't resolve default property of object Lib_GetField(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sV(i) = Lib_GetField(s, i)
			Select Case Vt(i)
				Case VariantType.Short
					'UPGRADE_WARNING: Couldn't resolve default property of object v(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v(i) = CShort(Val(sV(i)))
				Case VariantType.Double
					ok = Lib_ConvertStringToDouble(sV(i), v(i))
					If ok = False Then
						'UPGRADE_WARNING: Couldn't resolve default property of object v(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						v(i) = 0 'js99
					End If
				Case Else
					'UPGRADE_WARNING: Couldn't resolve default property of object v(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v(i) = sV(i)
			End Select
		Next i
		Select Case nParam
			Case 1
				If sV(1) = "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object d1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v1 = d1
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object v(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v1 = v(1)
				End If
			Case 2
				If sV(1) = "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object d1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v1 = d1
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object v(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v1 = v(1)
				End If
				If sV(2) = "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object d2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v2 = d2
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object v(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v2 = v(2)
				End If
			Case 3
				If sV(1) = "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object d1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v1 = d1
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object v(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v1 = v(1)
				End If
				If sV(2) = "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object d2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v2 = d2
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object v(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v2 = v(2)
				End If
				If sV(3) = "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object d3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v3 = d3
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object v(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v3 = v(3)
				End If
			Case 4
				If sV(1) = "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object d1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v1 = d1
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object v(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v1 = v(1)
				End If
				If sV(2) = "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object d2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v2 = d2
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object v(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v2 = v(2)
				End If
				If sV(3) = "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object d3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v3 = d3
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object v(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v3 = v(3)
				End If
				If sV(4) = "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object d4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v4 = d4
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object v(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object v4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					v4 = v(4)
				End If
		End Select
		Exit Function
		
ErrorL: 
		Lib_VBError("mLib", "Lib_GetVar", Erl(), "")
	End Function
	
	
	' This function does the same as:
	' #define MAKELANGID(p, s)       ((((WORD  )(s)) << 10) | (WORD  )(p))
	'
	Public Function Lib_MakeLangId(ByRef nLanguage As Short, ByRef nSubLuanguage As Short) As Short
		On Error GoTo ErrorL
		
		Lib_MakeLangId = ((nSubLuanguage * 1024) + nLanguage)
		
		Exit Function
ErrorL: 
		Lib_VBError("mLib", "SetLanguage", Erl(), "")
	End Function
	
	
	Public Sub Lib_SetLanguage()
		On Error GoTo ErrorL
		Dim Ret As Integer
		Dim sLanguage As String
		
		sLanguage = LCase(Trim(Lib_GetINIstr(INI_APPDATA, INIKEY_LANGUAGE, "norwegian", g_sAppIniFile)))
		'App.HelpFile = App.Path + "\" + App.EXEName + ".HLP"
		
		Select Case sLanguage
			Case "english", "e"
				Ret = SetThreadLocale(Lib_MakeLangId(LANG_ENGLISH, SUBLANG_ENGLISH_US))
				'            If (gKeyCodeOK And GetOption2Bits(OPTION2_LANGUAGE)) Or gKeyCodeOK = False Then
				'                Ret = SetThreadLocale(MAKELANGID(LANG_NORWEGIAN, SUBLANG_NORWEGIAN_BOKMAL))
				'                App.HelpFile = App.Path + "\" + App.EXEName + "DE.HLP"
				'            Else
				'                Ret = SetThreadLocale(MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US))
				'            End If
				
			Case Else
				Ret = SetThreadLocale(Lib_MakeLangId(LANG_NORWEGIAN, SUBLANG_NORWEGIAN_BOKMAL))
		End Select
		Exit Sub
		
ErrorL: 
		Lib_VBError("mLib", "SetLanguage", Erl(), "")
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
		Lib_VBError("mLib", "Lib_LoadResString", Erl(), "ID=" & VB6.Format(Index) & " FileId=" & VB6.Format(lngFileId))
		On Error Resume Next
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(strDefault) Then
			Lib_LoadResString = "Not Defined"
		Else
			Lib_LoadResString = strDefault
		End If
	End Function
	
	
	
	Public Function Lib_IsItemInList(ByRef TheList As System.Windows.Forms.Control, ByVal Item As String, ByRef lastindex As Short) As Short
		On Error GoTo ErrorL
		
		Dim i As Short
		Dim SearchString As String
		Lib_IsItemInList = False
		SearchString = UCase(Item)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object TheList.ListCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For i = lastindex To TheList.ListCount - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object TheList.List. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If SearchString = UCase(TheList.List(i)) Then
				Lib_IsItemInList = True
				lastindex = i
				Exit Function
			End If
		Next i
		lastindex = i
		Exit Function
		
ErrorL: 
		Lib_VBError("mLib", "Lib_IsItemInList", Erl(), "")
	End Function
	
	
	Public Function Lib_LoadResPicture(ByRef ID As Object, ByRef ImageType As Short) As System.Drawing.Image
		On Error GoTo ErrorL
		Select Case ImageType
			Case VB6.LoadResConstants.ResCursor
				'UPGRADE_ISSUE: Global method LoadResPicture was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
				Lib_LoadResPicture = VB6.LoadResPicture(ID, VB6.LoadResConstants.ResCursor)
			Case VB6.LoadResConstants.ResBitmap
				'UPGRADE_ISSUE: Global method LoadResPicture was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
				Lib_LoadResPicture = VB6.LoadResPicture(ID, VB6.LoadResConstants.ResBitmap)
			Case VB6.LoadResConstants.ResIcon
				'UPGRADE_ISSUE: Global method LoadResPicture was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
				Lib_LoadResPicture = VB6.LoadResPicture(ID, VB6.LoadResConstants.ResIcon)
		End Select
		
		Exit Function
ErrorL: 
		'UPGRADE_WARNING: Couldn't resolve default property of object ID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Lib_VBError("mLib", "Lib_LoadResPicture", Erl(), VB6.Format(ID), VB6.Format(ImageType))
	End Function
	
	
	
	Public Function Lib_GetFromImageList(ByRef ID As Object, ByRef ImageType As Short) As System.Drawing.Image
		On Error GoTo ErrorL
		Select Case ImageType
			Case VB6.LoadResConstants.ResCursor
				'UPGRADE_WARNING: Couldn't resolve default property of object ID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Lib_GetFromImageList = g_ilsImageList.Images.Item(KEY_CUR & VB6.Format(ID))
			Case VB6.LoadResConstants.ResBitmap
				'UPGRADE_WARNING: Couldn't resolve default property of object ID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Lib_GetFromImageList = g_ilsImageList.Images.Item(KEY_BMP & VB6.Format(ID))
			Case VB6.LoadResConstants.ResIcon
				'UPGRADE_WARNING: Couldn't resolve default property of object ID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Lib_GetFromImageList = g_ilsImageList.Images.Item(KEY_ICON & VB6.Format(ID))
		End Select
		Exit Function
		
ErrorL: 
		'UPGRADE_WARNING: Couldn't resolve default property of object ID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Lib_VBError("mLib", "Lib_GetFromImageList", Erl(), VB6.Format(ID), VB6.Format(ImageType))
	End Function
	
	Public Function Lib_AddtoImageList(ByRef ID As Object, ByRef ImageType As Short) As Integer
		On Error GoTo ErrorL
		Dim imgX As System.Drawing.Image ' Add images to ListImages collection.
		
		Select Case ImageType
			Case VB6.LoadResConstants.ResCursor
				'UPGRADE_WARNING: Couldn't resolve default property of object ID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				imgX = g_ilsImageList.Images.Add(KEY_CUR & VB6.Format(ID), Lib_LoadResPicture(ID, VB6.LoadResConstants.ResCursor))
			Case VB6.LoadResConstants.ResBitmap
				'UPGRADE_WARNING: Couldn't resolve default property of object ID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				imgX = g_ilsImageList.Images.Add(KEY_BMP & VB6.Format(ID), Lib_LoadResPicture(ID, VB6.LoadResConstants.ResBitmap))
			Case VB6.LoadResConstants.ResIcon
				'UPGRADE_WARNING: Couldn't resolve default property of object ID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				imgX = g_ilsImageList.Images.Add(KEY_ICON & VB6.Format(ID), Lib_LoadResPicture(ID, VB6.LoadResConstants.ResIcon))
		End Select
		Exit Function
		
ErrorL: 
		'UPGRADE_WARNING: Couldn't resolve default property of object ID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Lib_VBError("mLib", "Lib_AddtoImageList", Erl(), VB6.Format(ID), VB6.Format(ImageType))
	End Function
	
	
	Public Sub Lib_SetImageList(ByRef ImageList As System.Windows.Forms.ImageList)
		On Error GoTo ErrorL
		g_ilsImageList = ImageList
		Exit Sub
		
ErrorL: 
		Lib_VBError("mLib", "Lib_SetImageList", Erl(), "")
	End Sub
	
	Public Function Lib_Right(ByRef s As String, ByRef l As Short) As Object
		On Error GoTo ErrorL
		If Len(s) > l Then
			Lib_Right = Right(s, l)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Lib_Right. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Lib_Right = s
		End If
		Exit Function
		
ErrorL: 
		Lib_VBError("mLib", "Lib_Right", Erl(), "")
	End Function
	
	
	
	Public Function Lib_GetMainIcon() As System.Drawing.Image
		On Error GoTo ErrorL
		Lib_GetMainIcon = gfrmMain.Icon
		Exit Function
ErrorL: 
		Lib_VBError("mLib", "Lib_GetMainIcon", Erl(), "")
	End Function
	
	
	Public Function Lib_GetMainForm() As System.Windows.Forms.Form
		On Error GoTo ErrorL
		Lib_GetMainForm = gfrmMain
		Exit Function
ErrorL: 
		Lib_VBError("mLib", "Lib_GetMainForm", Erl(), "")
	End Function
	
	
	Public Function Lib_StrBlankIfZero(ByRef Value As Object) As String
		On Error GoTo ErrorL
		'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Value = 0 Then
			Lib_StrBlankIfZero = ""
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Lib_StrBlankIfZero = CStr(Value)
		End If
		Exit Function
ErrorL: 
		Lib_VBError("mLib", "Lib_StrBlankIfZero", Erl(), "")
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
		
		'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If IsNothing(StopDate) Or Len(StopDate) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object StopDate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			StopDate = Now
		End If
		'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If IsNothing(StartDate) Or Len(StartDate) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object StartDate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			StartDate = Now
		End If
		If IsDate(StopDate) And IsDate(StartDate) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object StartDate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object StopDate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Daydiff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Daydiff = System.Date.FromOADate(CDate(StopDate).ToOADate - CDate(StartDate).ToOADate)
			'UPGRADE_WARNING: Couldn't resolve default property of object Daydiff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CInt(Daydiff) < 10000 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Daydiff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object TotalSecDiff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TotalSecDiff = CInt(Daydiff * 86400)
				'UPGRADE_WARNING: Couldn't resolve default property of object TotalSecDiff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Lib_GetDiffSecond = CInt(TotalSecDiff)
			Else
				Lib_GetDiffSecond = 999999
			End If
		Else
			Lib_GetDiffSecond = 999999
		End If
		Exit Function
		
ErrorL: 
		Lib_VBError("mLib", "Lib_GetDiffSecond", Erl(), "")
	End Function
	
	Public Sub Lib_MakeDir(ByRef Path As String)
		On Error Resume Next
		Dim i As Short
		Dim sPath, s, sDrive, sFile As String
		Lib_SplitFilename(Path, sDrive, sPath, sFile)
		s = sDrive & sPath
		If Len(s) > 3 Then
			i = InStr(1, s, "\")
			While i > 0
				On Error Resume Next
				MkDir(Left(s, i - 1))
				i = InStr(i + 1, s, "\")
			End While
		End If
		MkDir(s)
		Err.Clear()
		Exit Sub
ErrorL: 
		Lib_VBError("mLib", "Lib_MakeDir", Erl(), g_sNoError, g_sNoError)
	End Sub
	
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
		Lib_VBError("mLib", "Lib_SplitFilename", Erl(), g_sNoError, g_sNoError)
	End Sub
	
	
	Public Function Lib_InitMousePointer() As Boolean
		On Error GoTo ErrorL
		nMousePointerCounter = 0
		Lib_InitMousePointer = True
		Exit Function
		
ErrorL: 
		Lib_InitMousePointer = False
		Lib_VBError("mLib", "Lib_InitMousePointer", Erl(), "")
	End Function
	
	
	Public Function Lib_ResetMousePointer() As Short
		On Error GoTo ErrorL
		nMousePointerCounter = 0
		Lib_ResetMousePointer = System.Windows.Forms.Cursors.Default
		Exit Function
		
ErrorL: 
		Lib_ResetMousePointer = System.Windows.Forms.Cursors.Default
		Lib_VBError("mLib", "Lib_ResetMousePointer", Erl(), "")
	End Function
	
	
	Public Function Lib_GetArrow() As Short
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
		Lib_VBError("mLib", "Lib_GetArrow", Erl(), "")
	End Function
	
	
	Public Function Lib_GetHourGlass() As Short
		On Error GoTo ErrorL
		nMousePointerCounter = nMousePointerCounter + 1
		Lib_GetHourGlass = System.Windows.Forms.Cursors.WaitCursor
		Exit Function
		
ErrorL: 
		Lib_GetHourGlass = System.Windows.Forms.Cursors.WaitCursor
		Lib_VBError("mLib", "Lib_GetHourGlass", Erl(), "")
	End Function
	
	
	Public Sub Lib_SetWindowPosTop(ByRef hwnd As Integer)
		On Error GoTo ErrorL
		SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
		Exit Sub
ErrorL: 
		Lib_VBError("mLib", "Lib_SetWindowPosTop", Erl(), "")
	End Sub
	
	
	Public Sub Lib_SetWindowPosTopMost(ByRef hwnd As Integer)
		On Error GoTo ErrorL
		SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
		Exit Sub
ErrorL: 
		Lib_VBError("mLib", "Lib_SetWindowPosTopMost", Erl(), "")
	End Sub
	
	
	Public Sub Lib_SetWindowPosNotTopMost(ByRef hwnd As Integer)
		On Error GoTo ErrorL
		SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
		Exit Sub
ErrorL: 
		Lib_VBError("mLib", "Lib_SetWindowPosNotTopMost", Erl(), "")
	End Sub
	
	
	Public Sub Lib_SetWindowPosBottom(ByRef hwnd As Integer)
		On Error GoTo ErrorL
		SetWindowPos(hwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
		
		Exit Sub
ErrorL: 
		Lib_VBError("mLib", "Lib_SetWindowPosBottom", Erl(), "")
	End Sub
	
	
	Public Sub Lib_SetWindowPosAfter(ByRef hwnd As Integer, ByRef hWndInsertAfter As Integer)
		On Error GoTo ErrorL
		SetWindowPos(hwnd, hWndInsertAfter, 0, 0, 0, 0, &H1 + &H2)
		Exit Sub
ErrorL: 
		Lib_VBError("mLib", "Lib_SetWindowPosAfter", Erl(), "")
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
		Lib_VBError("mLib", "SetTopMostWindow", Erl(), "")
	End Function
	
	
	
	Public Function Lib_WriteINIstr(ByVal Section As String, ByVal Entry As String, ByVal sValue As String, ByVal lplFileName As String) As Boolean
		On Error GoTo ErrorL
		Dim Ret As Short
		
		If Entry = "" Then
			Ret = WritePrivateProfileStringAPI(Section, 0, 0, lplFileName)
		Else
			If sValue = "" Then
				Ret = WritePrivateProfileStringAPI(Section, Entry, 0, lplFileName)
			Else
				Ret = WritePrivateProfileStringAPI(Section, Entry, sValue, lplFileName)
			End If
		End If
		Lib_WriteINIstr = (Ret <> 0)
		Exit Function
ErrorL: 
		Lib_VBError("mLib", "Lib_WriteINIstr", Erl(), "")
	End Function
	
	
	'UPGRADE_NOTE: Default was upgraded to Default_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function Lib_GetINIstr(ByVal Section As String, ByVal Entry As String, ByVal Default_Renamed As String, ByVal lplFileName As String) As String
		On Error GoTo ErrorL
		Dim Ret As Short
		Dim s As String
		Dim Pos As Short
		s = New String(Chr(0), 4096)
		
		If Entry = "" Then
			Ret = GetPrivateProfileStringAPI(Section, 0, Default_Renamed, s, Len(s), lplFileName)
		Else
			Ret = GetPrivateProfileStringAPI(Section, Entry, Default_Renamed, s, Len(s), lplFileName)
		End If
		
		If Ret > 0 Then
			If Entry <> "" Then
				Pos = InStr(s, Chr(0))
				If Pos > 1 Then
					Lib_GetINIstr = Trim(Left(s, Pos - 1))
				Else
					Lib_GetINIstr = ""
				End If
			Else
				Lib_GetINIstr = Trim(s)
			End If
		End If
		Exit Function
ErrorL: 
		Lib_VBError("mLib", "Lib_GetINIstr", Erl(), "")
	End Function
	
	
	Public Function Lib_NameOfUser(ByRef UserName As String) As Integer
		On Error GoTo ErrorL
		Dim NameSize As Integer
		Dim X As Integer
		Dim verinfo As OSVERSIONINFO
		
		UserName = Space(16)
		NameSize = Len(UserName)
		X = GetUserName(UserName, NameSize)
		If X Then
			UserName = Lib_StringC2VB(UserName)
		Else
			UserName = vbNullString
		End If
		
		Exit Function
ErrorL: 
		Lib_VBError("mLib", "Lib_NameOfUser", Erl(), UserName)
	End Function
	
	'FindWindowLike
	' - Finds the window handles of the windows matching the specified
	'   parameters
	'
	'hwndArray()
	' - An integer array used to return the window handles
	'
	'hWndStart
	' - The handle of the window to search under.
	' - The routine searches through all of this window's children and their
	'   children recursively.
	' - If hWndStart = 0 then the routine searches through all windows.
	'
	'WindowText
	' - The pattern used with the Like operator to compare window's text.
	'
	'ClassName
	' - The pattern used with the Like operator to compare window's class
	'   name.
	'
	'ID
	' - A child ID number used to identify a window.
	' - Can be a decimal number or a hex string.
	' - Prefix hex strings with "&H" or an error will occur.
	' - To ignore the ID pass the Visual Basic Null function.
	'
	'Returns
	' - The number of windows that matched the parameters.
	' - Also returns the window handles in hWndArray()
	'
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
				'UPGRADE_WARNING: Couldn't resolve default property of object sID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sID = CInt("&H" & Hex(R))
			Else
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				'UPGRADE_WARNING: Couldn't resolve default property of object sID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sID = System.DBNull.Value
			End If
			' Check that window matches the search parameters:
			If sWindowText Like WindowText And sClassname Like Classname Then
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IsDbNull(ID) Then
					' If find a match, increment counter and
					'  add handle to array:
					iFound = iFound + 1
					ReDim Preserve hWndArray(iFound)
					hWndArray(iFound) = hwnd
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				ElseIf Not IsDbNull(sID) Then 
					'UPGRADE_WARNING: Couldn't resolve default property of object ID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object sID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		Lib_VBError("mLib", "Lib_FindWindowLike", Erl(), "")
	End Function
	
	
	Public Sub Lib_FormAlignAtCursor(ByRef Theform As System.Windows.Forms.Form)
		'Align form at position of cursor
		On Error GoTo ErrorL
		Dim X, y As Integer
		Dim p As POINTAPI
		Dim Ret As Integer
		
		Ret = GetCursorPosAPI(p)
		X = p.X
		y = p.y
		
		X = X - (VB6.PixelsToTwipsX(Theform.Width) / 2)
		
		If X + VB6.PixelsToTwipsX(Theform.Width) > VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) Then
			X = VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Theform.Width)
		End If
		
		If y + VB6.PixelsToTwipsY(Theform.Height) > VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) Then
			y = VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Theform.Height)
		End If
		
		Theform.SetBounds(VB6.TwipsToPixelsX(X), VB6.TwipsToPixelsY(y), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		Exit Sub
ErrorL: 
		Lib_VBError("mLib", "Lib_FormAlignAtCursor", Erl(), g_sNoError, g_sNoError)
	End Sub
	
	
	Public Sub Lib_FormAlignAtCenter(ByRef CenterForm As System.Windows.Forms.Form, ByRef Theform As System.Windows.Forms.Form)
		On Error GoTo ErrorL
		'Align form at center of another form
		Dim w, l, t, H As Integer
		Dim Wn, ln, tn, Hn As Integer
		
		l = VB6.PixelsToTwipsX(CenterForm.Left)
		t = VB6.PixelsToTwipsY(CenterForm.Top)
		H = VB6.PixelsToTwipsY(CenterForm.Height)
		w = VB6.PixelsToTwipsX(CenterForm.Width)
		
		Wn = l + w / 2 'this is the x mid of our window
		Hn = t + H / 2 'this is the y mid of our window
		
		ln = Wn - (VB6.PixelsToTwipsX(Theform.Width) / 2) 'adjust according to width of form
		tn = Hn - (VB6.PixelsToTwipsY(Theform.Height) / 2) 'adjust according to height of form
		If ln > l Then
			l = ln
		End If
		If tn > t Then
			t = tn
		End If
		Theform.SetBounds(VB6.TwipsToPixelsX(l), VB6.TwipsToPixelsY(t), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		Exit Sub
		
ErrorL: 
		Lib_VBError("mLib", "Lib_FormAlignAtCenter", Erl(), "")
	End Sub
	
	'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub Lib_FormAlignAtLeft(ByRef Theform As System.Windows.Forms.Form, ByRef Left_Renamed As Short)
		On Error GoTo ErrorL
		Theform.SetBounds(VB6.TwipsToPixelsX(Left_Renamed), 0, 0, 0, Windows.Forms.BoundsSpecified.X)
		Exit Sub
		
ErrorL: 
		Lib_VBError("mLib", "Lib_FormAlignAtLeft", Erl(), "")
	End Sub
	
	Public Sub Lib_FormAlignAtTopLeft(ByRef CenterForm As System.Windows.Forms.Form, ByRef Theform As System.Windows.Forms.Form)
		On Error GoTo ErrorL
		'Align form at center of another form
		On Error GoTo ErrorL
		Dim w, l, t, H As Integer
		
		l = VB6.PixelsToTwipsX(CenterForm.Left)
		t = VB6.PixelsToTwipsY(CenterForm.Top)
		
		Theform.SetBounds(VB6.TwipsToPixelsX(l), VB6.TwipsToPixelsY(t), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		Exit Sub
ErrorL: 
		Lib_VBError("mLib", "Lib_FormAlignAtTopLeft", Erl(), g_sNoError, g_sNoError)
	End Sub
	
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
		Lib_VBError("mLib", "Lib_StringC2VB", Erl(), "")
	End Function
	
	
	Public Function Lib_NameOfPC(ByRef MachineName As String) As Integer
		On Error GoTo ErrorL
		
		Dim NameSize As Integer
		Dim X As Integer
		
		MachineName = Space(16)
		NameSize = Len(MachineName)
		X = GetComputerName(MachineName, NameSize)
		If X Then
			MachineName = Lib_StringC2VB(MachineName)
		Else
			MachineName = vbNullString
		End If
		Exit Function
ErrorL: 
		Lib_VBError("mLib", "Lib_NameOfPC", Erl(), g_sNoError, g_sNoError)
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
		Lib_VBError("mLib", "Lib_Crypt", Erl(), "")
	End Function
End Module