Option Strict Off
Option Explicit On
'Molab2Lims
Module ModMolabTextOperations
	
	Declare Function GetProfileString32 Lib "Kernel32"  Alias "GetProfileStringA"(ByVal strSection As String, ByVal strEntry As String, ByVal strDefault As String, ByVal strReturnBuffer As String, ByVal intReturnBuffer As Short) As Short
	
	Public gDesimal As String
	Public Function DecConvString(ByRef pInStr As String) As String
		Dim lWrongDesimal As String
		If Not gNoErrorHandling Then On Error GoTo Err_DecConvString
		Call GetDesimalSeparator()
		If gDesimal = "," Then
			lWrongDesimal = "."
		Else
			lWrongDesimal = ","
		End If
		DecConvString = Replace(pInStr, lWrongDesimal, gDesimal)
		
Exit_DecConvString: 
		Exit Function
		
Err_DecConvString: 
		'UpdateInfo
		'StoreInfo
        Call GenericError("DecConvString", "Feil ved lesing av desimalskilletegn fra Registry", Err.Number, Err.Description, Err.Source)
        Resume Exit_DecConvString
    End Function
	
	
	Public Sub GetDesimalSeparator()
		If Not gNoErrorHandling Then On Error GoTo Err_GetDesimalSeparator
		'Read from registry: "HKEY_CURRENT_USER\Control Panel\International\sDesimal"
		Dim strBuffer As String
		Dim intDResult As Short
		strBuffer = Space(10)
		intDResult = GetProfileString32("Intl", "sDecimal", ",", strBuffer, 10)
		gDesimal = Left(strBuffer, intDResult)
		gDesimal = "" & gDesimal & ""
		If gDesimal = "" Then
			gDesimal = ","
		End If
		
Exit_GetDesimalSeparator: 
		Exit Sub
		
Err_GetDesimalSeparator: 
		'UpdateInfo
		'StoreInfo
        Call GenericError("GetDesimalSeparator", "Feil ved lesing av desimalskilletegn fra Registry", Err.Number, Err.Description, Err.Source)
        Resume Exit_GetDesimalSeparator
    End Sub
	Public Function RemoveSpaces(ByRef pText As String) As String
		If Not gNoErrorHandling Then On Error GoTo Err_RemoveDoubleSpaces
		Dim i As Short
        Dim lText As String = String.Empty
		For i = 1 To Len(pText)
			If Mid(pText, i, 1) <> " " Then
				lText = lText & Mid(pText, i, 1)
			End If
		Next i
		RemoveSpaces = lText
		
Exit_RemoveDoubleSpaces: 
		Exit Function
		
Err_RemoveDoubleSpaces: 
		'UpdateInfo
		'StoreInfo
        Call GenericError("RemoveDoubleSpaces", "Feil ved fjerning av etterfølgende mellomrom", Err.Number, Err.Description, Err.Source)
        Resume Exit_RemoveDoubleSpaces
    End Function
	Public Sub ExtractParam(ByRef pText As String, ByRef pParam As String, ByRef pNumber As Single, Optional ByRef pIgnoreParenthesis As Boolean = False)
		'Extract the text (into pParam) from the begining of the line
		'(eliminate double spaces) from pText
		'Extract number from the end of the line
		Dim lNumberFound As Boolean
        Dim lStrNumber As String = String.Empty
		Dim lStrParam As String
		Dim lText As String
		Dim i As Byte
		Dim lInsideParenthesis As Boolean
		
		lInsideParenthesis = False
		lNumberFound = False
		lStrParam = " "
		i = Len(RTrim(pText))
		If i > 0 Then
			Do 
				If InStr(1, "0123456789.,", Mid(pText, i, 1)) > 0 Then
					If InStr(1, ".,", Mid(pText, i, 1)) > 0 Then
						lStrNumber = gDesimal & lStrNumber
					Else
						lStrNumber = Mid(pText, i, 1) & lStrNumber
					End If
					i = i - 1
				Else
					lNumberFound = True
				End If
			Loop While Not lNumberFound And i > 1
			lText = pText
			Do 
				If pIgnoreParenthesis Then
					If Mid(pText, i, 1) = ")" Then lInsideParenthesis = True
					If lInsideParenthesis And Mid(pText, i, 1) = "(" Then
						lInsideParenthesis = False
						lText = Mid(pText, 1, i - 1)
					End If
				End If
				If Not lInsideParenthesis Or Not pIgnoreParenthesis Then
					If (Mid(lText, i, 1) = " " And Left(lStrParam, 1) = " ") Or (Mid(lText, i, 1) = "=") Then
						'Avoid two trailing spaces, and ignore =
					Else
						lStrParam = Mid(lText, i, 1) & lStrParam
					End If
				End If
				i = i - 1
			Loop While i > 0
			
			pParam = Trim(lStrParam)
			If IsNumeric(lStrNumber) Then
				pNumber = CSng(lStrNumber)
			End If
		End If
		
	End Sub
End Module