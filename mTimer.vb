Option Strict Off
Option Explicit On
Module mTimer
	
	
	Declare Function QueryPerformanceCounter Lib "Kernel32" (ByRef X As Decimal) As Boolean
	Declare Function QueryPerformanceFrequency Lib "Kernel32" (ByRef X As Decimal) As Boolean
	
	Private TimerResolutionNanoSec As Integer
	Private Overhead As Decimal
	Private Ctr2, Ctr1, Freq As Decimal
	
	Sub Timer_Init()
		' Time QueryPerformanceCounter      '
		If QueryPerformanceCounter(Ctr1) Then
			QueryPerformanceCounter(Ctr2)
			QueryPerformanceFrequency(Freq)
			TimerResolutionNanoSec = CInt(1 / (Freq / 100000))
			Overhead = (Ctr2 - Ctr1)
		Else
			Debug.Print("High-resolution counter not supported.")
		End If
		
		Timer_Start()
		Call Timer_Calibrate_ms()
	End Sub
	
	Public Function Timer_GetTimerResolutionNanoSec() As Integer
		Timer_GetTimerResolutionNanoSec = TimerResolutionNanoSec
	End Function
	Public Sub Timer_Start()
		If Freq = 0 Then
			Timer_Init()
		End If
		QueryPerformanceCounter(Ctr1)
	End Sub
	
	Public Function Timer_Calibrate_ms() As Double
		QueryPerformanceCounter(Ctr2)
		Overhead = (Ctr2 - Ctr1)
		If Freq = 0 Then
			Timer_Init()
		End If
		Timer_Calibrate_ms = (Overhead / Freq) * 1000
        Debug.Print(("Timer_Calibrate_ms = " & Format(Timer_Calibrate_ms)) & " ms")
		
	End Function
	
	Public Function Timer_Stop_ms(ByRef text As String) As Double
        '2.0.0 Dim t As Decimal
		
		QueryPerformanceCounter(Ctr2)
		If Freq = 0 Then
			Timer_Init()
		End If
		Timer_Stop_ms = ((Ctr2 - Ctr1 - Overhead) / Freq) * 1000
        Debug.Print((text & " = " & Format(Timer_Stop_ms)) & " ms")
	End Function
End Module