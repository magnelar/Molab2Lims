Option Strict Off
Option Explicit On
'Imports System.Threading 
'Molab2Lims
'http://www.codeproject.com/Articles/31960/Upgrading-VB-to-VB-NET
Module Do_Events
    '

    Friend Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Integer, _
         ByVal nPriority As Integer) As Integer

    Friend Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Integer, _
         ByVal dwPriorityClass As Integer) As Integer

    Friend Declare Function GetCurrentThread Lib "kernel32" () As Integer
    Friend Declare Function GetCurrentProcess Lib "kernel32" () As Integer

    Friend Const THREAD_PRIORITY_HIGHEST As Short = 2
    Friend Const HIGH_PRIORITY_CLASS As Integer = &H80

    Public Sub test()
        'Dim th As New Thread(AddressOf WriteY)
        SetThreadPriority(GetCurrentThread, THREAD_PRIORITY_HIGHEST)
        SetPriorityClass(GetCurrentProcess, HIGH_PRIORITY_CLASS)
        Dim iLoops As Integer = 0
        'th.Priority = ThreadPriority.Lowest
        'th.Start()
        Do Until iLoops = 10000

            'Calling DoEvents() every 500 loops will greatly increase application performance
            If iLoops Mod 500 = 0 Then Application.DoEvents()

            iLoops += 1 'Add 1 to iLoops

        Loop
    End Sub
End Module
