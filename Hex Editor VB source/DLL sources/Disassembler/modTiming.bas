Attribute VB_Name = "modTiming"
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32.dll" (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32.dll" (ByRef lpFrequency As Currency) As Long

Dim perffreq As Currency
Dim starttime As Currency
Dim stoptime As Currency

Public Sub ResetTimer()
QueryPerformanceFrequency perffreq
starttime = 0
stoptime = 0
End Sub

Public Sub StartTimer()
QueryPerformanceCounter starttime
End Sub

Public Function StopTimer() As Single
QueryPerformanceCounter stoptime
StopTimer = (stoptime - starttime) / perffreq
End Function

