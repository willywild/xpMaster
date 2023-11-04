Attribute VB_Name = "IdleTime"
Option Explicit

Private Type LASTINPUTINFO
  cbSize As Long
  dwTime As Long
End Type

Private Declare Sub GetLastInputInfo Lib "user32" (ByRef plii As LASTINPUTINFO)
Private Declare Function GetTickCount Lib "kernel32" () As Long

Function IdleTime() As Single
  Dim a As LASTINPUTINFO
  a.cbSize = LenB(a)
  GetLastInputInfo a
  Debug.Print a.dwTime
  IdleTime = (GetTickCount - a.dwTime) / 1000
End Function

Sub PrintIdleTime1()
  Debug.Print IdleTime
    Application.OnTime Now + TimeSerial(0, 0, 5), "PrintIdleTime2"
End Sub

Sub PrintIdleTime2()
Debug.Print IdleTime
  Application.OnTime Now + TimeSerial(0, 0, 5), "PrintIdleTime1"
End Sub

Sub printTickCount()
    Debug.Print GetTickCount
End Sub
