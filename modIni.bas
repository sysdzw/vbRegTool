Attribute VB_Name = "modIni"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public iniFileName As String 
Function GetIniS(ByVal SectionName As String, ByVal KeyWord As String, Optional ByVal DefString As String) As String
    Dim ResultString As String * 144, Temp%
    Dim s$, i%
    Temp% = GetPrivateProfileString(SectionName, KeyWord, "", ResultString, 144, iniFileName)
   
    If Temp% > 0 Then
        For i = 1 To 144
            If Asc(Mid$(ResultString, i, 1)) <> 0 Then
                s = s & Mid$(ResultString, i, 1)
            End If
        Next
    Else
        Temp% = WritePrivateProfileString(SectionName, KeyWord, DefString, iniFileName)
        s = DefString
    End If
    GetIniS = s
End Function

Public Function SetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String) As Boolean
    SetIniS = WritePrivateProfileString(SectionName, KeyWord, ValStr, iniFileName)
End Function

Public Function DelIniSec(ByVal SectionName As String) As Boolean
    DelIniSec = WritePrivateProfileString(SectionName, 0&, "", iniFileName)
End Function
'delKeyWord
Public Function DelIniKey(ByVal SectionName As String, ByVal KeyWord As String) As Boolean
    DelIniKey = WritePrivateProfileString(SectionName, KeyWord, 0&, iniFileName)
End Function

