Attribute VB_Name = "modGlobals"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Const TASK_FILE_NAME As String = "Tasks.dat"

Global gstrTaskFile As String
Global gobjVBInstance  As VBIDE.VBE
Global gwinWindow   As VBIDE.Window
Global gblMouseClick As Integer

Public Function DoError(psModule As String, psProc As String, piErr As Integer)

   On Error Resume Next
   
   If piErr <> 0 Then
      MsgBox "Error: " & Str(piErr) & " - " & Error(piErr) & " in Module: " & psModule & " in Procedure: " & psProc, vbCritical
   End If
   
End Function

Function GetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String

   On Error GoTo GetFromIni_Error
   
   Dim strReturn As String
   
   strReturn = String(255, Chr(0))
   GetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
   
   Exit Function
   
GetFromIni_Error:
   DoError "modGlobals", "GetFromIni", Err
    
End Function

Function WriteToIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer

   On Error GoTo WriteToIni_Error
   WriteToIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)

   Exit Function
   
WriteToIni_Error:
   DoError "modGlobals", "WriteToIni", Err
   
End Function

Public Function FileExists(strFile As String) As Boolean
   
   On Error Resume Next 'Doesn't raise error - FileExists will be False
   
   FileExists = Dir(strFile, vbHidden) <> ""

End Function

Public Function GetToken(sSrc As String, sDelimit As String)
    
   Dim ilast As Integer, iLoop As Integer, iPos As Integer
   Dim sToken As String
   
   If sDelimit = "" Then sDelimit = ","
   
   ilast = 32767
   
   For iLoop = 1 To Len(sDelimit)
      iPos = InStr(sSrc, Mid$(sDelimit, iLoop, 1))
      If iPos <> 0 And iPos < ilast Then ilast = iPos
   Next
   
   If ilast <> 32767 Then
      If ilast = 1 Then
         sToken = ""
      Else
         sToken = Mid$(sSrc, 1, ilast - 1)
      End If
      sSrc = Mid$(sSrc, ilast + 1)
   Else
      sToken = sSrc
      sSrc = ""
   End If
   
   GetToken = sToken

End Function
