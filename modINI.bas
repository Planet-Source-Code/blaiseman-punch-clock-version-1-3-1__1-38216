Attribute VB_Name = "modINI"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public dtSelected As Date
Public intTextFieldLoc As Integer
Public strOptCWW As String
Public dblMainTop As Double
Public dblMainLeft As Double
Public strLunch As String

Public Function ReadINI(strsection As String, strkey As String, strfullpath As String) As String
   Dim strbuffer As String
   Let strbuffer$ = String$(750, Chr$(0&))
   Let ReadINI$ = Left$(strbuffer$, GetPrivateProfileString(strsection$, ByVal LCase$(strkey$), "", strbuffer, Len(strbuffer), strfullpath$))
End Function

Public Sub WriteINI(strsection As String, strkey As String, strkeyvalue As String, strfullpath As String)
    Call WritePrivateProfileString(strsection$, UCase$(strkey$), strkeyvalue$, strfullpath$)
End Sub

Public Function CenterMe(frm As Form) As String
    Dim x As String
    Dim y As String
    frm.Top = (Screen.Height / 2) - (frm.Height / 2)
    frm.Left = (Screen.Width / 2) - (frm.Width / 2)
    
    x = frm.Top
    y = frm.Left
    
    CenterMe = x & "*" & y
End Function

