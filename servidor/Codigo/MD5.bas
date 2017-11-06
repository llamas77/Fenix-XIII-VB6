Attribute VB_Name = "MD5"
Option Explicit

Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal T As Long, ByVal r As String)
 
Public Function MD5String(p As String) As String
'compute MD5 digest on a given string, returning the result
    Dim r As String * 32, T As Long
    r = Space(32)
    T = Len(p)
    MDStringFix p, T, r
    MD5String = r
End Function
