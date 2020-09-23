Attribute VB_Name = "Module_Arithmetic"
'***DLL_Arithmetic, for testing VB-Win32DLL Compiler,
'***made by KFGG, China.P.R
'***12/12/2004

Option Explicit

Function Plus(ByVal a As Long, ByVal b As Long) As Long
  Plus = a + b
End Function

Function Minus(ByVal a As Long, ByVal b As Long) As Long
  Minus = a - b
End Function

Function Multiply(ByVal a As Long, ByVal b As Long) As Long
  Multiply = a * b
End Function

Function Divide(ByVal a As Long, ByVal b As Long) As Long
  If b <> 0 Then Divide = a \ b
End Function

Sub About()
  MsgBox "Arithmetic DLL, for VB-Win32DLL Compiler test", vbInformation, "DLL_Arithmetic"
End Sub

Sub Main()
  MsgBox "Sub Main, this sub can be left empty.", vbInformation, "DLL_Arithmetic"
End Sub

