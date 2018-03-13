Attribute VB_Name = "Module1"
Option Explicit

Private Sub Main()
Form1.Show
End Sub

Public Function hex2(ByVal ulaz As Long) As String
hex2 = Right("0" + Hex(ulaz), 2) + "h"
End Function

Public Function hex4(ByVal ulaz As Long) As String
hex4 = Right("000" + Hex(ulaz), 4) + "h"
End Function
