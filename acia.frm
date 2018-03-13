VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ACIA Module"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   3975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   4320
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim object1 As Object
Dim doRecInt As Boolean
Dim doTransInt As Boolean
Dim recBuf As String
Dim lastSend As Long
Dim thisSend As Long
Dim lastTrans As Long
Dim port80 As Integer

Public Sub objectinit()
recBuf = ""
End Sub


Public Sub objectrefresh()
DoEvents
If object1.getclockcycles() > lastTrans + 300 And doTransInt Then
    Call object1.z80int
    lastTrans = object1.getclockcycles()
End If
If Len(recBuf) > 0 And object1.getclockcycles() > lastSend + 200 Then
    Call object1.z80int
    lastSend = object1.getclockcycles()
End If
End Sub

Public Sub writeio(port, data)
If port = &H80 Then
    port80 = data
    Label1.Caption = "80: " & CStr(port80)
    If (data And &H60) = &H20 Then
        doTransInt = True
    Else
        doTransInt = False
    End If
    If (data And &H80) = &H80 Then
        doRecInt = True
    Else
        doRecInt = False
    End If
ElseIf port = &H81 Then
    Text1.Text = Text1.Text + Chr(data)
    lastTrans = object1.getclockcycles()
End If
End Sub

Public Sub readio(port, ByRef data)
If port = &H80 Then
    If Len(recBuf) = 0 Then
        data = &HFE
    Else
        data = &HFF
    End If
ElseIf port = &H81 Then
    data = Asc(Left(recBuf, 1))
    recBuf = Mid(recBuf, 2)
End If
End Sub

Private Sub Form_Load()
On Error GoTo error1
Set object1 = CreateObject("z80simulatoride.server")

Exit Sub
error1:
MsgBox "Could not establish connection with Z80 Simulator IDE.", , "Z80 Module Template"
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
recBuf = recBuf & Chr(KeyAscii)
KeyAscii = 0
Label1.Caption = recBuf
End Sub
