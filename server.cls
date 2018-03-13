VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'The class name for Z80 Simulator IDE: z80moduletemplate.server

Public Sub objectinit()
Call Form1.objectinit
End Sub

Public Sub objectrefresh()
Call Form1.objectrefresh
End Sub

Public Sub writeio(port, data)
Call Form1.writeio(port, data)
End Sub

Public Sub readio(port, ByRef data)
Call Form1.readio(port, data)
End Sub

Public Sub objectterm()
End
End Sub
