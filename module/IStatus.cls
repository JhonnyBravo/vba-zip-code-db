VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Public Property Get code() As Long
End Property

Public Property Let code(lngCode As Long)
End Property

Public Property Get message() As String
End Property

Public Property Let message(strMessage As String)
End Property
