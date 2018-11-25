VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IStatus

Private propCode As Long
Private propMessage As String

'Properties

'@see IStatus#code
Public Property Let IStatus_code(lngCode As Long)
    propCode = lngCode
End Property

'@see IStatus#code
Public Property Get IStatus_code() As Long
    IStatus_code = propCode
End Property

'@see IStatus#message
Public Property Let IStatus_message(strMessage As String)
    propMessage = strMessage
End Property

'@see IStatus#message
Public Property Get IStatus_message() As String
    IStatus_message = propMessage
End Property

'Methods

'@see IStatus#errorTerminate
Public Sub IStatus_errorTerminate()
    IStatus_code = 1
    Debug.Print IStatus_message
    MsgBox IStatus_message
End Sub

'@see IStatus#initStatus
Public Sub IStatus_initStatus()
    IStatus_code = 0
    IStatus_message = ""
End Sub
