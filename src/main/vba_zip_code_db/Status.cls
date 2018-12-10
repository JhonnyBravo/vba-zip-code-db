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

'@see IStatus#setCode
Public Function IStatus_setCode(lngCode As Long) As IStatus
    propCode = lngCode
    Set IStatus_setCode = Me
End Function

'@see IStatus#getCode
Public Function IStatus_getCode() As Long
    IStatus_getCode = propCode
End Function

'@see IStatus#errorTerminate
Public Function IStatus_errorTerminate(strMessage As String) As IStatus
    propCode = 1
    Set IStatus_errorTerminate = Me
    
    Debug.Print strMessage
    MsgBox strMessage
End Function

'@see IStatus#initStatus
Public Function IStatus_initStatus() As IStatus
    propCode = 0
    Set IStatus_initStatus = Me
End Function
