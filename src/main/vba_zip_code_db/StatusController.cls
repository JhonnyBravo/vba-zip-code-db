VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StatusController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Database

Implements IStatus
Private objStatus As New Status

Public Property Let IStatus_code(lngCode As Long)
    objStatus.code = lngCode
End Property

Public Property Get IStatus_code() As Long
    IStatus_code = objStatus.code
End Property

Public Property Let IStatus_message(strMessage As String)
    objStatus.message = strMessage
End Property

Public Property Get IStatus_message() As String
    IStatus_message = objStatus.message
End Property

'終了コードとエラーメッセージを初期化する。
Public Sub initStatus()
    IStatus_code = 0
    IStatus_message = ""
End Sub

'終了コードを 1 に設定し、エラーメッセージを出力する。
Public Sub errorTerminate()
    IStatus_code = 1
    Debug.Print IStatus_message
    MsgBox IStatus_message
End Sub
