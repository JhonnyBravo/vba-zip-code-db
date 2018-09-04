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

Private propCode As Long
Private propMessage As String

'@return code メソッド終了直後の終了コードを返す。
Public Property Get code() As Long
    code = propCode
End Property

'@param lngCode メソッド終了直後の終了コードを設定する。
'   0: 異常もリソースの変更もなく終了した状態を表す。
'   1: 異常終了した状態を表す。
'   2: リソースの変更に成功した状態を表す。
Public Property Let code(lngCode As Long)
    propCode = lngCode
End Property

'@return message エラー発生時に出力するメッセージを返す。
Public Property Get message() As String
    message = propMessage
End Property

'@param strMessage エラー発生時に出力するメッセージを設定する。
Public Property Let message(strMessage As String)
    propMessage = strMessage
End Property

'終了コードとエラーメッセージを初期化する。
Public Sub initStatus()
    code = 0
    message = ""
End Sub

'終了コードを 1 に設定し、エラーメッセージを出力する。
Public Sub errorTerminate()
    code = 1
    Debug.Print message
    MsgBox message
End Sub
