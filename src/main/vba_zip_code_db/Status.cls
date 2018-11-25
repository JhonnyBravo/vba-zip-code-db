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

Private propCode As Long
Private propMessage As String

'@return code メソッド終了直後の終了コードを返す。
Public Property Get code() As Long
    code = propCode
End Property

'@param lngCode メソッド終了直後の終了コードを指定する。
'   0: 成功も失敗もなく、何もせずに終了した状態を表す。
'   1: 異常終了した状態を表す。
'   2: メソッドの実行に成功した状態を表す。
Public Property Let code(lngCode As Long)
    propCode = lngCode
End Property

'@return message メソッド終了直後のメッセージを返す。
Public Property Get message() As String
    message = propMessage
End Property

'@param strMessage メソッド終了直後のメッセージを指定する。
'   原則的にエラー終了した場合に限り設定し、
'   それ以外の場合は空文字を設定するものとする。
Public Property Let message(strMessage As String)
    propMessage = strMessage
End Property
