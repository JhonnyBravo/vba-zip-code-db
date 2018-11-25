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

'@return code メソッド実行直後の終了コードを返す。
Public Property Get code() As Long

End Property

'@param lngCode メソッド実行直後の終了コードを指定する。
'   0: 成功もエラーもなく何も実行せずに終了した状態を表す。
'   1: 異常終了した状態を表す。
'   2: メソッドの実行に成功した状態を表す。
Public Property Let code(lngCode As Long)

End Property

'@return message 異常終了時のエラーメッセージを返す。
Public Property Get message() As String

End Property

'@param strMessage 異常終了時のエラーメッセージを指定する。
Public Property Let message(strMessage As String)

End Property

'終了コードとエラーメッセージを初期化する。
Public Sub initStatus()

End Sub

'終了コードを 1 に自動設定し、エラーメッセージを出力する。
Public Sub errorTerminate()

End Sub
