VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StatusResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private propCode As Integer
Private propMessage As String

'''
'@return code メソッドの終了コードを返す。
'''
Public Property Get code() As Integer
    code = propCode
End Property

'''
'@param code メソッドの終了コードを指定する。
'   * 0: 成功も失敗もなく終了した状態を表す。
'   * 1: 異常終了した状態を表す。
'   * 2: メソッドの実行に成功した状態を表す。
'''
Public Property Let code(code As Integer)
    propCode = code
End Property

'''
'@return message 終了メッセージを返す。
'''
Public Property Get message() As String
    message = propMessage
End Property

'''
'@param message 終了メッセージを指定する。
'''
Public Property Let message(message As String)
    propMessage = message
End Property

'''
'終了メッセージをコンソールへ出力する。
'
'@param code 終了コードとして設定する整数を指定する。
'@param message 終了メッセージとして設定する文字列を指定する。
'''
Public Sub printMessage(code As Integer, message As String)
    With Me
        .code = code
        .message = message
        Debug.Print .message
        
        If .code = 1 Then
            MsgBox .message
        End If
    End With
End Sub
