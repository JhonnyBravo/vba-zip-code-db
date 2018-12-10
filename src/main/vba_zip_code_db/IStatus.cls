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
Public Function getCode() As Long

End Function

'@param lngCode メソッド実行直後の終了コードを指定する。
'   0: 失敗も成功もなく何もせずに終了した状態を表す。
'   1: 異常終了した状態を表す。
'   2: メソッドの実行に成功した状態を表す。
'
'@return IStatus
Public Function setCode(lngCode As Long) As IStatus

End Function

'終了コードを初期化する。
'
'@return IStatus
Public Function initStatus() As IStatus

End Function

'エラー発生時にダイアログとコンソールへメッセージを出力する。
'
'@param strMessage エラーメッセージとして表示する文字列を指定する。
'@return IStatus
Public Function errorTerminate(strMessage As String) As IStatus

End Function
