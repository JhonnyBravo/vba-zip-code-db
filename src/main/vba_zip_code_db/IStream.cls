VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'IStatus の機能を継承するためのプロパティ。
'
'@return IStatus
Public Property Get Status() As IStatus

End Property

'TextStream を開く。
'
'@param objFile 操作対象とする File を指定する。
'@param strMode 取得対象とする TextStream の種類を指定する。
'   read: 読取専用で開く。
'   write: 上書きモードで開く。
'   append: 追記モードで開く。
'
'@return IStream
Public Function openStream(ByRef objFile As File, strMode As String) As IStream

End Function

'TextStream を閉じる。
'
'@return IStream
Public Function closeStream() As IStream

End Function

'@return TextStream 生成した TextStream を取得する。
Public Function getStream() As TextStream

End Function
