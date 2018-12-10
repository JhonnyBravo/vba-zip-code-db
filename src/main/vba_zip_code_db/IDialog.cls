VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDialog"
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

'FileDialog を生成する。
'
'@param strType 生成する FileDialog の種類を指定する。
'   file: ファイルピッカーダイアログを生成する。
'   directory: フォルダーピッカーダイアログを生成する。
'
'@return IDialog
Public Function initDialog(strType As String) As IDialog

End Function

'生成した FileDialog を起動する。
'
'@return IDialog
Public Function openDialog() As IDialog

End Function

'拡張子フィルターへ登録されたフィルター定義を全件削除する。
'
'@return IDialog
Public Function clearFilters() As IDialog

End Function

'拡張子フィルターへフィルター定義を登録する。
'
'@param strName 拡張子フィルターの名前を指定する。
'@param strDefinition 拡張子フィルターへ登録するフィルター定義を ``;`` 区切りで指定する。
'
'@return IDialog
Public Function addFilters(strName As String, strDefinition As String) As IDialog

End Function

'@return path FileDialog で選択したファイルまたはディレクトリのパスを返す。
Public Function getPath() As String

End Function
