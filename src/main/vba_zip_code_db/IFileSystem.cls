VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFileSystem"
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

'@param strPath 取得対象とするファイルまたはディレクトリのパスを指定する。
'@return Object 任意のオブジェクトを返す。
Public Function getItem(strPath As String) As Object

End Function
