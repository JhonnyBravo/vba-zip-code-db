VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'@see IStatus#code
Public Property Get code() As Long

End Property

'@param strKey キーとして使用する文字列を指定する。
'@param varValue パラメータ値として格納する値を指定する。
Public Sub setVariant(strKey As String, varValue As Variant)

End Sub

'@param strKey 取得対象とするパラメータ値のキーを指定する。
'@return Variant 指定したキーに紐づくパラメータ値を返す。
Public Function getVariant(strKey As String) As Variant

End Function

'@param strKey キーとして使用する文字列を指定する。
'@param objInstance パラメータ値として格納するオブジェクトを指定する。
Public Sub setObject(strKey As String, ByRef objInstance As Object)

End Sub

'@param strKey 取得対象とするパラメータ値のキーを指定する。
'@return Object 指定したキーに紐づくパラメータオブジェクトを返す。
Public Function getObject(strKey As String) As Object

End Function
