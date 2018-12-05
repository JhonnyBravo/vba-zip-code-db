VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IReturnVariant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'@see IStatus#code
Public Property Get code() As Long

End Property

'@param strKey パラメータ値のキーとして使用する文字列を指定する。
'@param varValue パラメータ値として格納する値を指定する。
Public Sub setParamValue(strKey As String, varValue As Variant)

End Sub

'@param strKey パラメータ値のキーとして使用する文字列を指定する。
'@param objInstance パラメータ値として格納するオブジェクトを指定する。
Public Sub setParamObject(strKey As String, objInstance As Object)

End Sub

'@return 処理を実行し、値を生成して返す。
Public Function run() As Variant

End Function
