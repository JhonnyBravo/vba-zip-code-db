VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'@return Boolean 次の要素が存在するかを確認する。
'   True: 次の要素が存在することを表す。
'   False: 次の要素が存在しないことを表す。
Public Function hasNext() As Boolean

End Function

'@return array 配列を生成して返す。
Public Function nextArray() As Variant

End Function
