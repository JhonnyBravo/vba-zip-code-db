VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private propName As String
Private propDefinition As String

'''
'@return filterName 拡張子フィルターの名前を返す。
'''
Public Property Get filterName() As String
    filterName = propName
End Property

'''
'@param name 拡張子フィルターの名前を指定する。
'   例) Excel
'''
Public Property Let filterName(name As String)
    propName = name
End Property

'''
'@return definition 拡張子フィルターの定義を返す。
'''
Public Property Get definition() As String
    definition = propDefinition
End Property

'''
'@param definition 拡張子フィルターの定義を指定する。
'   例) *.xls;*.xlsx;*.xlsm
'''
Public Property Let definition(definition As String)
    propDefinition = definition
End Property
