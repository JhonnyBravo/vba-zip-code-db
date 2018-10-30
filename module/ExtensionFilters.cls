VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtensionFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private objCollection As Collection

'拡張子フィルターの定義配列を納めるコレクションを初期化する。
Public Sub initCollection()
    Set objCollection = New Collection
End Sub

'拡張子フィルターの定義配列をコレクションへ追加する。
'
'@param strLabel 拡張子フィルターの定義名を指定する。
'   例) Excel
'@param strExtension 拡張子フィルターの捕捉対象とする拡張子を ``;`` 区切りで指定する。
'   例) *.xls;*.xlsx;*.xlsm
Public Sub addDefinition(strLabel As String, strExtension As String)
    Dim varDefinition As Variant
    
    varDefinition = Array(strLabel, strExtension)
    objCollection.Add varDefinition
End Sub

'拡張子フィルターの定義配列を納めたコレクションを返す。
'
'@return Collection 拡張子フィルターの定義配列を納めたコレクション。
Public Function getCollection() As Collection
    Set getCollection = objCollection
End Function

Private Sub Class_Initialize()
    initCollection
End Sub
