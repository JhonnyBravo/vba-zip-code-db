VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EntityDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private propEntity As String
Private objCollection As Collection

'@return 操作対象とするテーブルの名前を返す。
Public Property Get entity() As String
    entity = propEntity
End Property

'@param strEntity 操作対象とするテーブルの名前を指定する。
Public Property Let entity(strEntity As String)
    propEntity = strEntity
End Property

'フィールド定義を格納するコレクションを初期化する。
Public Sub initCollection()
    Set objCollection = New Collection
End Sub

'フィールド定義配列をコレクションへ追加する。
'
'@param lngIndex フィールドの列番号を指定する。
'@param strFieldName フィールドの名前を指定する。
Public Sub addFieldDefinition(lngIndex As Long, strFieldName As String)
    objCollection.Add Array(lngIndex, strFieldName)
End Sub

'@return Collection フィールド定義配列を格納したコレクションを返す。
Public Function getCollection() As Collection
    Set getCollection = objCollection
End Function

'フィールド定義を格納するコレクションを初期化する。
Private Sub Class_Initialize()
    initCollection
End Sub
