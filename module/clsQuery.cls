VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private objDB As DAO.Database
Private objQD As DAO.QueryDef

Private propCode As Long
Private propMessage As String
Private propQueryName As String
Private propParamCollection As Collection

'/**
' * コマンド終了直後の終了コードを返す。
' *
' * @return
' *     0: 変更も異常もなく終了したことを表す。
' *     1: 異常終了したことを表す。
' *     2: 操作対象に対して何らかの変更を加えて正常終了したことを表す。
' */
Public Property Get code() As Long
    code = propCode
End Property

'/**
' * @return コマンド終了直後の終了メッセージを返す。
' */
Public Property Get message() As String
    message = propMessage
End Property

'/**
' * 実行するクエリの名前を指定する。
' *
' * @param strName 実行するクエリの名前。
' */
Public Property Let queryName(strName As String)
    propQueryName = strName
End Property

'/**
' * @return クエリの名前。
' */
Public Property Get queryName() As String
    queryName = propQueryName
End Property

'/**
' * クエリのパラメータ名と値を紐づける配列を格納する。
' *
' * @param varArray
' *     Array(ParamName, ParamValue)
' *
' *         ParamName: パラメータの名前。
' *         ParamValue: パラメータへ渡す値。
' */
Public Property Let paramCollection(varArray As Variant)
    propParamCollection.Add varArray
End Property

'/**
' * @return テーブルのフィールド名と CSV ファイルの列番号を紐づける配列を格納したコレクション。
' */
Public Property Get paramCollection() As Collection
    Set paramCollection = propParamCollection
End Property

'/**
' * クラス初期化処理。
' */
Private Sub Class_Initialize()
    Set objDB = Application.CurrentDb
    Set propParamCollection = New Collection
End Sub

'/**
' * CSV から Access へデータをインポートする。
' */
Public Sub executeQuery()
    On Error GoTo errExit2
    
    Set objQD = objDB.QueryDefs(queryName)
    
    Dim varParam As Variant
    
    For Each varParam In paramCollection
        If UBound(varParam) <> 1 Then
            propMessage = "clsCSV.executeQuery paramCollection の値が不正です。"
            GoTo errExit1
        End If
        
        objQD.Parameters(varParam(0)).Value = varParam(1)
    Next
    
    objQD.Execute
    
    propCode = 2
    propMessage = "clsCSV.executeQuery [" & queryName & "] クエリを実行しました。"
    Debug.Print message
    Exit Sub
    
errExit1:
    propCode = 1
    MsgBox message
    Exit Sub

errExit2:
    propCode = 1
    propMessage = "clsCSV.executeQuery エラーが発生しました。" _
        & " コード: " & Err.Number _
        & " メッセージ: " & Err.Description
    MsgBox message
End Sub

'/**
' * クラス終了処理。
' */
Private Sub Class_Terminate()
    Set objDB = Nothing
End Sub
