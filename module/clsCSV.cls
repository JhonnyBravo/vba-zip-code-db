VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private objDB As DAO.Database
Private objRS As DAO.Recordset

Private propCode As Long
Private propMessage As String
Private propPath As String
Private propDestination As String
Private propFieldCollection As Collection
Private propAutoNumber As Boolean

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
' * 取込み対象とする CSV ファイルのパスを指定する。
' *
' * @param strPath CSV ファイルのパス。
' */
Public Property Let path(strPath As String)
    propPath = strPath
End Property

'/**
' * @return CSV ファイルのパス。
' */
Public Property Get path() As String
    path = propPath
End Property

'/**
' * レコード登録先となるテーブルの名前を指定する。
' *
' * @param strDestination テーブルの名前。
' */
Public Property Let destination(strDestination As String)
    propDestination = strDestination
End Property

'/**
' * @return テーブルの名前。
' */
Public Property Get destination() As String
    destination = propDestination
End Property

'/**
' * テーブルのフィールド名と CSV ファイルの列番号を紐づける配列を格納する。
' *
' * @param varArray
' *     Array(FieldName, ColNumber)
' *
' *         FieldName: テーブルのフィールド名。
' *         ColNumber: CSV の列番号。(0 から始まる番号)
' */
Public Property Let fieldCollection(varArray As Variant)
    propFieldCollection.Add varArray
End Property

'/**
' * @return テーブルのフィールド名と CSV ファイルの列番号を紐づける配列を格納したコレクション。
' */
Public Property Get fieldCollection() As Collection
    Set fieldCollection = propFieldCollection
End Property

'/**
' * オートナンバーを有効にするかどうかを指定する。
' *
' * @param boolFlag
' *     True: オートナンバーを有効にする。
' *     False: オートナンバーを無効にする。(既定値)
' */
Public Property Let autoNumber(boolFlag As Boolean)
    propAutoNumber = boolFlag
End Property

'/**
' * @return オートナンバーの設定値。
' */
Public Property Get autoNumber() As Boolean
    autoNumber = propAutoNumber
End Property

'/**
' * クラス初期化処理。
' */
Private Sub Class_Initialize()
    Set objDB = Application.CurrentDb
    Set propFieldCollection = New Collection
    propAutoNumber = False
End Sub

'/**
' * レコードセットを開く。
' */
Private Sub openRecordset()
    On Error GoTo errExit
    
    Set objRS = objDB.openRecordset(destination)
    
    propCode = 2
    propMessage = "clsCSV.openRecordset [" & destination & "] レコードセットを開きました。"
    Exit Sub
    
errExit:
    propCode = 1
    propMessage = "clsCSV.openRecordset エラーが発生しました。" _
        & " コード: " & Err.Number _
        & " メッセージ: " & Err.Description
End Sub

'/**
' * レコードセットを閉じる。
' */
Private Sub closeRecordset()
    On Error GoTo errExit
    
    objRS.Close
    
    propCode = 2
    propMessage = "clsCSV.closeRecordset [" & destination & "] レコードセットを閉じました。"
    Exit Sub

errExit:
    propCode = 1
    propMessage = "clsCSV.closeRecordset エラーが発生しました。" _
        & " コード: " & Err.Number _
        & " メッセージ: " & Err.Description
End Sub

'/**
' * CSV から Access へデータをインポートする。
' */
Public Sub importRecord()
    On Error GoTo errExit2
    
    openRecordset
    
    If code = 1 Then
        GoTo errExit1
    Else
        Debug.Print message
    End If
    
    Dim strLine As String
    Dim lngFileNumber As Long
    Dim varRecord As Variant
    Dim varField As Variant
    
    lngFileNumber = FreeFile
    Open path For Input As #lngFileNumber
    
    Do Until EOF(lngFileNumber)
        Line Input #lngFileNumber, strLine
        
        varRecord = Split(Replace(strLine, """", ""), ",")
        
        With objRS
            .AddNew
            
            For Each varField In fieldCollection
                If UBound(varField) <> 1 Then
                    propMessage = "clsCSV.importRecord fieldCollection の値が不正です。"
                    GoTo errExit1
                End If
                
                If autoNumber Then
                    .Fields("ID").Value = .RecordCount + 1
                End If
                
                If varRecord(varField(1)) <> "" Then
                    .Fields(varField(0)).Value = varRecord(varField(1))
                End If
            Next
            
            .Update
        End With
    Loop
    
    Close lngFileNumber
    
    closeRecordset
    
    If code = 1 Then
        MsgBox message
        Exit Sub
    Else
        Debug.Print message
    End If
    
    propCode = 2
    propMessage = "clsCSV.importRecord [" & path & "] レコードをインポートしました。"
    Debug.Print message
    Exit Sub
    
errExit1:
    MsgBox message
    
    closeRecordset
    propCode = 1
    Exit Sub

errExit2:
    propMessage = "clsCSV.importRecord エラーが発生しました。" _
        & " コード: " & Err.Number _
        & " メッセージ: " & Err.Description
    MsgBox message
    
    closeRecordset
    propCode = 1
End Sub

'/**
' * クラス終了処理。
' */
Private Sub Class_Terminate()
    Set objRS = Nothing
    Set objDB = Nothing
End Sub
