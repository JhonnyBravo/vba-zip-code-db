﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private propCode As Long
Private objSC As New StatusController

'@return codeメソッド実行直後の終了コードを返す。
'   0: エラーもなく、リソースの変更もなく終了した状態を表す。
'   1: エラー終了した状態を表す。
'   2: リソースを変更し、正常終了した状態を表す。
Public Property Get code() As Long
    code = propCode
End Property

'@param strName レコードセット取得対象とするテーブル/クエリの名前を指定する。
'@return Recordset テーブル/クエリのレコードセットを返す。
Public Function runCommand(strName As String) As DAO.Recordset
    objSC.initStatus
    
    Dim objDB As DAO.Database
    Dim objRS As DAO.Recordset
    
    Set objDB = Application.CurrentDb
    
    On Error GoTo errExit
    
    Debug.Print strName & " を開きます。"
    
    Set objRS = objDB.openRecordset(strName)
    
    With objSC
        .code = 2
        .message = ""
        propCode = .code
    End With
    
    Set runCommand = objRS
    Exit Function
errExit:
    With objSC
        .message = "エラーが発生しました。 " & Err.Description
        .errorTerminate
        propCode = .code
    End With
End Function