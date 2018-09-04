VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetQueryDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private propCode As Long
Private objSC As New StatusController

'メソッド実行直後の終了コードを返す。
'   0: エラーもなく、リソースの変更もなく終了した状態を表す。
'   1: エラー終了した状態を表す。
'   2: リソースを変更し、正常終了した状態を表す。
Public Property Get code() As Long
    code = propCode
End Property

'@param strName QueryDef オブジェクトとして取得するクエリの名前を指定する。
'@return QueryDef を返す。
Public Function runCommand(strName As String) As DAO.QueryDef
    objSC.initStatus
    
    Dim objDB As DAO.Database
    Dim objQD As DAO.QueryDef
    
    Set objDB = Application.CurrentDb
    
    On Error GoTo errExit
    
    Debug.Print strName & " を実行します。"
    
    Set objQD = objDB.QueryDefs(strName)
    
    With objSC
        .code = 2
        .message = ""
        propCode = .code
    End With
    
    Set runCommand = objQD
    Exit Function
errExit:
    With objSC
        .message = "エラーが発生しました。 " & Err.Description
        .errorTerminate
        propCode = .code
    End With
End Function
