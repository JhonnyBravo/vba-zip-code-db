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

Implements ICommand

Private objSC As IStatus

Private propConnection As DAO.Database
Private propQuery As DAO.QueryDef
Private propSQL As String

'Properties

'@see ICommand#code
Public Property Get ICommand_code() As Long
    ICommand_code = objSC.code
End Property

'@return query 生成された QueryDef を返す。
Public Property Get query() As DAO.QueryDef
    Set query = propQuery
End Property

'Methods

Private Sub Class_Initialize()
    Set objSC = New Status
End Sub

'@param connection 操作対象とする Database を指定する。
'@param sql 取得対象とするクエリオブジェクトの名前を指定する。または SQL を直接記述する。
Public Sub init(ByRef connection As DAO.Database, sql As String)
    Set propConnection = connection
    propSQL = sql
End Sub

'QueryDef を生成する。
Public Sub ICommand_run()
    objSC.initStatus
    
    On Error GoTo errExit
    
    Debug.Print propSQL & " を開いています......"
    Set propQuery = propConnection.QueryDefs(propSQL)
    
    objSC.code = 2
    
    Exit Sub
errExit:
    With objSC
        .message = "エラーが発生しました。 " & Err.description
        .errorTerminate
    End With
End Sub
