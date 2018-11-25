VERSION 1.0 CLASS
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

Implements ICommand
Private objSC As IStatus
Private propConnection As DAO.Database
Private propEntityName As String
Private propRecordset As DAO.recordset

'Properties

'@see ICommand#code
Public Property Get ICommand_code() As Long
    ICommand_code = objSC.code
End Property

'@return recordset 生成された Recordset を返す。
Public Property Get recordset() As DAO.recordset
    Set recordset = propRecordset
End Property

'Methods

Private Sub Class_Initialize()
    Set objSC = New Status
End Sub

Private Sub Class_Terminate()
    Set propRecordset = Nothing
End Sub

'@param connection 操作対象とする DataBase を指定する。
'@param entityName Recordset として取得するテーブルオブジェクトまたはクエリオブジェクトの名前を指定する。
Public Sub init(ByRef connection As DAO.Database, entityName As String)
    Set propConnection = connection
    propEntityName = entityName
End Sub

'Recordset を生成する。
Public Sub ICommand_run()
    objSC.initStatus
    
    On Error GoTo errExit
    
    Debug.Print propEntityName & " を開いています......"
    Set propRecordset = propConnection.OpenRecordset(propEntityName)
    
    objSC.code = 2
    Exit Sub
errExit:
    With objSC
        .message = "エラーが発生しました。 " & Err.description
        .errorTerminate
    End With
End Sub
