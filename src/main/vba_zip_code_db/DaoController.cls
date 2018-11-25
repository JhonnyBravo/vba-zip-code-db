VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaoController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements ICommand

Private objStatus As IStatus
Private objGC As GetConnection
Private objGQD As GetQueryDef
Private objGR As getRecordset

'Properties

'@see ICommand#code
Public Property Get ICommand_code() As Long
    ICommand_code = objStatus.code
End Property

'Methods

Private Sub Class_Initialize()
    Set objStatus = New Status
    Set objGC = New GetConnection
    Set objGQD = New GetQueryDef
    Set objGR = New getRecordset
End Sub

'@param sql 取得対象とするクエリオブジェクトの名前または SQL を直接記述する。
'@return query QueryDef を生成して返す。
Public Function getQuery(sql As String) As DAO.QueryDef
    objStatus.initStatus
    
    With objGC
        .ICommand_run
        objGQD.init .connection, sql
    End With
    
    Dim objQD As DAO.QueryDef
    
    With objGQD
        .ICommand_run
        objStatus.code = .ICommand_code
        
        If objStatus.code = 1 Then
            Set getQuery = objQD
            Exit Function
        End If
        
        Set objQD = .query
    End With
    
    Set getQuery = objQD
End Function

'@param entityName Recordset として開くテーブルオブジェクトまたはクエリオブジェクトの名前を指定する。
'@return recordset Recordset を生成して返す。
Public Function getRecordset(entityName As String) As DAO.recordset
    objStatus.initStatus
    
    With objGC
        .ICommand_run
        objGR.init .connection, entityName
    End With
    
    Dim objRecordset As DAO.recordset
    
    With objGR
        .ICommand_run
        objStatus.code = .ICommand_code
        
        If objStatus.code = 1 Then
            Set getRecordset = objRecordset
            Exit Function
        End If
        
        Set objRecordset = .recordset
    End With
    
    Set getRecordset = objRecordset
End Function

Private Sub ICommand_run()

End Sub
