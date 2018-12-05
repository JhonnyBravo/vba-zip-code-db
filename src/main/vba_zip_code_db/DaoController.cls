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

Implements IController

Private objStatus As IStatus
Private objGC As GetConnection
Private objGQD As GetQueryDef
Private objGR As getRecordset

'Properties

Public Property Get IController_code() As Long
    IController_code = objStatus.code
End Property

'Methods

Private Sub Class_Initialize()
    Set objStatus = New Status
    Set objGC = New GetConnection
    Set objGQD = New GetQueryDef
    Set objGR = New getRecordset
End Sub

'@param sql 取得対象とするクエリオブジェクトの名前または SQL を直接記述する。
'@return DAO.QueryDef を生成して返す。
Public Function getQuery(sql As String) As DAO.QueryDef
    objStatus.initStatus
    
    Dim objDB As DAO.Database
    Set objDB = objGC.IReturnObject_run
    
    Dim objQD As DAO.QueryDef
    
    With objGQD
        .IReturnObject_setParamObject "connection", objDB
        .IReturnObject_setParamValue "sql", sql
        Set objQD = .IReturnObject_run
        
        objStatus.code = .IReturnObject_code
        
        If objStatus.code = 1 Then
            Exit Function
        End If
    End With
    
    Set getQuery = objQD
End Function

'@param entityName Recordset として開くテーブルオブジェクトまたはクエリオブジェクトの名前を指定する。
'@return DAO.Recordset を生成して返す。
Public Function getRecordset(entityName As String) As DAO.recordset
    objStatus.initStatus
    
    Dim objDB As DAO.Database
    Dim objRecordset As DAO.recordset
    
    Set objDB = objGC.IReturnObject_run
    
    With objGR
        .IReturnObject_setParamObject "connection", objDB
        .IReturnObject_setParamValue "entityName", entityName
        Set objRecordset = .IReturnObject_run
        
        objStatus.code = .IReturnObject_code
        
        If objStatus.code = 1 Then
            Exit Function
        End If
    End With
    
    Set getRecordset = objRecordset
End Function
