VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueryResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IContext

Private propStatus As StatusResource
Private propQD As DAO.QueryDef
Private propConnection As IContext
Private propEntityName As String

'''
'@return ConnectionResource
'''
Public Property Get connection() As ConnectionResource
    Set connection = propConnection
End Property

'''
'@param connection 操作対象とする ConnectionResource を指定する。
'''
Public Property Let connection(ByRef connection As ConnectionResource)
    Set propConnection = connection
End Property

'''
'@return entityName
'''
Public Property Get entityName() As String
    entityName = propEntityName
End Property

'''
'@param entityName 操作対象とするクエリの名前を指定する。
'''
Public Property Let entityName(entityName As String)
    propEntityName = entityName
End Property

Public Property Get status() As StatusResource
    Set status = propStatus
End Property

'''
'@return IContext
'''
Public Property Get context() As IContext
    Set context = Me
End Property

Private Sub Class_Initialize()
    Set propStatus = New StatusResource
End Sub

Private Sub Class_Terminate()
    IContext_closeContext
End Sub

'''
'クエリを閉じる。
'''
Private Sub IContext_closeContext()
    status.code = 0
    
    If propQD Is Nothing = False Then
        status.printMessage 0, propQD.name & " を閉じています......"
        propQD.Close
        Set propQD = Nothing
        status.code = 2
    End If
    
    Set propConnection = Nothing
End Sub

'''
'@return DAO.QueryDef クエリを開いて返す。
'''
Private Function IContext_openContext() As Object
    Dim objDB As DAO.Database
    Set objDB = connection.context.openContext
    status.code = connection.status.code
    
    If status.code = 1 Then
        Set IContext_openContext = propQD
        Exit Function
    End If
    
    On Error GoTo catch
    
    status.printMessage 0, entityName & " を開いています......"
    Set propQD = objDB.QueryDefs(entityName)
    status.code = 2
    GoTo finally
    
catch:
    status.printMessage 1, "エラーが発生しました。 " & Err.Description
finally:
    Set IContext_openContext = propQD
End Function
