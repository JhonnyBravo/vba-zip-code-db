VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConnectionResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IContext

Private propPath As String
Private propDB As DAO.Database
Private propStatus As StatusResource

'''
'@return path 操作対象とする DB のパスを返す。
'''
Public Property Get path() As String
    path = propPath
End Property

'''
'@param path 操作対象とする DB のパスを指定する。
'''
Public Property Let path(path As String)
    propPath = path
End Property

'''
'@return StatusResource
'''
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
'DB 接続を切断する。
'''
Private Sub IContext_closeContext()
    status.code = 0
    
    If propDB Is Nothing = False Then
        status.printMessage 0, propDB.name & " を閉じています......"
        propDB.Close
        Set propDB = Nothing
        status.code = 2
    End If
End Sub

'''
'DB 接続を開く。
'
'@return DAO.Database
'''
Private Function IContext_openContext() As Object
    On Error GoTo catch
    
    If path <> "" Then
        status.printMessage 0, path & " を開いています......"
        Set propDB = OpenDatabase(path)
        status.code = 2
    Else
        status.printMessage 0, Application.CurrentDb.name & " を開いています......"
        Set propDB = Application.CurrentDb
        status.code = 2
    End If
    
    GoTo finally
    
catch:
    status.printMessage 1, "エラーが発生しました。 " & Err.Description
finally:
    Set IContext_openContext = propDB
End Function
