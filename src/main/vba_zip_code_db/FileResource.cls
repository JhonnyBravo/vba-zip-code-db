VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IBase
Implements IContext

Private propFSO As FileSystemObject
Private propStream As TextStream

Private propStatus As StatusResource
Private propPath As String

'''
'@return StatusResource
'''
Public Property Get status() As StatusResource
    Set status = propStatus
End Property

'''
'@return path 操作対象とするファイルのパスを返す。
'''
Public Property Get path() As String
    path = propPath
End Property

'''
'@param path 操作対象とするファイルのパスを指定する。
'''
Public Property Let path(path As String)
    propPath = path
End Property

'''
'@return IBase
'''
Public Property Get base() As IBase
    Set base = Me
End Property

'''
'@return IContext
'''
Public Property Get context() As IContext
    Set context = Me
End Property

Private Sub Class_Initialize()
    Set propStatus = New StatusResource
    Set propFSO = New FileSystemObject
End Sub

'''
'ファイルを作成する。
'''
Private Sub IBase_createItem()
    If propFSO.FileExists(path) Then
        status.code = 0
        Exit Sub
    Else
        status.printMessage 0, path & " を作成しています......"
        propFSO.CreateTextFile path
        status.code = 2
    End If
End Sub

'''
'ファイルを削除する。
'''
Private Sub IBase_deleteItem()
    If propFSO.FileExists(path) = False Then
        status.code = 0
    Else
        status.printMessage 0, path & " を削除しています......"
        propFSO.DeleteFile path
        status.code = 2
    End If
End Sub

'''
'TextStream を閉じる。
'''
Private Sub IContext_closeContext()
    If propStream Is Nothing Then
        status.code = 0
    Else
        status.printMessage 0, path & " を閉じています......"
        propStream.Close
        Set propStream = Nothing
        status.code = 2
    End If
End Sub

'''
'@return TextStream を開いて返す。
'''
Private Function IContext_openContext() As Object
    If propFSO.FileExists(path) = False Then
        status.printMessage 1, path & " が見つかりません。"
        Set IContext = propStream
        Exit Function
    Else
        status.printMessage 0, path & " を開いています......"
        Set propStream = propFSO.OpenTextFile(path, ForReading)
        Set IContext_openContext = propStream
        status.code = 2
    End If
End Function

Private Sub Class_Terminate()
    IContext_closeContext
End Sub
