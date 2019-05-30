VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DirectoryResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IBase

Private propFSO As FileSystemObject

Private propStatus As StatusResource
Private propPath As String

'''
'@return StatusResource
'''
Public Property Get status() As StatusResource
    Set status = propStatus
End Property

'''
'@return path 操作対象とするディレクトリのパスを返す。
'''
Public Property Get path() As String
    path = propPath
End Property

'''
'@param path 操作対象とするディレクトリのパスを指定する。
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

Private Sub Class_Initialize()
    Set propStatus = New StatusResource
    Set propFSO = New FileSystemObject
End Sub

'''
'ディレクトリを作成する。
'''
Private Sub IBase_createItem()
    If propFSO.FolderExists(path) Then
        status.code = 0
        Exit Sub
    Else
        status.printMessage 0, path & " を作成しています......"
        propFSO.CreateFolder path
        status.code = 2
    End If
End Sub

'''
'ディレクトリを削除する。
'''
Private Sub IBase_deleteItem()
    If propFSO.FolderExists(path) = False Then
        status.code = 0
    Else
        status.printMessage 0, path & " を削除しています......"
        propFSO.DeleteFolder path
        status.code = 2
    End If
End Sub

'''
'@return files ディレクトリ直下に存在するファイルの一覧を取得する。
'''
Public Function getFiles() As Files
    Dim objFiles As Files
    
    If propFSO.FolderExists(path) = False Then
        status.printMessage 1, path & " が見つかりません。"
        Set getFiles = objFiles
    Else
        status.printMessage 0, "ファイル一覧を取得しています......"
        Set objFiles = propFSO.GetFolder(path).Files
        status.code = 2
        Set getFiles = objFiles
    End If
End Function
