VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IFileSystem

Private objStatus As IStatus
Private objFSO As FileSystemObject

Private Sub Class_Initialize()
    Set objStatus = New Status
    Set objFSO = New FileSystemObject
End Sub

'指定したディレクトリ配下に存在するファイルのパスを一覧取得する。
'
'@param strPath ファイル一覧取得対象とするディレクトリのパスを指定する。
'@return Files
Public Function IFileSystem_getItem(strPath As String) As Object
    objStatus.initStatus
    
    Dim objFiles As files
    
    With objFSO
        If .FolderExists(strPath) = False Then
            objStatus.errorTerminate strPath & " が見つかりません。"
        Else
            Set objFiles = .GetFolder(strPath).files
            objStatus.setCode 2
        End If
    End With
    
    Set IFileSystem_getItem = objFiles
End Function

'@see IFileSystem#Status
Public Property Get IFileSystem_Status() As IStatus
    Set IFileSystem_Status = objStatus
End Property
