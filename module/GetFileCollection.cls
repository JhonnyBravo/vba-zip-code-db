VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetFileCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IStatus
Private objSC As New StatusController

Private Property Let IStatus_code(lngCode As Long)
    objSC.IStatus_code = lngCode
End Property

Public Property Get IStatus_code() As Long
    IStatus_code = objSC.IStatus_code
End Property

Private Property Let IStatus_message(strMessage As String)
    objSC.IStatus_message = strMessage
End Property

Private Property Get IStatus_message() As String
    IStatus_message = objSC.IStatus_message
End Property

'指定したディレクトリ配下に存在するファイル群を返す。
'
'@param strPath 走査対象とするディレクトリのパスを指定する。
'@return Files 指定したディレクトリ配下に存在するファイル群を返す。
Public Function runCommand(strPath As String) As Files
    objSC.initStatus
    
    Dim objFSO As New FileSystemObject
    Dim objFiles As Files
    
    With objFSO
        If .FolderExists(strPath) = False Then
            IStatus_message = strPath & " が見つかりません。"
            objSC.errorTerminate
            Exit Function
        Else
            Set objFiles = .GetFolder(strPath).Files
            IStatus_code = 2
            Set runCommand = objFiles
        End If
    End With
End Function
