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

Private propCode As Long
Private objSC As New StatusController

'@return code メソッド実行直後の終了コードを返す。
'
'   0: エラーもなく、リソースの変更もなく終了した状態を表す。
'   1: エラー終了した状態を表す。
'   2: リソースを変更し、正常終了した状態を表す。
Public Property Get code() As Long
    code = propCode
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
            With objSC
                .message = strPath & " が見つかりません。"
                .errorTerminate
                propCode = .code
            End With
            
            Exit Function
        Else
            Set objFiles = .GetFolder(strPath).Files
            
            With objSC
                .code = 2
                .message = ""
                propCode = .code
            End With
            
            Set runCommand = objFiles
        End If
    End With
End Function
