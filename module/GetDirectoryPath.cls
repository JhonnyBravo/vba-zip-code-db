VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetDirectoryPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Database

Private objSC As New StatusController
Private propCode As Long

'@return code メソッド実行直後の終了コードを返す。
'   0: エラーもなく、リソースの変更もなく終了した状態を表す。
'   1: エラー終了した状態を表す。
'   2: リソースを変更し、正常終了した状態を表す。
Public Property Get code() As Long
    code = propCode
End Property

'ファイルピッカーを起動し、選択したディレクトリのパスを取得して返す。
'
'@return path 選択したディレクトリのパスを返す。
Public Function runCommand() As String
    Dim strPath As String
    strPath = ""
    
    objSC.initStatus
    
    Dim objFD As FileDialog
    Set objFD = Application.FileDialog(msoFileDialogFolderPicker)
    
    With objFD
        .InitialFileName = Application.CurrentProject.path
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            strPath = .SelectedItems(1)
            
            With objSC
                .code = 2
                .message = ""
            End With
        Else
            strPath = ""
            objSC.initStatus
        End If
    End With
    
    propCode = objSC.code
    runCommand = strPath
End Function
