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

Implements ICommand

Private objSC As IStatus
Private objFSO As FileSystemObject

Private propDirectory As Folder
Private propFiles As files

'Properties

'@see ICommand#code
Public Property Get ICommand_code() As Long
    ICommand_code = objSC.code
End Property

'@retern files 取得した Files を返す。
Public Property Get files() As files
    Set files = propFiles
End Property

'Methods

Private Sub Class_Initialize()
    Set objSC = New Status
    Set objFSO = New FileSystemObject
End Sub

'@param objDirectory 操作対象とする Folder を指定する。
Public Sub init(ByRef objDirectory As Folder)
    Set propDirectory = objDirectory
End Sub

'指定した Folder 配下に存在する Files を取得する。
Public Sub ICommand_run()
    objSC.initStatus
    
    If objFSO.FolderExists(propDirectory.path) Then
        Debug.Print propDirectory.path & " 配下のファイル一覧を取得しています......"
        Set propFiles = propDirectory.files
        objSC.code = 2
    Else
        With objSC
            .message = propDirectory.path & " が見つかりません。"
            .errorTerminate
        End With
    End If
End Sub
