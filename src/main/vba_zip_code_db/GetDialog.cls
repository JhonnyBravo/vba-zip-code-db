VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements ICommand

Private objStatus As IStatus
Private propDialogType As String
Private propDialog As FileDialog

'Properties

'@see ICommand#code
Public Property Get ICommand_code() As Long
    ICommand_code = objStatus.code
End Property

'@return dialog 生成された FileDialog を返す。
Public Property Get dialog() As FileDialog
    Set dialog = propDialog
End Property

'Methods

Private Sub Class_Initialize()
    Set objStatus = New Status
End Sub

'@param dialogType 生成するダイアログの種類を指定する。
'   file: ファイルピッカーダイアログを生成する。
'   directory: フォルダーピッカーダイアログを生成する。
Public Sub init(dialogType As String)
    propDialogType = dialogType
End Sub

'ダイアログを生成する。
Public Sub ICommand_run()
    objStatus.initStatus
    
    If propDialogType = "file" Then
        Set propDialog = Application.FileDialog(msoFileDialogFilePicker)
        objStatus.code = 2
    ElseIf propDialogType = "directory" Then
        Set propDialog = Application.FileDialog(msoFileDialogFolderPicker)
        objStatus.code = 2
    Else
        With objStatus
            .message = "dialogType の値が不正です。"
            .errorTerminate
        End With
    End If
End Sub
