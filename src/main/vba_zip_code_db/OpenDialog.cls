VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OpenDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements ICommand

Private objStatus As IStatus
Private propDialog As FileDialog
Private propPath As String

'Properties

'@see ICommand#code
Public Property Get ICommand_code() As Long
    ICommand_code = objStatus.code
End Property

'@return ダイアログから取得したパスを返す。
Public Property Get path() As String
    path = propPath
End Property

'Methods

Private Sub Class_Initialize()
    Set objStatus = New Status
End Sub

'@param dialog 操作対象とする FileDialog を指定する。
Public Sub init(ByRef dialog As FileDialog)
    Set propDialog = dialog
End Sub

'拡張子フィルターを空にする。
Public Sub clearFilters()
    objStatus.initStatus
    propDialog.Filters.Clear
    objStatus.code = 2
End Sub

'拡張子フィルターを登録する。
'
'@param description 拡張子フィルターの名前を指定する。
'   例) Excel
'@param definition 拡張子フィルターの定義を ``;`` 区切りで指定する。
'   例) *.xls;*.xlsx;*.xlsm
Public Sub addFilters(description As String, definition As String)
    objStatus.initStatus
    
    On Error GoTo errExit
    
    propDialog.Filters.Add description, definition
    objStatus.code = 2
    Exit Sub
    
errExit:
    With objStatus
        .message = "エラーが発生しました。 " & Err.description
        .errorTerminate
    End With
End Sub

'FileDialog を開く。
Public Sub ICommand_run()
    objStatus.initStatus
    
    With propDialog
        .AllowMultiSelect = False
        .InitialFileName = Application.CurrentProject.path
        
        If .Show = -1 Then
            propPath = .SelectedItems(1)
            objStatus.code = 2
        Else
            propPath = ""
        End If
    End With
End Sub
