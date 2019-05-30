VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DialogResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private propStatus As StatusResource
Private propFD As FileDialog
Private propType As String
Private propFilters As Collection

'''
'@return dialogType 生成するダイアログの種類を返す。
'''
Public Property Get dialogType() As String
    dialogType = propType
End Property

'''
'@param dialogType 生成するダイアログの種類を指定する。
'   * file: ファイルピッカーダイアログを生成する。
'   * directory: ディレクトリピッカーダイアログを生成する。
'''
Public Property Let dialogType(dialogType As String)
    propType = dialogType
End Property

'''
'@return StatusResource
'''
Public Property Get status() As StatusResource
    Set status = propStatus
End Property

Private Sub Class_Initialize()
    Set propStatus = New StatusResource
    Set propFilters = New Collection
End Sub

'''
'ファイルピッカーダイアログを生成する。
'''
Private Sub openFileDialog()
    Set propFD = Application.FileDialog(msoFileDialogFilePicker)
End Sub

'''
'ディレクトリピッカーダイアログを生成する。
'''
Private Sub openDirectoryDialog()
    Set propFD = Application.FileDialog(msoFileDialogFolderPicker)
End Sub

'''
'拡張子フィルターを登録する。
'
'@param resource 登録対象とする FilterResource を指定する。
'''
Public Sub addFilter(ByRef resource As FilterResource)
    propFilters.Add resource
End Sub

'''
'@return path ダイアログ上で選択したファイルまたはディレクトリのパスを返す。
'''
Public Function getPath() As String
    Dim strPath As String
    
    If dialogType = "file" Then
        status.printMessage 0, "ファイルピッカーを起動しています......"
        openFileDialog
    ElseIf dialogType = "directory" Then
        status.printMessage 0, "ディレクトリピッカーを起動しています......"
        openDirectoryDialog
    Else
        status.printMessage 1, "不正な dialogType です。 file または directory を指定してください。"
        getPath = strPath
        Exit Function
    End If
    
    Dim objFR As FilterResource
    
    With propFD
        .AllowMultiSelect = False
        .InitialFileName = Application.CurrentProject.path
        .filters.Clear
        
        If propFilters.Count > 0 Then
            For Each objFR In propFilters
                .filters.Add objFR.filterName, objFR.definition
            Next
        End If
        
        If .Show = -1 Then
            strPath = .SelectedItems(1)
            status.code = 2
        End If
    End With
    
    getPath = strPath
End Function
