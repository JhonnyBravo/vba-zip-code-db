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

Implements IReturnVariant

Private objStatus As IStatus
Private objParams As IParameters

Private objDialog As FileDialog
Private boolDialog As Boolean

'Properties

'@see IReturnVariant#code
Public Property Get IReturnVariant_code() As Long
    IReturnVariant_code = objStatus.code
End Property

'Methods

Private Sub Class_Initialize()
    Set objStatus = New Status
    Set objParams = New Parameters
    boolDialog = False
End Sub

'拡張子フィルターを空にする。
Public Sub clearFilters()
    objStatus.initStatus
    objDialog.Filters.Clear
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
    
    objDialog.Filters.Add description, definition
    objStatus.code = 2
    Exit Sub
    
errExit:
    With objStatus
        .message = "エラーが発生しました。 " & Err.description
        .errorTerminate
    End With
End Sub

'@return path FileDialog 上で選択したファイルまたはディレクトリのパスを取得して返す。
Public Function IReturnVariant_run() As Variant
    objStatus.initStatus
    
    If boolDialog = False Then
        With objStatus
            .message = "dialog が設定されていません。"
            .errorTerminate
            Exit Function
        End With
    End If
    
    Dim strPath As String
    
    With objDialog
        .AllowMultiSelect = False
        .InitialFileName = Application.CurrentProject.path

        If .Show = -1 Then
            strPath = .SelectedItems(1)
            objStatus.code = 2
        End If
    End With
    
    IReturnVariant_run = strPath
End Function

'@param strKey 有効なキー: dialog
'@param objInstance 操作対象とする FileDialog を指定する。
Public Sub IReturnVariant_setParamObject(strKey As String, objInstance As Object)
    If strKey = "dialog" Then
        boolDialog = True
        Set objDialog = objInstance
    End If
End Sub

Private Sub IReturnVariant_setParamValue(strKey As String, varValue As Variant)

End Sub
