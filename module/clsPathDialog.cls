VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPathDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private objFD As FileDialog

Private propCode As Long
Private propMessage As String
Private propPathType As String
Private propExtensionFilter As Collection

'/**
' * コマンド終了直後の終了コードを返す。
' *
' * @return
' *     0: 変更も異常もなく終了したことを表す。
' *     1: 異常終了したことを表す。
' *     2: 操作対象に対して何らかの変更を加えて正常終了したことを表す。
' */
Public Property Get code() As Long
    code = propCode
End Property

'/**
' * @return コマンド終了直後の終了メッセージを返す。
' */
Public Property Get message() As String
    message = propMessage
End Property

'/**
' * 呼び出すダイアログの種類を指定する。
' *
' * @param strType
' *     file: ファイルピッカーを呼び出す。
' *     directory: フォルダーピッカーを呼び出す。
' */
Public Property Let pathType(strType As String)
    propPathType = strType
End Property


'/**
' * @return 呼び出すダイアログの種類を返す。
' */
Public Property Get pathType() As String
    pathType = propPathType
End Property

'/**
' * ファイルピッカー向けの拡張子フィルターを設定する。
' *
' * @param varArray
' *     Array(Description, ExtensionList)
' *
' *         Description: フィルターの説明。
' *         ExtensionList: 表示を許可する拡張子のリスト。
' */
Public Property Let extensionFilter(varArray As Variant)
    propExtensionFilter.Add varArray
End Property

'/**
' * @return ファイルピッカー向けの拡張子フィルターの定義を納めたコレクション。
' */
Public Property Get extensionFilter() As Collection
    Set extensionFilter = propExtensionFilter
End Property

'/**
' * クラス初期化処理。
' */
Private Sub Class_Initialize()
    Set propExtensionFilter = New Collection
End Sub

'/**
' * ファイルピッカーまたはフォルダーピッカーを起動し、
' * 選択したファイルまたはフォルダーのパスを返す。
' *
' * @return ファイルまたはフォルダーのパスを返す。
' */
Public Function getPath() As String
    Dim result As String
    Dim varArray As Variant
    
    '--▼ ファイルピッカー初期化処理。
    If pathType = "file" Then
        Set objFD = Application.FileDialog(msoFileDialogFilePicker)
        
        objFD.Filters.Clear
        
        If extensionFilter.Count > 0 Then
            For Each varArray In extensionFilter
                If UBound(varArray) <> 1 Then
                    propMessage = "clsPathDialog.getPath extensionFilter の値が不正です。"
                    GoTo errExit
                End If
                
                objFD.Filters.Add CStr(varArray(0)), CStr(varArray(1))
            Next
        End If
    '--▲ ------
    '--▼ フォルダーピッカー初期化処理
    ElseIf pathType = "directory" Then
        Set objFD = Application.FileDialog(msoFileDialogFolderPicker)
    Else
        propMessage = "clsPathDialog.getPath pathType の値が不正です。"
        GoTo errExit
    End If
    '--▲ ------
    
    With objFD
        .InitialFileName = Application.CurrentProject.path
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            result = .SelectedItems(1)
            propCode = 2
            propMessage = "clsPathDialog.getPath [" & result & "] パスを取得しました。"
            
            Debug.Print message
        Else
            propCode = 0
            propMessage = ""
        End If
    End With
    
    getPath = result
    Exit Function

errExit:
    propCode = 1
    getPath = result
    MsgBox message
End Function
