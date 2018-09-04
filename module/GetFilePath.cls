VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetFilePath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Database

Private objSC As New StatusController
Private propFilterCollection As Collection
Private propCode As Long

'@return code メソッド実行直後の終了コードを返す。
'   0: エラーもなく、リソースの変更もなく終了した状態を表す。
'   1: エラー終了した状態を表す。
'   2: リソースを変更し、正常終了した状態を表す。
Public Property Get code() As Long
    code = propCode
End Property

'@return ファイルピッカーの拡張子フィルターへ渡す配列を格納したコレクションを返す。
'   例: Array("Images", "*.gif; *.jpg; *.jpeg")
Public Property Get filterCollection() As Collection
    Set filterCollection = propFilterCollection
End Property

Private Sub Class_Initialize()
    Set propFilterCollection = New Collection
End Sub

'ファイルピッカーを起動し、選択したファイルのパスを取得して返す。
'
'@return path 選択したファイルのパスを返す。
Public Function runCommand() As String
    Dim strPath As String
    strPath = ""
    
    objSC.initStatus
    
    Dim objFD As FileDialog
    Set objFD = Application.FileDialog(msoFileDialogFilePicker)
    
    With objFD
        .Filters.Clear
    End With
    
    Dim varFilter As Variant
    
    If filterCollection.Count > 0 Then
        For Each varFilter In filterCollection
            If UBound(varFilter) <> 1 Then
                With objSC
                    .message = "拡張子フィルターの定義が不正です。"
                    .errorTerminate
                    propCode = .code
                End With
                
                Exit Function
            End If
            
            objFD.Filters.Add varFilter(0), varFilter(1)
        Next
    End If
    
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
