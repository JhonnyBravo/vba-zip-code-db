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

Implements IReturnObject

Private objStatus As IStatus
Private objParams As IParameters

'Properties

'@see IReturnObject#code
Public Property Get IReturnObject_code() As Long
    IReturnObject_code = objStatus.code
End Property

'Methods

Private Sub Class_Initialize()
    Set objStatus = New Status
    Set objParams = New Parameters
End Sub

'@return FileDialog FileDialog を生成して返す。
Public Function IReturnObject_run() As Object
    objStatus.initStatus
    
    Dim strType As String
    
    strType = objParams.getVariant("dialogType")
    
    With objStatus
        .code = objParams.code
        If .code = 1 Then
            Exit Function
        End If
    End With
    
    Dim objDialog As FileDialog
    
    If strType = "file" Then
        Set objDialog = Application.FileDialog(msoFileDialogFilePicker)
        objStatus.code = 2
    ElseIf strType = "directory" Then
        Set objDialog = Application.FileDialog(msoFileDialogFolderPicker)
        objStatus.code = 2
    Else
        With objStatus
            .message = "dialogType の値が不正です。"
            .errorTerminate
            Exit Function
        End With
    End If
    
    Set IReturnObject_run = objDialog
End Function

Private Sub IReturnObject_setParamObject(strKey As String, objInstance As Object)

End Sub

'@param strKey 有効なキー: dialogType
'@param varValue
'   file: ファイルピッカーダイアログを生成する。
'   directory: フォルダーピッカーダイアログを生成する。
Public Sub IReturnObject_setParamValue(strKey As String, varValue As Variant)
    objStatus.initStatus
    objParams.setVariant strKey, varValue
    objStatus.code = objParams.code
End Sub
