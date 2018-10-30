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

Implements IStatus
Private objSC As New StatusController

Private Property Let IStatus_code(lngCode As Long)
    objSC.IStatus_code = lngCode
End Property

Public Property Get IStatus_code() As Long
    IStatus_code = objSC.IStatus_code
End Property

Private Property Let IStatus_message(strMessage As String)
    objSC.IStatus_message = strMessage
End Property

Private Property Get IStatus_message() As String
    IStatus_message = objSC.IStatus_message
End Property

'ファイルピッカーを起動し、選択したディレクトリのパスを取得して返す。
'
'@return path 選択したディレクトリのパスを返す。
Public Function runCommand() As String
    Dim strPath As String
    
    objSC.initStatus
    
    Dim objFD As FileDialog
    Set objFD = Application.FileDialog(msoFileDialogFolderPicker)
    
    With objFD
        .InitialFileName = Application.CurrentProject.path
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            strPath = .SelectedItems(1)
            IStatus_code = 2
        End If
    End With
    
    runCommand = strPath
End Function
