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

Implements IStatus
Implements IExtensionFilters

Private objSC As New StatusController
Private objEF As New ExtensionFilters

Private Property Let IStatus_code(lngCode As Long)
    objSC.IStatus_code = lngCode
End Property

Public Property Get IStatus_code() As Long
    IStatus_code = objSC.IStatus_code
End Property

Private Property Let IStatus_message(strMessage As String)
    IStatus_message = strMessage
End Property

Private Property Get IStatus_message() As String
    IStatus_message = objSC.IStatus_message
End Property

Public Sub IExtensionFilters_addDefinition(strLabel As String, strExtension As String)
    objEF.addDefinition strLabel, strExtension
End Sub

Private Function IExtensionFilters_getCollection() As Collection
    Set IExtensionFilters_getCollection = objEF.getCollection
End Function

Public Sub IExtensionFilters_initCollection()
    objEF.initCollection
End Sub

'ファイルピッカーを起動し、選択したファイルのパスを取得して返す。
'
'@return path 選択したファイルのパスを返す。
Public Function runCommand() As String
    objSC.initStatus
    
    Dim objFD As FileDialog
    Set objFD = Application.FileDialog(msoFileDialogFilePicker)
    objFD.Filters.Clear
    
    Dim objFilterCollection As Collection
    Dim varFilter As Variant
    
    Set objFilterCollection = IExtensionFilters_getCollection
    
    If objFilterCollection.Count > 0 Then
        For Each varFilter In objFilterCollection
            objFD.Filters.Add varFilter(0), varFilter(1)
        Next
    End If
    
    Dim strPath As String
    
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
