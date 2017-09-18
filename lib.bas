Attribute VB_Name = "lib"
Option Compare Database
Option Explicit

Public Function getDirPath() As String
    Dim result As String
    
    result = ""
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.CurrentProject.Path
        
        If .Show = -1 Then
            result = .SelectedItems(1)
        End If
        
    End With
    
    getDirPath = result
End Function

Public Function getFileCollection(strPath As String) As Collection
    Dim objFSO As New FileSystemObject
    Dim fileCollection As Files
    Dim objFile As File
    Dim result As New Collection
    
    Set fileCollection = objFSO.GetFolder(strPath).Files
    
    For Each objFile In fileCollection
        result.Add objFile.Name
    Next
    
    Set getFileCollection = result
End Function

