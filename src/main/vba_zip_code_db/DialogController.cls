VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DialogController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IDialog

Private objStatus As IStatus
Private objFD As FileDialog
Private strPath As String

Private Sub Class_Initialize()
    Set objStatus = New Status
End Sub

'@see IDialog#addFilters
Public Function IDialog_addFilters(strName As String, strDefinition As String) As IDialog
    objStatus.initStatus
    
    On Error GoTo catch
    
    objFD.Filters.Add strName, strDefinition
    objStatus.setCode 2
    GoTo finally
catch:
    objStatus.errorTerminate "エラーが発生しました。 " & Err.description
finally:
    Set IDialog_addFilters = Me
End Function

'@see IDialog#clearFilters
Public Function IDialog_clearFilters() As IDialog
    objStatus.initStatus
    
    On Error GoTo catch
    
    objFD.Filters.Clear
    objStatus.setCode 2
    GoTo finally
catch:
    objStatus.errorTerminate "エラーが発生しました。 " & Err.description
finally:
    Set IDialog_clearFilters = Me
End Function

'@see IDialog#getPath
Public Function IDialog_getPath() As String
    IDialog_getPath = strPath
End Function

'@see IDialog#initDialog
Public Function IDialog_initDialog(strType As String) As IDialog
    objStatus.initStatus
    
    If strType = "file" Then
        Set objFD = Application.FileDialog(msoFileDialogFilePicker)
        objStatus.setCode 2
    ElseIf strType = "directory" Then
        Set objFD = Application.FileDialog(msoFileDialogFolderPicker)
        objStatus.setCode 2
    Else
        objStatus.errorTerminate "strType の値が不正です。"
    End If
    
    Set IDialog_initDialog = Me
End Function

'@see IDialog#openDialog
Public Function IDialog_openDialog() As IDialog
    objStatus.initStatus
    strPath = ""
    
    On Error GoTo catch
    
    With objFD
        .AllowMultiSelect = False
        .InitialFileName = Application.CurrentProject.path
        
        If .Show = -1 Then
            strPath = .SelectedItems(1)
            objStatus.setCode 2
        End If
    End With
    
    GoTo finally
catch:
    objStatus.errorTerminate "エラーが発生しました。 " & Err.description
finally:
    Set IDialog_openDialog = Me
End Function

'@see IDialog#Status
Public Property Get IDialog_Status() As IStatus
    Set IDialog_Status = objStatus
End Property
