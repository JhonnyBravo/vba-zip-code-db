Attribute VB_Name = "DCTest"
Option Compare Database
Option Explicit

'DialogController#getCsvPath の動作確認
Public Sub DCTest1()
    Dim objDC As New DialogController
    
    With objDC
        Debug.Print "Path: " & .getCsvPath
        Debug.Print "Code: " & .IStatus_code
    End With
End Sub

'DialogController#getExcelPath の動作確認
Public Sub DCTest2()
    Dim objDC As New DialogController
    
    With objDC
        Debug.Print "Path: " & .getExcelPath
        Debug.Print "Code: " & .IStatus_code
    End With
End Sub

'DialogController#getDirectoryPath の動作確認
Public Sub DCTest3()
    Dim objDC As New DialogController
    
    With objDC
        Debug.Print "Path: " & .getDirectoryPath
        Debug.Print "Code: " & .IStatus_code
    End With
End Sub
