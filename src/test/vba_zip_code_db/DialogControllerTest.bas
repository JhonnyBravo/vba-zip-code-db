Attribute VB_Name = "DialogControllerTest"
Option Compare Database
Option Explicit

Public Sub dcTest1()
    Dim strPath As String
    Dim objDialog As IDialog
    
    Set objDialog = New DialogController
    strPath = objDialog.initDialog("file").clearFilters.addFilters("CSV", "*.csv").openDialog.getPath
    
    With objDialog.Status
        Select Case .getCode
            Case 0
                If strPath = "" Then
                    Debug.Print "dcTest1: OK"
                Else
                    Debug.Print "dcTest1: NG"
                End If
            Case 2
                If strPath <> "" Then
                    Debug.Print "dcTest1: OK"
                Else
                    Debug.Print "dcTest1: NG"
                End If
            Case Else
                Debug.Print "dcTest1: NG"
        End Select
        
        Debug.Print "code: " & .getCode & " path: " & strPath
    End With
End Sub

Public Sub dcTest2()
    Dim strPath As String
    Dim objDialog As IDialog
    
    Set objDialog = New DialogController
    strPath = objDialog.initDialog("file").clearFilters.addFilters("Excel", "*.xls;*.xlsx;*.xlsm").openDialog.getPath
    
    With objDialog.Status
        Select Case .getCode
            Case 0
                If strPath = "" Then
                    Debug.Print "dcTest2: OK"
                Else
                    Debug.Print "dcTest2: NG"
                End If
            Case 2
                If strPath <> "" Then
                    Debug.Print "dcTest2: OK"
                Else
                    Debug.Print "dcTest2: NG"
                End If
            Case Else
                Debug.Print "dcTest2: NG"
        End Select
        
        Debug.Print "code: " & .getCode & " path: " & strPath
    End With
End Sub

Public Sub dcTest3()
    Dim strPath As String
    Dim objDialog As IDialog
    
    Set objDialog = New DialogController
    strPath = objDialog.initDialog("directory").openDialog.getPath
    
    With objDialog.Status
        Select Case .getCode
            Case 0
                If strPath = "" Then
                    Debug.Print "dcTest3: OK"
                Else
                    Debug.Print "dcTest3: NG"
                End If
            Case 2
                If strPath <> "" Then
                    Debug.Print "dcTest3: OK"
                Else
                    Debug.Print "dcTest3: NG"
                End If
            Case Else
                Debug.Print "dcTest3: NG"
        End Select
        
        Debug.Print "code: " & .getCode & " path: " & strPath
    End With
End Sub
