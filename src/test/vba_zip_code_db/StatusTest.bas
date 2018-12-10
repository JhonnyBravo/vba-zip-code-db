Attribute VB_Name = "StatusTest"
Option Compare Database
Option Explicit

Private objStatus As IStatus

Public Sub statusTest1()
    Set objStatus = New Status
    
    With objStatus
        .errorTerminate "エラーテスト"
        
        If .getCode = 1 Then
            Debug.Print "errorTerminate: OK"
        Else
            Debug.Print "errorTerminate: NG"
            Exit Sub
        End If
        
        .initStatus
        
        If .getCode = 0 Then
            Debug.Print "initStatus: OK"
        Else
            Debug.Print "initStatus: NG"
            Exit Sub
        End If
        
        .setCode 2
        
        If .getCode = 2 Then
            Debug.Print "code: OK"
        Else
            Debug.Print "code: NG"
            Exit Sub
        End If
    End With
End Sub
