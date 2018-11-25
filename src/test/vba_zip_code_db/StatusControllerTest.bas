Attribute VB_Name = "StatusControllerTest"
Option Compare Database
Option Explicit

Private objSC As New StatusController

Public Sub initStatusTest()
    With objSC
        .initStatus
        
        If .IStatus_code = 0 And .IStatus_message = "" Then
            Debug.Print "initStatus: OK" _
                & " code: " & .IStatus_code _
                & " message: " & .IStatus_message
        Else
            Debug.Print "initStatus: NG" _
                & " code: " & .IStatus_code _
                & " message: " & .IStatus_message
        End If
    End With
End Sub

Public Sub errorTerminateTest()
    With objSC
        .IStatus_message = "エラーが発生しました。"
        .errorTerminate
        
        If .IStatus_code = 1 And .IStatus_message = "エラーが発生しました。" Then
            Debug.Print "errorTerminate: OK" _
                & " code: " & .IStatus_code _
                & " message: " & .IStatus_message
        Else
            Debug.Print "errorTerminate: NG" _
                & " code: " & .IStatus_code _
                & " message: " & .IStatus_message
        End If
    End With
End Sub
