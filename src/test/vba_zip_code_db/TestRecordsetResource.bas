Attribute VB_Name = "TestRecordsetResource"
Option Compare Database
Option Explicit

Public Sub testRecordset1()
    Dim objRR As recordsetResource
    Dim objRS As DAO.Recordset
    
    Set objRR = NewRecordsetResource(NewConnectionResource(), "tblAddressMaster")
    Set objRS = objRR.context.openContext
    
    With objRR
        Debug.Print "code = " & .status.code
        
        If .status.code = 2 Then
            Debug.Print "OK"
        Else
            Debug.Print "NG"
        End If
    End With
    
    If objRS.name = "tblAddressMaster" Then
        Debug.Print "OK"
    Else
        Debug.Print "NG"
    End If
    
    Debug.Print "レコード数 = " & objRS.RecordCount
End Sub
