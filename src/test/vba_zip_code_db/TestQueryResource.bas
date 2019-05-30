Attribute VB_Name = "TestQueryResource"
Option Compare Database
Option Explicit

Public Sub testQuery1()
    Dim objQR As QueryResource
    Dim objQD As DAO.QueryDef
    
    Set objQR = NewQueryResource(NewConnectionResource(), "selectAddData")
    Set objQD = objQR.context.openContext
    
    With objQR
        Debug.Print "code = " & .status.code
        
        If .status.code = 2 Then
            Debug.Print "OK"
        Else
            Debug.Print "NG"
        End If
    End With
    
    If objQD.name = "selectAddData" Then
        Debug.Print "OK"
    Else
        Debug.Print "NG"
    End If
    
    Dim objRS As DAO.Recordset
    Set objRS = objQD.openRecordset
    
    With objRS
        .MoveLast
        .MoveFirst
    End With
    
    Debug.Print "レコード数 = " & objRS.RecordCount
End Sub
