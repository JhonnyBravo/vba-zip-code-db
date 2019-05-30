Attribute VB_Name = "TestConnectionResource"
Option Compare Database
Option Explicit

Public Sub testConnection1()
    Dim objCR As ConnectionResource
    Dim objDB As DAO.Database
    
    Set objCR = NewConnectionResource
    
    With objCR
        Set objDB = .context.openContext
        
        Debug.Print "code = " & .status.code
        
        If .status.code = 2 Then
            Debug.Print "OK"
        Else
            Debug.Print "NG"
        End If
    End With
    
    If objDB.name = Application.CurrentDb.name Then
        Debug.Print "OK"
    Else
        Debug.Print "NG"
    End If
End Sub

Public Sub testConnection2()
    Dim objFSO As New FileSystemObject
    Dim strBasePath As String
    Dim strFileName As String
    Dim strPath As String
    
    strBasePath = Application.CurrentProject.path
    strFileName = "src\test\resources\vba-zip-code-db_v3.0.0.accdb"
    strPath = objFSO.BuildPath(strBasePath, strFileName)
    
    Dim objCR As ConnectionResource
    Dim objDB As DAO.Database
    
    Set objCR = NewConnectionResource(strPath)
    Set objDB = objCR.context.openContext
    
    With objCR
        Debug.Print "code = " & .status.code
        
        If .status.code = 2 Then
            Debug.Print "OK"
        Else
            Debug.Print "NG"
        End If
    End With
    
    If objDB.name = strPath Then
        Debug.Print "OK"
    Else
        Debug.Print "NG"
    End If
End Sub
