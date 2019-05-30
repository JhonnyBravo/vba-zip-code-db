Attribute VB_Name = "TestFileResource"
Option Compare Database
Option Explicit

Private objFSO As New FileSystemObject

Public Sub testFile1()
    Dim basePath As String
    Dim fileName As String
    Dim path As String
    
    basePath = Application.CurrentProject.path
    fileName = "test.txt"
    path = objFSO.BuildPath(basePath, fileName)
    
    Dim objFR As fileResource
    Set objFR = NewFileResource(path)
    
    With objFR
        'ファイルを作成できること。
        .base.createItem
        
        If objFSO.FileExists(path) Then
            Debug.Print "OK"
        Else
            Debug.Print "NG"
        End If
        
        '終了コードが 2 であること。
        Debug.Print "code = " & .status.code
        
        If .status.code = 2 Then
            Debug.Print "OK"
        Else
            Debug.Print "NG"
        End If
        
        'ファイルが既に存在する場合に終了コードが 0 であること。
        .base.createItem
        
        Debug.Print "code = " & .status.code
        
        If .status.code = 0 Then
            Debug.Print "OK"
        Else
            Debug.Print "NG"
        End If
        
        'ファイルの削除ができること。
        .base.deleteItem
        
        If objFSO.FileExists(path) = False Then
            Debug.Print "OK"
        Else
            Debug.Print "NG"
        End If
        
        '終了コードが 2 であること。
        Debug.Print "code = " & .status.code
        
        If .status.code = 2 Then
            Debug.Print "OK"
        Else
            Debug.Print "NG"
        End If
        
        'ファイルが存在しない場合に終了コードが 0 であること。
        .base.deleteItem
        
        Debug.Print "code = " & .status.code
        
        If .status.code = 0 Then
            Debug.Print "OK"
        Else
            Debug.Print "NG"
        End If
    End With
End Sub

Public Sub testFile2()
    Dim basePath As String
    Dim fileName As String
    Dim path As String
    
    basePath = Application.CurrentProject.path
    fileName = "src\test\resources\csv\ADD_test.csv"
    path = objFSO.BuildPath(basePath, fileName)
    
    Dim objFR As fileResource
    Dim objTS As TextStream
    Dim objLines As New Collection
    
    'ファイルの読込ができること。
    Set objFR = NewFileResource(path)
    Set objTS = objFR.context.openContext
    
    Debug.Print "code = " & objFR.status.code
    
    '終了コードが 2 であること。
    If objFR.status.code = 2 Then
        Debug.Print "OK"
    Else
        Debug.Print "NG"
    End If
    
    With objTS
        While .AtEndOfStream = False
            objLines.Add .ReadLine
        Wend
    End With
    
    Debug.Print "行数 = " & objLines.Count
    
    If objLines.Count = 8 Then
        Debug.Print "OK"
    Else
        Debug.Print "NG"
    End If
End Sub
