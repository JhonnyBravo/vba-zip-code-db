Attribute VB_Name = "TestDirectoryResource"
Option Compare Database
Option Explicit

Private objFSO As New FileSystemObject

Public Sub testDir1()
    Dim basePath As String
    Dim fileName As String
    Dim path As String
    
    basePath = Application.CurrentProject.path
    fileName = "test_dir"
    path = objFSO.BuildPath(basePath, fileName)
    
    Dim objDR As DirectoryResource
    Set objDR = NewDirectoryResource(path)
    
    With objDR
        'ディレクトリを作成できること。
        .base.createItem
        
        If objFSO.FolderExists(path) Then
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
        
        'ディレクトリが既に存在する場合に終了コードが 0 であること。
        .base.createItem
        
        Debug.Print "code = " & .status.code
        
        If .status.code = 0 Then
            Debug.Print "OK"
        Else
            Debug.Print "NG"
        End If
        
        'ディレクトリの削除ができること。
        .base.deleteItem
        
        If objFSO.FolderExists(path) = False Then
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
        
        'ディレクトリが存在しない場合に終了コードが 0 であること。
        .base.deleteItem
        
        Debug.Print "code = " & .status.code
        
        If .status.code = 0 Then
            Debug.Print "OK"
        Else
            Debug.Print "NG"
        End If
    End With
End Sub

Public Sub testDir2()
    Dim basePath As String
    Dim fileName As String
    Dim path As String
    
    basePath = Application.CurrentProject.path
    fileName = "src\test\resources\csv"
    path = objFSO.BuildPath(basePath, fileName)
    
    Dim objDR As DirectoryResource
    Dim objFiles As Files
    Dim objFile As File
    
    'ファイルの読込ができること。
    Set objDR = NewDirectoryResource(path)
    Set objFiles = objDR.getFiles
    
    Debug.Print "code = " & objDR.status.code
    
    '終了コードが 2 であること。
    If objDR.status.code = 2 Then
        Debug.Print "OK"
    Else
        Debug.Print "NG"
    End If
    
    Debug.Print "ファイル数 = " & objFiles.Count
    
    If objFiles.Count = 2 Then
        Debug.Print "OK"
    Else
        Debug.Print "NG"
    End If
End Sub
