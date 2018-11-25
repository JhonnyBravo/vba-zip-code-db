VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmUpdateData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnImportCSV_Click()
    Debug.Print "作業開始: " & Time
    
    'インポート対象ファイル群の取得
    Dim strPath As String
    
    strPath = Me.txtFilePath
    
    Dim objGF As New GetFiles
    Dim objFSO As New FileSystemObject
    
    Dim objFiles As files
    Dim objFile As File
    
    With objGF
        .init objFSO.GetFolder(strPath)
        .ICommand_run
        
        If .ICommand_code = 1 Then
            Exit Sub
        End If
        
        Set objFiles = .files
    End With
    
    'tblUpdateData 初期化処理
    Dim objDAOC As New DaoController
    Dim objQD As DAO.QueryDef
    
    With objDAOC
        Set objQD = .getQuery("deleteUpdateData")
        
        If .ICommand_code = 1 Then
            Exit Sub
        End If
    End With
    
    objQD.Execute
    
    Dim objGTS As New GetTextStream
    Dim objTS As TextStream
    Dim objRS As DAO.recordset
    
    Dim objICSV As New ImportCsv
    
    '更新ファイルインポート処理
    For Each objFile In objFiles
        If LCase(objFile.path) Like "*del_*.csv" Or _
            LCase(objFile.path) Like "*add_*.csv" _
        Then
            With objGTS
                .init objFile, "read"
                .ICommand_run
                
                If .ICommand_code = 1 Then
                    Exit Sub
                End If
                
                Set objTS = .stream
            End With
            
            With objDAOC
                Set objRS = .getRecordset("tblUpdateData")
                
                If .ICommand_code = 1 Then
                    Exit Sub
                End If
            End With
            
            With objICSV
                .init objRS, objTS
                .ICommand_run
            End With
            
            objTS.Close
            objRS.Close
        End If
    Next
    
    'tblAddressMaster 更新処理
    On Error GoTo errExit
    
    With objDAOC
        Set objQD = .getQuery("deleteAddressMaster")
        
        If .ICommand_code = 1 Then
            Exit Sub
        End If
        
        objQD.Execute
        
        Set objQD = .getQuery("insertAddressMaster")
        
        If .ICommand_code = 1 Then
            Exit Sub
        End If
        
        objQD.Execute
    End With
    
    Debug.Print "作業終了: " & Time
    MsgBox "CSV 取込が完了しました。"
    Exit Sub
    
errExit:
    Debug.Print "エラーが発生しました。 " & Err.description
    MsgBox "エラーが発生しました。 " & Err.description
End Sub

Private Sub btnSetFilePath_Click()
    Dim strPath As String
    Dim objDC As New DialogController
    
    With objDC
        strPath = .getDirectoryPath
        
        If .ICommand_code = 2 Then
            Me.txtFilePath = strPath
        End If
    End With
End Sub

Private Sub Form_Load()
    Me.txtFilePath = Application.CurrentProject.path
End Sub
