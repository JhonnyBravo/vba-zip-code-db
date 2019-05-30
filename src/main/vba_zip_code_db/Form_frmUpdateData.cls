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
    
    With Me
        If IsNull(.txtFilePath) Then
            MsgBox "ファイルパスを指定してください。"
            Exit Sub
        End If
    End With
    
    strPath = Me.txtFilePath
    
    Dim objDR As DirectoryResource
    Dim objFiles As Files
    Dim objFile As File
    
    Set objDR = NewDirectoryResource(strPath)
    Set objFiles = objDR.getFiles
        
    If objDR.status.code = 1 Then
        Exit Sub
    End If
    
    'tblUpdateData 初期化処理
    Dim objQR As QueryResource
    Dim objQD As DAO.QueryDef
    
    Set objQR = NewQueryResource(NewConnectionResource(), "deleteUpdateData")
    Set objQD = objQR.context.openContext
    
    If objQR.status.code = 1 Then
        Exit Sub
    End If
    
    objQD.Execute
    objQR.context.closeContext
    
    Dim objIC As ImportCsv
    
    '更新ファイルインポート処理
    For Each objFile In objFiles
        If LCase(objFile.path) Like "*del_*.csv" Or _
            LCase(objFile.path) Like "*add_*.csv" _
        Then
            Set objIC = New ImportCsv
            
            With objIC
                .fileResource = NewFileResource(objFile.path)
                .recordsetResource = NewRecordsetResource(NewConnectionResource(), "tblUpdateData")
                .import
                
                If .status.code = 1 Then
                    Exit Sub
                End If
            End With
        End If
    Next
    
    'tblAddressMaster 更新処理
    On Error GoTo errExit
    
    Set objQR = NewQueryResource(NewConnectionResource(), "deleteAddressMaster")
    Set objQD = objQR.context.openContext
    
    If objQR.status.code = 1 Then
        Exit Sub
    End If
    
    objQD.Execute
    objQR.context.closeContext
    
    Set objQR = NewQueryResource(NewConnectionResource(), "insertAddressMaster")
    Set objQD = objQR.context.openContext
        
    If objQR.status.code = 1 Then
        Exit Sub
    End If
    
    objQD.Execute
    objQR.context.closeContext
    
    Debug.Print "作業終了: " & Time
    MsgBox "CSV 取込が完了しました。"
    Exit Sub
    
errExit:
    Debug.Print "エラーが発生しました。 " & Err.Description
    MsgBox "エラーが発生しました。 " & Err.Description
End Sub

Private Sub btnSetFilePath_Click()
    Dim strPath As String
    Dim objDR As DialogResource
    
    Set objDR = NewDialogResource("directory")
    
    With objDR
        strPath = .getPath
        
        If .status.code = 2 Then
            Me.txtFilePath = strPath
        End If
    End With
End Sub

Private Sub Form_Load()
    Me.txtFilePath = Application.CurrentProject.path
End Sub
