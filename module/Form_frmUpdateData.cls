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
    
    Dim objGFC As New GetFileCollection
    Dim objFiles As Files
    Dim objFile As File
    
    With objGFC
        Set objFiles = .runCommand(strPath)
        
        If .code = 1 Then
            Exit Sub
        End If
    End With

    'tblUpdateData 初期化処理
    Dim objT01Ctrl As New Tbl01Controller
    
    With objT01Ctrl
        .deleteRecord
        
        If .code = 1 Then
            Exit Sub
        End If
    End With
    
    '更新ファイルインポート処理
    For Each objFile In objFiles
        If LCase(objFile.path) Like "*del_*.csv" Or _
            LCase(objFile.path) Like "*add_*.csv" _
        Then
            With objT01Ctrl
                .importRecord objFile.path
                
                If .code = 1 Then
                    Exit Sub
                End If
            End With
        End If
    Next
    
    'tblAddressMaster 更新処理
    On Error GoTo errExit
    
    Dim objGQD As New GetQueryDef
    Dim objQD As DAO.QueryDef
    
    With objGQD
        Set objQD = .runCommand("deleteAddressMaster")
        
        If .code = 1 Then
            Exit Sub
        End If

        objQD.Execute
        
        Set objQD = .runCommand("insertAddressMaster")
        
        If .code = 1 Then
            Exit Sub
        End If

        objQD.Execute
    End With
    
    Debug.Print "作業終了: " & Time
    MsgBox "CSV 取込が完了しました。"
    Exit Sub
    
errExit:
    Debug.Print "エラーが発生しました。 " & Err.Description
    MsgBox "エラーが発生しました。 " & Err.Description
End Sub

Private Sub btnSetFilePath_Click()
    Dim strPath As String
    Dim objDC As New DialogController
    
    With objDC
        strPath = .getDirectoryPath
        
        If .code = 2 Then
            Me.txtFilePath = strPath
        End If
    End With
End Sub

Private Sub Form_Load()
    Me.txtFilePath = Application.CurrentProject.path
End Sub
