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
    
    Dim objGF As IFileSystem

    Dim objFiles As files
    Dim objFile As File
    
    Set objGF = New GetFiles
    Set objFiles = objGF.getItem(strPath)
    
    If objGF.Status.getCode = 1 Then
        Exit Sub
    End If
    
    'tblUpdateData 初期化処理
    Dim objDAOC As IDao
    Dim objQD As DAO.QueryDef
    
    Set objDAOC = New DaoController
    Set objQD = objDAOC.openConnection.openQueryDef("deleteUpdateData").getQueryDef
    
    If objDAOC.Status.getCode = 1 Then
        Exit Sub
    End If
    
    objQD.Execute
    objDAOC.closeQueryDef
    
    Dim objGTS As IStream
    Dim objICSV As IImportCsv
    
    Dim objTS As TextStream
    Dim objRS As DAO.Recordset
    
    Set objGTS = New GetTextStream
    Set objICSV = New ImportCsv
    
    '更新ファイルインポート処理
    For Each objFile In objFiles
        If LCase(objFile.path) Like "*del_*.csv" Or _
            LCase(objFile.path) Like "*add_*.csv" _
        Then
            Set objTS = objGTS.openStream(objFile, "read").getStream
            
            If objGTS.Status.getCode = 1 Then
                Exit Sub
            End If
            
            Set objRS = objDAOC.openRecordset("tblUpdateData").getRecordset
            
            If objDAOC.Status.getCode = 1 Then
                Exit Sub
            End If
            
            objICSV.setRecordset(objRS).setStream(objTS).import
            
            If objICSV.Status.getCode = 1 Then
                Exit Sub
            End If
            
            objGTS.closeStream
            objDAOC.closeRecordset
        End If
    Next
    
    'tblAddressMaster 更新処理
    On Error GoTo errExit
    
    Set objQD = objDAOC.openQueryDef("deleteAddressMaster").getQueryDef
    
    If objDAOC.Status.getCode = 1 Then
        Exit Sub
    End If
    
    objQD.Execute
    objDAOC.closeQueryDef
    
    Set objQD = objDAOC.openQueryDef("insertAddressMaster").getQueryDef
    
    If objDAOC.Status.getCode = 1 Then
        Exit Sub
    End If
    
    objQD.Execute
    objDAOC.closeQueryDef
    
    Debug.Print "作業終了: " & Time
    MsgBox "CSV 取込が完了しました。"
    Exit Sub
    
errExit:
    Debug.Print "エラーが発生しました。 " & Err.description
    MsgBox "エラーが発生しました。 " & Err.description
End Sub

Private Sub btnSetFilePath_Click()
    Dim strPath As String
    Dim objDialog As IDialog
    
    Set objDialog = New DialogController
    strPath = objDialog.initDialog("directory").openDialog.getPath
    
    If objDialog.Status.getCode = 2 Then
        Me.txtFilePath = strPath
    End If
End Sub

Private Sub Form_Load()
    Me.txtFilePath = Application.CurrentProject.path
End Sub
