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
    Dim objDaoLib As New daoLib
    Dim objCsvLib As New csvLib
    
    Dim fileCollection As Collection
    Dim varFileName As Variant
    
    Dim strDirPath As String
    Dim strFilePath As String
    
    'ファイル一覧の取得
    strDirPath = Me.txtFilePath
    
    Set fileCollection = getFileCollection(strDirPath)
    
    'tblUpdateData からレコードを削除
    objDaoLib.invokeQuery ("deleteUpdateData")
     
    If objDaoLib.code <> 0 Then
        MsgBox objDaoLib.message
        Exit Sub
    End If

    Dim recordCollection As Collection
    Dim varRecord As Variant
    
    For Each varFileName In fileCollection
        strFilePath = strDirPath & "\" & CStr(varFileName)
        
        'tblUpdateData へレコードを登録する
        If LCase(CStr(varFileName)) Like "del_*.csv" Or _
            LCase(CStr(varFileName)) Like "add_*.csv" _
        Then
            Set recordCollection = objCsvLib.getRecordCollection(strFilePath)
            
            If objCsvLib.code <> 0 Then
                MsgBox objCsvLib.message
                Exit Sub
            End If
            
            For Each varRecord In recordCollection
                objDaoLib.insertRecord "insertUpdateData", DCount("ID", "tblUpdateData") + 1, _
                    CStr(varRecord(0)), CStr(varRecord(1)), CStr(varRecord(2)), _
                    CInt(varRecord(3)), CInt(varRecord(4)), Date
                
                If objDaoLib.code <> 0 Then
                    MsgBox objDaoLib.message
                    Exit Sub
                End If
            Next
        End If
    Next
    
    '削除対象レコードを tblAddressMaster から削除する。
    objDaoLib.invokeQuery ("deleteAddressMaster")
    
    If objDaoLib.code <> 0 Then
        MsgBox objDaoLib.message
        Exit Sub
    End If
    
    '追加対象レコードを tblAddressMaster へ登録する。
    objDaoLib.invokeQuery ("insertAddressMaster")
    
    If objDaoLib.code <> 0 Then
        MsgBox objDaoLib.message
        Exit Sub
    End If
    
    MsgBox "CSV 取込が完了しました。"
End Sub

Private Sub btnSetFilePath_Click()
    Dim strFilePath As String
    strFilePath = getDirPath
    
    If strFilePath <> "" Then
        Me.txtFilePath.Value = strFilePath
    End If
End Sub

Private Sub Form_Load()
    Me.txtFilePath = Application.CurrentProject.Path
End Sub
