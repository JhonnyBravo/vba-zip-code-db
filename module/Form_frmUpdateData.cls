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
    Dim objCsvLib As New csvLib
    
    Dim fileCollection As Collection
    Dim varFileName As Variant
    
    Dim strDirPath As String
    Dim strFilePath As String
    
    'ファイル一覧の取得
    strDirPath = Me.txtFilePath
    
    Set fileCollection = getFileCollection(strDirPath)
    
    'tblUpdateData からレコードを削除
    Dim objQuery As New clsQuery
    
    With objQuery
        .queryName = "deleteUpdateData"
        .executeQuery
        
        If .code = 1 Then
            Exit Sub
        End If
    End With

    Dim recordCollection As Collection
    Dim varRecord As Variant
    
    For Each varFileName In fileCollection
        strFilePath = strDirPath & "\" & CStr(varFileName)
        
        'tblUpdateData へレコードを登録する
        If LCase(CStr(varFileName)) Like "del_*.csv" Or _
            LCase(CStr(varFileName)) Like "add_*.csv" _
        Then
            Dim objCSV As New clsCSV
            
            With objCSV
                .path = strFilePath
                .destination = "tblUpdateData"
                .autoNumber = True
                .fieldCollection = Array("郵便番号", 2)
                .fieldCollection = Array("都道府県名-カナ", 3)
                .fieldCollection = Array("市区町村名-カナ", 4)
                .fieldCollection = Array("町域名-カナ", 5)
                .fieldCollection = Array("都道府県名", 6)
                .fieldCollection = Array("市区町村名", 7)
                .fieldCollection = Array("町域名", 8)
                .fieldCollection = Array("更新の表示", 13)
                .fieldCollection = Array("変更理由", 14)
                .importRecord
                
                If .code = 1 Then
                    Exit Sub
                End If
            End With
        End If
    Next
    
    '削除対象レコードを tblAddressMaster から削除する。
    With objQuery
        .queryName = "deleteAddressMaster"
        .executeQuery
        
        If .code = 1 Then
            Exit Sub
        End If
    End With
    
    '追加対象レコードを tblAddressMaster へ登録する。
    With objQuery
        .queryName = "insertAddressMaster"
        .executeQuery
        
        If .code = 1 Then
            Exit Sub
        End If
    End With
    
    MsgBox "CSV 取込が完了しました。"
End Sub

Private Sub btnSetFilePath_Click()
    Dim objPD As New clsPathDialog
    Dim strPath As String
    
    With objPD
        .pathType = "directory"
        strPath = .getPath
        
        If .code = 2 Then
            Me.txtFilePath = strPath
        End If
    End With
End Sub

Private Sub Form_Load()
    Me.txtFilePath = Application.CurrentProject.path
End Sub
