VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tbl01Controller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private propCode As Long
Private objSC As New StatusController

'@return code メソッド実行直後の終了コードを返す。
'   0: エラーもなく、リソースの変更もなく終了した状態を表す。
'   1: エラー終了した状態を表す。
'   2: リソースを変更し、正常終了した状態を表す。
Public Property Get code() As Long
    code = propCode
End Property

'tblUpdateData からレコードを全件削除する。
Public Sub deleteRecord()
    objSC.initStatus
    
    Dim objGQD As New GetQueryDef
    Dim objQD As DAO.QueryDef
    
    With objGQD
        Set objQD = .runCommand("deleteUpdateData")
        propCode = .code
    End With
    
    If code = 1 Then
        Exit Sub
    End If
    
    On Error GoTo errExit
    
    objQD.Execute
    Exit Sub
errExit:
    With objSC
        .message = "エラーが発生しました。 " & Err.Description
        .errorTerminate
        propCode = .code
    End With
End Sub

'tblUpdateData へレコードを登録する。
'{add, del}_yymm.csv -> tblUpdateData
'
'@param strPath {add, del}_yymm.csv のパスを指定する。
Public Sub importRecord(strPath As String)
    objSC.initStatus
    
    Dim objGTS As New GetTextStream
    Dim objTS As TextStream
    
    With objGTS
        Set objTS = .runCommand(strPath)
        propCode = .code
    End With
    
    If code = 1 Then
        Exit Sub
    End If
    
    Dim objGR As New GetRecordset
    Dim objRS As DAO.Recordset
    
    With objGR
        Set objRS = .runCommand("tblUpdateData")
        propCode = .code
    End With
    
    If code = 1 Then
        Exit Sub
    End If
    
    On Error GoTo errExit
    
    Dim lngLineNumber As Long
    Dim strLine As String
    Dim varRecordset As Variant
    Dim lngID As Long
    
    If objRS.RecordCount = 0 Then
        lngID = 0
    Else
        lngID = DMax("ID", "tblUpdateData")
    End If
    
    With objTS
        Do Until .AtEndOfStream
            strLine = .ReadLine
            lngLineNumber = .Line
            varRecordset = Split(Replace(strLine, """", ""), ",")
            
            lngID = lngID + 1
            
            With objRS
                .AddNew
                
                .Fields("ID").Value = lngID
                
                If varRecordset(2) <> "" Then
                    .Fields("郵便番号").Value = varRecordset(2)
                End If
                
                If varRecordset(3) <> "" Then
                    .Fields("都道府県名-カナ").Value = varRecordset(3)
                End If
                
                If varRecordset(4) <> "" Then
                    .Fields("市区町村名-カナ").Value = varRecordset(4)
                End If
                
                If varRecordset(5) <> "" Then
                    .Fields("町域名-カナ").Value = varRecordset(5)
                End If
                
                If varRecordset(6) <> "" Then
                    .Fields("都道府県名").Value = varRecordset(6)
                End If
                
                If varRecordset(7) <> "" Then
                    .Fields("市区町村名").Value = varRecordset(7)
                End If
                
                If varRecordset(8) <> "" Then
                    .Fields("町域名").Value = varRecordset(8)
                End If
                
                If varRecordset(13) <> "" Then
                    .Fields("更新の表示").Value = varRecordset(13)
                End If
                
                If varRecordset(14) <> "" Then
                    .Fields("変更理由").Value = varRecordset(14)
                End If
                
                .Update
            End With
        Loop
    End With
    
    objTS.Close
    objRS.Close
    Set objRS = Nothing
    
    propCode = 2
    Exit Sub
errExit:
    With objSC
        .message = "エラーが発生しました。 " & Err.Description
        .errorTerminate
        Debug.Print lngLineNumber & " 行目: " & strLine
        propCode = .code
    End With
End Sub
