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

Implements IStatus
Private objSC As New StatusController
Private objED As New EntityDefinition

Private Property Let IStatus_code(lngCode As Long)
    objSC.IStatus_code = lngCode
End Property

Public Property Get IStatus_code() As Long
    IStatus_code = objSC.IStatus_code
End Property

Private Property Let IStatus_message(strMessage As String)
    objSC.IStatus_message = strMessage
End Property

Private Property Get IStatus_message() As String
    IStatus_message = objSC.IStatus_message
End Property

Private Sub Class_Initialize()
    With objED
        .entity = "tblUpdateData"
        .addFieldDefinition 2, "郵便番号"
        .addFieldDefinition 3, "都道府県名-カナ"
        .addFieldDefinition 4, "市区町村名-カナ"
        .addFieldDefinition 5, "町域名-カナ"
        .addFieldDefinition 6, "都道府県名"
        .addFieldDefinition 7, "市区町村名"
        .addFieldDefinition 8, "町域名"
        .addFieldDefinition 13, "更新の表示"
        .addFieldDefinition 14, "変更理由"
    End With
End Sub

'tblUpdateData からレコードを全件削除する。
Public Sub deleteRecord()
    objSC.initStatus
    
    Dim objGQD As New GetQueryDef
    Dim objQD As DAO.QueryDef
    
    With objGQD
        Set objQD = .runCommand("deleteUpdateData")
        IStatus_code = .IStatus_code
    End With
    
    If IStatus_code = 1 Then
        Exit Sub
    End If
    
    On Error GoTo errExit
    
    objQD.Execute
    Exit Sub
errExit:
    IStatus_message = "エラーが発生しました。 " & Err.Description
    objSC.errorTerminate
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
        IStatus_code = .IStatus_code
    End With
    
    If IStatus_code = 1 Then
        Exit Sub
    End If
    
    Dim objGR As New GetRecordset
    Dim objRS As DAO.Recordset
    
    With objGR
        Set objRS = .runCommand(objED.entity)
        IStatus_code = .IStatus_code
    End With
    
    If IStatus_code = 1 Then
        Exit Sub
    End If
    
    On Error GoTo errExit
    
    Dim varRecordset As Variant
    Dim lngID As Long
    
    If objRS.RecordCount = 0 Then
        lngID = 0
    Else
        lngID = DMax("ID", objED.entity)
    End If
    
    Dim objFieldCollection As Collection
    Dim varField As Variant
    
    Set objFieldCollection = objED.getCollection
    
    Dim lngLineNumber As Long
    Dim strLine As String
    
    With objTS
        Do Until .AtEndOfStream
            strLine = .ReadLine
            lngLineNumber = .Line
            varRecordset = Split(Replace(strLine, """", ""), ",")
            
            lngID = lngID + 1
            
            With objRS
                .AddNew
                
                .Fields("ID").Value = lngID
                
                For Each varField In objFieldCollection
                    If varRecordset(varField(0)) <> "" Then
                        .Fields(varField(1)).Value = varRecordset(varField(0))
                    End If
                Next
                
                .Update
            End With
        Loop
    End With
    
    objTS.Close
    objRS.Close
    Set objRS = Nothing
    
    IStatus_code = 2
    Exit Sub
errExit:
    IStatus_message = "エラーが発生しました。 " & Err.Description
    objSC.errorTerminate
    Debug.Print lngLineNumber & " 行目: " & strLine
End Sub
