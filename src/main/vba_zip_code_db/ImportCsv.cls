VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImportCsv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private propRR As recordsetResource
Private propFR As fileResource
Private propStatus As StatusResource

'''
'@return RecordsetResource
'''
Public Property Get recordsetResource() As recordsetResource
    Set recordsetResource = propRR
End Property

'''
'@param resource 操作対象とする RecordsetResource を指定する。
'''
Public Property Let recordsetResource(ByRef resource As recordsetResource)
    Set propRR = resource
End Property

'''
'@return FileResource
'''
Public Property Get fileResource() As fileResource
    Set fileResource = propFR
End Property

'''
'@param resource 操作対象とする FileResource を指定する。
'''
Public Property Let fileResource(ByRef resource As fileResource)
    Set propFR = resource
End Property

'''
'@return StatusResource
'''
Public Property Get status() As StatusResource
    Set status = propStatus
End Property

Private Sub Class_Initialize()
    Set propStatus = New StatusResource
End Sub

'''
'CSV をテーブルへインポートする。
'''
Public Sub import()
    Dim objRS As DAO.Recordset
    Dim objTS As TextStream
    
    Set objRS = recordsetResource.context.openContext
    status.code = recordsetResource.status.code
    
    If status.code = 1 Then
        Exit Sub
    End If
    
    Set objTS = fileResource.context.openContext
    status.code = fileResource.status.code
    
    If status.code = 1 Then
        Exit Sub
    End If
    
    Dim strLine As String
    Dim varArray As Variant
    Dim lngId As Long
    
    On Error GoTo catch
    
    While objTS.AtEndOfStream = False
        strLine = objTS.ReadLine
        varArray = Split(Replace(strLine, """", ""), ",")
        
        With objRS
            .AddNew
            
            If .RecordCount = 0 Then
                lngId = 1
            Else
                lngId = DMax("ID", "tblUpdateData") + 1
            End If
            
            .Fields("ID").Value = lngId
            
            If varArray(2) <> "" Then
                .Fields("郵便番号").Value = varArray(2)
            End If
            
            If varArray(3) <> "" Then
                .Fields("都道府県名-カナ").Value = varArray(3)
            End If
            
            If varArray(4) <> "" Then
                .Fields("市区町村名-カナ").Value = varArray(4)
            End If
            
            If varArray(5) <> "" Then
                .Fields("町域名-カナ").Value = varArray(5)
            End If
            
            If varArray(6) <> "" Then
                .Fields("都道府県名").Value = varArray(6)
            End If
            
            If varArray(7) <> "" Then
                .Fields("市区町村名").Value = varArray(7)
            End If
            
            If varArray(8) <> "" Then
                .Fields("町域名").Value = varArray(8)
            End If
            
            If varArray(13) <> "" Then
                .Fields("更新の表示").Value = varArray(13)
            End If
            
            If varArray(14) <> "" Then
                .Fields("変更理由").Value = varArray(14)
            End If
            
            .Update
        End With
    Wend
    
    status.code = 2
    GoTo finally
    
catch:
    status.printMessage 1, "エラーが発生しました。 " & Err.Description
    Debug.Print objTS.Line & " 行目: " & strLine
finally:
    fileResource.context.closeContext
    recordsetResource.context.closeContext
End Sub
