VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xlLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private objXL As Excel.Application
Private objWorkbook As Workbook

Private lngCode As Long
Private strMessage As String

Public Property Get code() As Long
    code = lngCode
End Property

Public Property Get message() As String
    message = strMessage
End Property

Private Sub Class_Initialize()
    Set objXL = New Excel.Application
End Sub

Public Sub openWorkbook(strPath As String)
    On Error Resume Next
    
    objXL.Visible = True
    Set objWorkbook = objXL.Workbooks.Open(strPath)
    
    lngCode = Err.Number
    
    If lngCode = 0 Then
        strMessage = strPath & "を開きました。"
    Else
        strMessage = Err.Description
    End If
End Sub

Public Function getRecordCollection() As Collection
    Dim recordCollection As New Collection
    Dim lngRecordCount As Long
    
    Dim strZipCode As String
    Dim strPhonetic As String
    Dim strAddress As String
    Dim lngUpdateCode As Long
    Dim lngReasonCode As Long
    
    On Error GoTo errExit
    
    objWorkbook.Sheets(1).Range("A1").Activate
    
    If objXL.ActiveCell.Value = "" Then
        lngCode = 1
        strMessage = "データが見つかりませんでした。"
    End If
    
    lngRecordCount = 1
    
    While objXL.ActiveCell.Value <> ""
        With objXL.ActiveCell
            strZipCode = Format(.Cells(1, 3).Value, "0000000")
            strPhonetic = .Cells(1, 4).Value & .Cells(1, 5).Value & .Cells(1, 6).Value
            strAddress = .Cells(1, 7).Value & .Cells(1, 8).Value & .Cells(1, 9).Value
            lngUpdateCode = .Cells(1, 14).Value
            lngReasonCode = .Cells(1, 15).Value
        End With
        
        recordCollection.Add Array(strZipCode, strPhonetic, strAddress, lngUpdateCode, lngReasonCode) ', registrationDate)
        lngRecordCount = lngRecordCount + 1
        
        objXL.ActiveCell.Cells(2, 1).Activate
    Wend
    
    lngCode = 0
    strMessage = lngRecordCount & " 件のデータを取得しました。"
    Set getRecordCollection = recordCollection
    
    Exit Function
    
errExit:
    lngCode = Err.Number
    strMessage = Err.Description
End Function

Public Sub closeWorkbook()
    objWorkbook.Close
    objXL.Quit
    
    Set objWorkbook = Nothing
End Sub

Private Sub Class_Terminate()
    Set objXL = Nothing
End Sub
