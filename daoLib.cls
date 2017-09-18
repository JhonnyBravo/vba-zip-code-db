VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "daoLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private objDB As DAO.Database
Private objQD As QueryDef

Private lngCode As Long
Private strMessage As String

Public Property Get code() As Long
    code = lngCode
End Property

Public Property Get message() As String
    message = strMessage
End Property

Private Sub Class_Initialize()
    Set objDB = CurrentDb
End Sub

Public Sub invokeQuery(strQueryName As String)
    On Error Resume Next
    
    Set objQD = objDB.QueryDefs(strQueryName)
    objQD.Execute
    
    lngCode = Err.Number
    
    If lngCode = 0 Then
        strMessage = strQueryName & " の実行に成功しました。"
    Else
        strMessage = Err.Description
    End If
    
    Set objQD = Nothing
End Sub

Public Sub insertRecord(strQueryName As String, _
    lngID As Long, strZipCode As String, strPhonetic As String, strAddress As String, _
    intUpdateCode As Integer, intReasonCode As Integer, resistrationDate As Date)
    
    On Error Resume Next
    
    Set objQD = objDB.QueryDefs(strQueryName)
    
    With objQD
        .Parameters("ID").Value = lngID
        .Parameters("zipCode").Value = strZipCode
        .Parameters("phonetic").Value = strPhonetic
        .Parameters("address").Value = strAddress
        .Parameters("updateCode").Value = intUpdateCode
        .Parameters("reasonCode").Value = intReasonCode
        .Parameters("resistrationDate").Value = resistrationDate
    End With
    
    objQD.Execute
    
    lngCode = Err.Number
    
    If lngCode = 0 Then
        strMessage = strQueryName & " の実行に成功しました。"
    Else
        strMessage = Err.Description
    End If
    
    Set objQD = Nothing
End Sub

Private Sub Class_Terminate()
    Set objDB = Nothing
End Sub
