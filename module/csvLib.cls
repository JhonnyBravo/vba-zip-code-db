VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csvLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private lngCode As Long
Private strMessage As String

Private lngFileNumber As Long

Private strZipCode As String
Private strPhonetic As String
Private strAddress As String
Private lngUpdateCode As Long
Private lngReasonCode As Long

Public Property Get code() As Long
    code = lngCode
End Property

Public Property Get message() As String
    message = strMessage
End Property

Public Function getRecordCollection(strPath As String) As Collection
    Dim recordCollection As New Collection
    Dim strLine As String
    Dim varRecord As Variant
    Dim lngRecordCount As Long
    
    On Error GoTo errExit
    
    lngFileNumber = FreeFile
    Open strPath For Input As #lngFileNumber
    
    Do Until EOF(lngFileNumber)
        Line Input #lngFileNumber, strLine
        varRecord = Split(Replace(strLine, """", ""), ",")
        
        strZipCode = Format(varRecord(2), "0000000")
        strPhonetic = varRecord(3) & varRecord(4) & varRecord(5)
        strAddress = varRecord(6) & varRecord(7) & varRecord(8)
        lngUpdateCode = CLng(varRecord(13))
        lngReasonCode = CLng(varRecord(14))
        
        recordCollection.Add Array(strZipCode, strPhonetic, strAddress, lngUpdateCode, lngReasonCode)
        lngRecordCount = lngRecordCount + 1
    Loop
    
    Set getRecordCollection = recordCollection
    Close lngFileNumber
    
    lngCode = 0
    strMessage = lngRecordCount & " 件のデータを取得しました。"
    Exit Function
    
errExit:
    lngCode = Err.Number
    strMessage = Err.Description
End Function
