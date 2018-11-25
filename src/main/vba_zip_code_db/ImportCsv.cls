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

Implements IRecord
Implements IIterator
Implements ICommand

Private objStatus As IStatus

Private propRecordset As DAO.recordset
Private propStream As TextStream
Private propArray As Variant

'Properties

'@see ICommand#code
Public Property Get ICommand_code() As Long
    ICommand_code = objStatus.code
End Property

'@see IRecord#zipCode
Private Property Let IRecord_zipCode(strValue As String)
    propRecordset.Fields("郵便番号").Value = strValue
End Property

'@see IRecord#zipCode
Private Property Get IRecord_zipCode() As String
    IRecord_zipCode = propArray(2)
End Property

'@see IRecord#prefecturePhonetic
Private Property Let IRecord_prefecturePhonetic(strValue As String)
    propRecordset.Fields("都道府県名-カナ").Value = strValue
End Property

'@see IRecord#prefecturePhonetic
Private Property Get IRecord_prefecturePhonetic() As String
    IRecord_prefecturePhonetic = propArray(3)
End Property

'@see IRecord#cityPhonetic
Private Property Let IRecord_cityPhonetic(strValue As String)
    propRecordset.Fields("市区町村名-カナ").Value = strValue
End Property

'@see IRecord#cityPhonetic
Private Property Get IRecord_cityPhonetic() As String
    IRecord_cityPhonetic = propArray(4)
End Property

'@see IRecord#areaPhonetic
Private Property Let IRecord_areaPhonetic(strValue As String)
    propRecordset.Fields("町域名-カナ").Value = strValue
End Property

'@see IRecord#areaPhonetic
Private Property Get IRecord_areaPhonetic() As String
    IRecord_areaPhonetic = propArray(5)
End Property

'@see IRecord#prefecture
Private Property Let IRecord_prefecture(strValue As String)
    propRecordset.Fields("都道府県名").Value = strValue
End Property

'@see IRecord#prefecture
Private Property Get IRecord_prefecture() As String
    IRecord_prefecture = propArray(6)
End Property

'@see IRecord#city
Private Property Let IRecord_city(strValue As String)
    propRecordset.Fields("市区町村名").Value = strValue
End Property

'@see IRecord#city
Private Property Get IRecord_city() As String
    IRecord_city = propArray(7)
End Property

'@see IRecord#area
Private Property Let IRecord_area(strValue As String)
    propRecordset.Fields("町域名").Value = strValue
End Property

'@see IRecord#area
Private Property Get IRecord_area() As String
    IRecord_area = propArray(8)
End Property

'@see IRecord#updateFlag
Private Property Let IRecord_updateFlag(lngValue As Long)
    propRecordset.Fields("更新の表示").Value = lngValue
End Property

'@see IRecord#updateFlag
Private Property Get IRecord_updateFlag() As Long
    IRecord_updateFlag = propArray(13)
End Property

'@see IRecord#reasonFlag
Private Property Let IRecord_reasonFlag(lngValue As Long)
    propRecordset.Fields("変更理由").Value = lngValue
End Property

'@see IRecord#reasonFlag
Private Property Get IRecord_reasonFlag() As Long
    IRecord_reasonFlag = propArray(14)
End Property

'Methods

Private Sub Class_Initialize()
    Set objStatus = New Status
End Sub

'@param objRecordset 操作対象とする Recordset を指定する。
'@param objStream 操作対象とする TextStream を指定する。
Public Sub init(ByRef objRecordset As DAO.recordset, ByRef objStream As TextStream)
    Set propRecordset = objRecordset
    Set propStream = objStream
End Sub

'CSV からレコードを取得し、 Access のテーブルオブジェクトへ登録する。
Public Sub ICommand_run()
    objStatus.initStatus
    Debug.Print "レコードを登録しています......"
    
    On Error GoTo errExit
    
    While IIterator_hasNext
        IIterator_nextArray
        propRecordset.AddNew
        
        IRecord_zipCode = IRecord_zipCode
        
        IRecord_prefecturePhonetic = IRecord_prefecturePhonetic
        IRecord_cityPhonetic = IRecord_cityPhonetic
        IRecord_areaPhonetic = IRecord_areaPhonetic
        
        IRecord_prefecture = IRecord_prefecture
        IRecord_city = IRecord_city
        IRecord_area = IRecord_area
        
        IRecord_updateFlag = IRecord_updateFlag
        IRecord_reasonFlag = IRecord_reasonFlag
        
        propRecordset.Update
    Wend
    
    objStatus.code = 2
    Exit Sub
errExit:
    With objStatus
        .message = "エラーが発生しました。 " & Err.description
        .errorTerminate
    End With
End Sub

'@return Boolean CSV に次の行が存在するかを確認する。
'   True: 次の行が存在することを表す。
'   False: 次の行が存在しないことを表す。
Private Function IIterator_hasNext() As Boolean
    If propStream.AtEndOfStream = False Then
        IIterator_hasNext = True
    Else
        IIterator_hasNext = False
    End If
End Function

'@return array CSV から一行読込み、配列に変換して返す。
Private Function IIterator_nextArray() As Variant
    Dim strLine As String
    Dim varArray As Variant
    
    strLine = propStream.ReadLine
    propArray = Split(Replace(strLine, """", ""), ",")
    IIterator_nextArray = propArray
End Function
