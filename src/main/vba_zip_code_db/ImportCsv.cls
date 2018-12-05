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

Implements IReturnObject
Implements IIterator

Private objStatus As IStatus

Private boolRecordset As Boolean
Private boolStream As Boolean
Private objRecordset As DAO.recordset
Private objStream As TextStream

Private varArray As Variant
Private strLine As String

'Properties

'@see IReturnObject#code
Public Property Get IReturnObject_code() As Long
    IReturnObject_code = objStatus.code
End Property

'Methods

Private Sub Class_Initialize()
    Set objStatus = New Status
    boolRecordset = False
    boolStream = False
End Sub

'@return Boolean CSV に次の行が存在するかを確認する。
'   True: 次の行が存在することを表す。
'   False: 次の行が存在しないことを表す。
Private Function IIterator_hasNext() As Boolean
    If objStream.AtEndOfStream = False Then
        IIterator_hasNext = True
    Else
        IIterator_hasNext = False
    End If
End Function

'@return array CSV から一行読込み、配列に変換して返す。
Private Function IIterator_nextArray() As Variant
    strLine = objStream.ReadLine
    varArray = Split(Replace(strLine, """", ""), ",")
    IIterator_nextArray = varArray
End Function

'CSV からレコードを取得し、 Access のテーブルオブジェクトへ登録する。
Public Function IReturnObject_run() As Object
    objStatus.initStatus
    
    If boolRecordset = False Then
        With objStatus
            .message = "recordset を指定してください。"
            .errorTerminate
            Exit Function
        End With
    End If
    
    If boolStream = False Then
        With objStatus
            .message = "stream を指定してください。"
            .errorTerminate
            Exit Function
        End With
    End If
    
    Debug.Print "レコードを登録しています......"

    On Error GoTo errExit
    
    Dim lngID As Long
    
    While IIterator_hasNext
        IIterator_nextArray
        
        With objRecordset
            .AddNew
            
            If .RecordCount = 0 Then
                lngID = 1
            Else
                lngID = DMax("ID", "tblUpdateData") + 1
            End If
            
            .Fields("ID").Value = lngID
            
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

            If varArray(2) <> "" Then
                .Fields("更新の表示").Value = varArray(13)
            End If

            If varArray(2) <> "" Then
                .Fields("変更理由").Value = varArray(14)
            End If
            
            .Update
        End With
    Wend

    objStatus.code = 2
    Exit Function
errExit:
    With objStatus
        .message = "エラーが発生しました。 " & Err.description
        .errorTerminate
    End With
End Function

'@param strKey 有効なキー: recordset, stream
'@param objInstance
'   recordset: 操作対象とする DAO.Recordset を指定する。
'   stream: 操作対象とする TextStream を指定する。
Public Sub IReturnObject_setParamObject(strKey As String, objInstance As Object)
    objStatus.initStatus
    
    If strKey = "recordset" Then
        Set objRecordset = objInstance
        boolRecordset = True
        objStatus.code = 2
    ElseIf strKey = "stream" Then
        Set objStream = objInstance
        boolStream = True
        objStatus.code = 2
    End If
End Sub

Private Sub IReturnObject_setParamValue(strKey As String, varValue As Variant)

End Sub
