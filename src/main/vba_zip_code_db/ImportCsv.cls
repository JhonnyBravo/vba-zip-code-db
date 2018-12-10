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

Implements IImportCsv
Implements IIterator

Private objStatus As IStatus

Private propRecordset As DAO.Recordset
Private propStream As TextStream

Private varArray As Variant
Private strLine As String

Private Sub Class_Initialize()
    Set objStatus = New Status
End Sub

'@see IImportCsv#setRecordset
Public Function IImportCsv_setRecordset(objRecordset As DAO.Recordset) As IImportCsv
    objStatus.initStatus
    Set propRecordset = objRecordset
    objStatus.setCode 2
    
    Set IImportCsv_setRecordset = Me
End Function

'@see IImportCsv#setStream
Public Function IImportCsv_setStream(objStream As TextStream) As IImportCsv
    objStatus.initStatus
    Set propStream = objStream
    objStatus.setCode 2
    
    Set IImportCsv_setStream = Me
End Function

'@see IImportCsv#Status
Public Property Get IImportCsv_Status() As IStatus
    Set IImportCsv_Status = objStatus
End Property

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
    strLine = propStream.ReadLine
    varArray = Split(Replace(strLine, """", ""), ",")
    IIterator_nextArray = varArray
End Function

'@see IImportCsv#import
Public Function IImportCsv_import() As IImportCsv
    objStatus.initStatus
    
    If propStream Is Nothing Then
        objStatus.errorTerminate "TextStream を指定してください。"
        GoTo finally
    End If
    
    If propRecordset Is Nothing Then
        objStatus.errorTerminate "Recordset を指定してください。"
        GoTo finally
    End If
    
    Debug.Print "レコードを登録しています......"

    On Error GoTo catch
    
    Dim lngID As Long
    
    While IIterator_hasNext
        IIterator_nextArray
        
        With propRecordset
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

            If varArray(13) <> "" Then
                .Fields("更新の表示").Value = varArray(13)
            End If

            If varArray(14) <> "" Then
                .Fields("変更理由").Value = varArray(14)
            End If
            
            .Update
        End With
    Wend

    objStatus.setCode 2
    GoTo finally
catch:
    objStatus.errorTerminate "エラーが発生しました。 " & Err.description
finally:
    Set IImportCsv_import = Me
End Function
