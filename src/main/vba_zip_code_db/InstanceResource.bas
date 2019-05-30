Attribute VB_Name = "InstanceResource"
Option Compare Database
Option Explicit

'''
'@param path 操作対象とするファイルのパスを指定する。
'@return FileResource
'''
Public Function NewFileResource(path As String) As fileResource
    Dim objFR As New fileResource
    objFR.path = path
    Set NewFileResource = objFR
End Function

'''
'@param path 操作対象とするディレクトリのパスを指定する。
'@return DirectoryResource
'''
Public Function NewDirectoryResource(path As String) As DirectoryResource
    Dim objDR As New DirectoryResource
    objDR.path = path
    Set NewDirectoryResource = objDR
End Function

'''
'@param filterName 拡張子フィルターの名前を指定する。
'   例) Excel
'@param definition 拡張子フィルターの定義を指定する。
'   例) *.xls;*.xlsx;*.xlsm
'@return FilterResource
'''
Public Function NewFilterResource(filterName As String, definition As String) As FilterResource
    Dim objFR As New FilterResource
    
    With objFR
        .filterName = filterName
        .definition = definition
    End With
    
    Set NewFilterResource = objFR
End Function

'''
'@param dialogType 生成するダイアログの種類を指定する。
'   * file: ファイルピッカーダイアログを生成する。
'   * directory: ディレクトリピッカーダイアログを生成する。
'@return DialogResource
'''
Public Function NewDialogResource(dialogType As String) As DialogResource
    Dim objDR As New DialogResource
    objDR.dialogType = dialogType
    Set NewDialogResource = objDR
End Function

'''
'@param path 操作対象とする DB のパスを指定する。
'@return ConnectionResource
'''
Public Function NewConnectionResource(Optional path As String = "") As ConnectionResource
    Dim objCR As New ConnectionResource
    objCR.path = path
    Set NewConnectionResource = objCR
End Function

'''
'@param connection 操作対象とする ConnectionResource を指定する。
'@param entityName 操作対象とするクエリまたはテーブルの名前を指定する。
'@return RecordsetResource
'''
Public Function NewRecordsetResource(ByRef connection As IContext, entityName As String) As recordsetResource
    Dim objRR As New recordsetResource
    
    With objRR
        .connection = connection
        .entityName = entityName
    End With
    
    Set NewRecordsetResource = objRR
End Function

'''
'@param connection 操作対象とする ConnectionResource を指定する。
'@param entityName 操作対象とするクエリの名前を指定する。
'@return QueryResource
'''
Public Function NewQueryResource(ByRef connection As IContext, entityName As String) As QueryResource
    Dim objQR As New QueryResource
    
    With objQR
        .connection = connection
        .entityName = entityName
    End With
    
    Set NewQueryResource = objQR
End Function
