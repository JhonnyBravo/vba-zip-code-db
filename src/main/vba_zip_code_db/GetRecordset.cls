VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IReturnObject

Private objSC As IStatus
Private objParams As IParameters

Private objRS As DAO.recordset

'Properties

'@see IReturnObject#code
Public Property Get IReturnObject_code() As Long
    IReturnObject_code = objSC.code
End Property

'Methods

Private Sub Class_Initialize()
    Set objSC = New Status
    Set objParams = New Parameters
End Sub

Private Sub Class_Terminate()
    Set objRS = Nothing
End Sub

'@return DAO.Recordset を生成して返す。
Public Function IReturnObject_run() As Object
    objSC.initStatus
    
    Dim objDB As DAO.Database
    Dim strEntity As String
    
    With objParams
        Set objDB = objParams.getObject("connection")
        objSC.code = .code
        
        If objSC.code = 1 Then
            Exit Function
        End If
    End With
    
    With objParams
        strEntity = objParams.getVariant("entityName")
        
        objSC.code = .code
        
        If objSC.code = 1 Then
            Exit Function
        End If
    End With
    
    On Error GoTo errExit
    
    Debug.Print strEntity & " を開いています......"
    Set objRS = objDB.OpenRecordset(strEntity)
    Set IReturnObject_run = objRS
    
    objSC.code = 2
    Exit Function
errExit:
    With objSC
        .message = "エラーが発生しました。 " & Err.description
        .errorTerminate
    End With
End Function

'@param strKey 有効なキー: connection
'@param objInstance 操作対象とする DAO.DataBase を指定する。
Public Sub IReturnObject_setParamObject(strKey As String, objInstance As Object)
    objSC.initStatus
    objParams.setObject strKey, objInstance
    objSC.code = 2
End Sub

'@param strKey 有効なキー: entityName
'@param varValue 操作対象とするテーブルオブジェクトまたはクエリオブジェクトの名前を指定する。
Public Sub IReturnObject_setParamValue(strKey As String, varValue As Variant)
    objSC.initStatus
    objParams.setVariant strKey, varValue
    objSC.code = 2
End Sub
