VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetQueryDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IReturnObject

Private objSC As IStatus
Private objParams As IParameters

Private objQD As DAO.QueryDef

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

'@return DAO.QueryDef を生成して返す。
Public Function IReturnObject_run() As Object
    objSC.initStatus
    
    Dim objDB As DAO.Database
    Dim strSQL As String
    
    Set objDB = objParams.getObject("connection")
    
    With objSC
        .code = objParams.code
        
        If .code = 1 Then
            Exit Function
        End If
    End With
    
    strSQL = objParams.getVariant("sql")
    
    With objSC
        .code = objParams.code
        
        If .code = 1 Then
            Exit Function
        End If
    End With
    
    On Error GoTo errExit

    Debug.Print strSQL & " を開いています......"
    Set objQD = objDB.QueryDefs(strSQL)
    Set IReturnObject_run = objQD
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
    objSC.code = objParams.code
End Sub

'@param strKey 有効なキー: sql
'@param varValue 操作対象とするクエリオブジェクトの名前または SQL を直接記述する。
Public Sub IReturnObject_setParamValue(strKey As String, varValue As Variant)
    objSC.initStatus
    objParams.setVariant strKey, varValue
    objSC.code = objParams.code
End Sub
