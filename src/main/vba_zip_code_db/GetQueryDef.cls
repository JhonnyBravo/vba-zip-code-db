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

Implements IStatus
Private objSC As New StatusController

Private Property Let IStatus_code(lngCode As Long)
    objSC.IStatus_code = lngCode
End Property

Public Property Get IStatus_code() As Long
    IStatus_code = objSC.IStatus_code
End Property

Private Property Let IStatus_message(strMessage As String)
    objSC.IStatus_message = strMessage
End Property

Private Property Get IStatus_message() As String
    IStatus_message = objSC.IStatus_message
End Property

'@param strName QueryDef オブジェクトとして取得するクエリの名前を指定する。
'@return QueryDef を返す。
Public Function runCommand(strName As String) As DAO.QueryDef
    objSC.initStatus
    
    Dim objDB As DAO.Database
    Dim objQD As DAO.QueryDef
    
    Set objDB = Application.CurrentDb
    
    On Error GoTo errExit
    
    Debug.Print strName & " を実行します。"
    
    Set objQD = objDB.QueryDefs(strName)
    IStatus_code = 2
    
    Set runCommand = objQD
    Exit Function
errExit:
    IStatus_message = "エラーが発生しました。 " & Err.Description
    objSC.errorTerminate
End Function
