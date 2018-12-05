VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IReturnObject

Private objStatus As IStatus
Private objConnection As DAO.Database

'Properties

'@see IReturnObject#code
Public Property Get IReturnObject_code() As Long
    IReturnObject_code = objStatus.code
End Property

'Methods

Private Sub Class_Initialize()
    Set objStatus = New Status
End Sub

Private Sub Class_Terminate()
    Set objConnection = Nothing
End Sub

'@return DAO.Dabase を生成して返す。
Public Function IReturnObject_run() As Object
    objStatus.initStatus

    Debug.Print "接続を確立しています......"
    Set objConnection = Application.CurrentDb
    Set IReturnObject_run = objConnection
    
    objStatus.code = 2
End Function

Private Sub IReturnObject_setParamObject(strKey As String, objInstance As Object)

End Sub

Private Sub IReturnObject_setParamValue(strKey As String, varValue As Variant)

End Sub
