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

Implements ICommand

Private objStatus As IStatus
Private propConnection As DAO.Database

'Properties

'@see ICommand#code
Public Property Get ICommand_code() As Long
    ICommand_code = objStatus.code
End Property

'@return connection 生成された Database を返す。
Public Property Get connection() As DAO.Database
    Set connection = propConnection
End Property

'Methods

Private Sub Class_Initialize()
    Set objStatus = New Status
End Sub

Private Sub Class_Terminate()
    Set propConnection = Nothing
End Sub

'Dabase を生成する。
Public Sub ICommand_run()
    objStatus.initStatus
    
    Debug.Print "接続を確立しています......"
    Set propConnection = Application.CurrentDb
    
    objStatus.code = 2
End Sub
