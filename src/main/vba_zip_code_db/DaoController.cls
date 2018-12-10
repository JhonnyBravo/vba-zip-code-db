VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaoController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IDao

Private objStatus As IStatus

Private objConnection As DAO.Database
Private objQD As DAO.QueryDef
Private objRS As DAO.Recordset

Private Sub Class_Initialize()
    Set objStatus = New Status
End Sub

Private Sub Class_Terminate()
    IDao_closeQueryDef
    IDao_closeRecordset
    IDao_closeConnection
End Sub

'@see IDao#openConnection
Public Function IDao_openConnection() As IDao
    objStatus.initStatus
    Set objConnection = Application.CurrentDb
    objStatus.setCode 2
    Set IDao_openConnection = Me
End Function

'@see IDao#closeConnection
Public Function IDao_closeConnection() As IDao
    objStatus.initStatus
    
    If objConnection Is Nothing = False Then
        objConnection.Close
        Set objConnection = Nothing
        objStatus.setCode 2
    End If
    
    Set IDao_closeConnection = Me
End Function

'@see IDao#getConnection
Public Function IDao_getConnection() As DAO.Database
    Set IDao_getConnection = objConnection
End Function

'@see IDao#openQueryDef
Public Function IDao_openQueryDef(strEntity As String) As IDao
    objStatus.initStatus
    
    If objConnection Is Nothing Then
        objStatus.errorTerminate "データベース接続が確立されていません。"
        GoTo finally
    End If
    
    On Error GoTo catch
    
    Debug.Print strEntity & " を開いています......"
    Set objQD = objConnection.QueryDefs(strEntity)
    objStatus.setCode 2
    GoTo finally
catch:
    objStatus.errorTerminate "エラーが発生しました。 " & Err.description
finally:
    Set IDao_openQueryDef = Me
End Function

'@see IDao#closeQueryDef
Public Function IDao_closeQueryDef() As IDao
    objStatus.initStatus
    
    If objQD Is Nothing = False Then
        objQD.Close
        Set objQD = Nothing
        objStatus.setCode 2
    End If
    
    Set IDao_closeQueryDef = Me
End Function

'@see IDao#getQueryDef
Public Function IDao_getQueryDef() As DAO.QueryDef
    Set IDao_getQueryDef = objQD
End Function

'@see IDao#openRecordset
Public Function IDao_openRecordset(strEntity As String) As IDao
    objStatus.initStatus
    
    If objConnection Is Nothing Then
        objStatus.errorTerminate "データベース接続が確立されていません。"
        GoTo finally
    End If
    
    On Error GoTo catch
    
    Debug.Print strEntity & " を開いています......"
    Set objRS = objConnection.openRecordset(strEntity)
    objStatus.setCode 2
    GoTo finally
catch:
    objStatus.errorTerminate "エラーが発生しました。 " & Err.description
finally:
    Set IDao_openRecordset = Me
End Function

'@see IDao#closeRecordset
Public Function IDao_closeRecordset() As IDao
    objStatus.initStatus
    
    If objRS Is Nothing = False Then
        objRS.Close
        Set objRS = Nothing
        objStatus.setCode 2
    End If
    
    Set IDao_closeRecordset = Me
End Function

'@see IDao#getRecordset
Public Function IDao_getRecordset() As DAO.Recordset
    Set IDao_getRecordset = objRS
End Function

'@see IDao#Status
Public Property Get IDao_Status() As IStatus
    Set IDao_Status = objStatus
End Property
