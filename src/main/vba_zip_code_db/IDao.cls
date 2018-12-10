VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'IStatus の機能を継承するためのプロパティ。
'
'@return IStatus
Public Property Get Status() As IStatus

End Property

'DB 接続を確立する。
'
'@return IDao
Public Function openConnection() As IDao

End Function

'DB 接続を切断する。
'
'@return IDao
Public Function closeConnection() As IDao

End Function

'@return DAO.Database
Public Function getConnection() As DAO.Database

End Function

'QueryDef を生成する。
'
'@param strEntity 操作対象とするクエリオブジェクトの名前または SQL を直接記述する。
'@return IDao
Public Function openQueryDef(strEntity As String) As IDao

End Function

'QueryDef を閉じる。
'
'@return IDao
Public Function closeQueryDef() As IDao

End Function

'@return DAO.QueryDef
Public Function getQueryDef() As DAO.QueryDef

End Function

'Recordset を開く。
'
'@param strEntity 操作対象とするテーブルまたはクエリの名前を指定する。
'
'@return IDao
Public Function openRecordset(strEntity As String) As IDao

End Function

'Recordset を閉じる。
'
'@return IDao
Public Function closeRecordset() As IDao

End Function

'@return DAO.Recordset
Public Function getRecordset() As DAO.Recordset

End Function
