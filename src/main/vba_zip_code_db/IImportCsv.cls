VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IImportCsv"
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

'@param objRecordset 操作対象とする Recordset を指定する。
'
'@return IImportCsv
Public Function setRecordset(ByRef objRecordset As DAO.Recordset) As IImportCsv

End Function

'@param objStream 操作対象とする TextStream を指定する。
'
'@return IImportCsv
Public Function setStream(ByRef objStream As TextStream) As IImportCsv

End Function

'CSV から Access テーブルへデータをインポートする。
'
'@return IImportCsv
Public Function import() As IImportCsv

End Function
