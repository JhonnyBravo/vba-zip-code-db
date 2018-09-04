VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DialogController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private propCode As Long

'@return code メソッド実行直後の終了コードを返す。
'   0: エラーもなく、リソースの変更もなく終了した状態を表す。
'   1: エラー終了した状態を表す。
'   2: リソースを変更して正常終了した状態を表す。
Public Property Get code() As Long
    code = propCode
End Property

'ファイルピッカーを起動し、選択した CSV ファイルのパスを取得して返す。
'
'@return path 選択したファイルのパスを返す。
Public Function getCsvPath() As String
    Dim objGFP As New GetFilePath
    Dim strPath As String
    
    With objGFP
        .filterCollection.Add Array("CSV", "*.csv")
        strPath = .runCommand
        propCode = .code
    End With
    
    getCsvPath = strPath
End Function

'ファイルピッカーを起動し、選択した Excel ファイルのパスを取得して返す。
'
'@return path 選択したファイルのパスを返す。
Public Function getExcelPath() As String
    Dim objGFP As New GetFilePath
    Dim strPath As String
    
    With objGFP
        .filterCollection.Add Array("Excel", "*.xls;*.xlsx;*.xlsm")
        strPath = .runCommand
        propCode = .code
    End With
    
    getExcelPath = strPath
End Function

'ディレクトリピッカーを起動し、選択したディレクトリのパスを取得して返す。
'
'@return path 選択したディレクトリのパスを返す。
Public Function getDirectoryPath() As String
    Dim objGDP As New getDirectoryPath
    Dim strPath As String
    
    With objGDP
        strPath = .runCommand
        propCode = .code
    End With
    
    getDirectoryPath = strPath
End Function
