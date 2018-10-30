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

'ファイルピッカーを起動し、選択した CSV ファイルのパスを取得して返す。
'
'@return path 選択したファイルのパスを返す。
Public Function getCsvPath() As String
    Dim objGFP As New GetFilePath
    Dim strPath As String
    
    With objGFP
        .IExtensionFilters_addDefinition "CSV", "*.csv"
        strPath = .runCommand
        IStatus_code = .IStatus_code
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
        .IExtensionFilters_addDefinition "Excel", "*.xls;*.xlsx;*.xlsm"
        strPath = .runCommand
        IStatus_code = .IStatus_code
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
        IStatus_code = .IStatus_code
    End With
    
    getDirectoryPath = strPath
End Function
