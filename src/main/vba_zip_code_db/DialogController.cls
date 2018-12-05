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

Implements IController

Private objSC As IStatus
Private objGD As GetDialog
Private objOD As OpenDialog

'Properties

Public Property Get IController_code() As Long
    IController_code = objSC.code
End Property

'Methods

Private Sub Class_Initialize()
    Set objSC = New Status
    Set objGD = New GetDialog
    Set objOD = New OpenDialog
End Sub

'ファイルピッカーダイアログを起動し、選択した CSV ファイルのパスを取得して返す。
'
'@return path 選択したファイルのパスを返す。
Public Function getCsvPath() As String
    Dim strPath As String
    Dim objDialog As FileDialog
    
    With objGD
        .IReturnObject_setParamValue "dialogType", "file"
        Set objDialog = .IReturnObject_run
        
        objSC.code = .IReturnObject_code
        
        If objSC.code = 1 Then
            getCsvPath = strPath
            Exit Function
        End If
    End With
    
    With objOD
        .IReturnVariant_setParamObject "dialog", objDialog
        .clearFilters
        .addFilters "CSV", "*.csv"

        strPath = .IReturnVariant_run
        objSC.code = .IReturnVariant_code
    End With
    
    getCsvPath = strPath
End Function

'ファイルピッカーダイアログを起動し、選択した Excel ファイルのパスを取得して返す。
'
'@return path 選択したファイルのパスを返す。
Public Function getExcelPath() As String
    Dim strPath As String
    Dim objDialog As FileDialog
    
    With objGD
        .IReturnObject_setParamValue "dialogType", "file"
        Set objDialog = .IReturnObject_run
        
        objSC.code = .IReturnObject_code
        
        If objSC.code = 1 Then
            getExcelPath = strPath
            Exit Function
        End If
    End With
    
    With objOD
        .IReturnVariant_setParamObject "dialog", objDialog
        .clearFilters
        .addFilters "Excel", "*.xls;*.xlsx;*.xlsm"
        
        strPath = .IReturnVariant_run
        objSC.code = .IReturnVariant_code
    End With
    
    getExcelPath = strPath
End Function

'フォルダーピッカーダイアログを起動し、選択したディレクトリのパスを取得して返す。
'
'@return path 選択したディレクトリのパスを返す。
Public Function getDirectoryPath() As String
    Dim strPath As String
    Dim objDialog As FileDialog
    
    With objGD
        .IReturnObject_setParamValue "dialogType", "directory"
        Set objDialog = .IReturnObject_run
        
        objSC.code = .IReturnObject_code
        
        If objSC.code = 1 Then
            getDirectoryPath = strPath
            Exit Function
        End If
    End With
    
    With objOD
        .IReturnVariant_setParamObject "dialog", objDialog
        strPath = .IReturnVariant_run
        objSC.code = .IReturnVariant_code
    End With
    
    getDirectoryPath = strPath
End Function
