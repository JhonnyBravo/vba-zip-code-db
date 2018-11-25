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

Implements ICommand

Private objSC As IStatus
Private objGD As GetDialog
Private objOD As OpenDialog

'Properties

'@see ICommand#code
Public Property Get ICommand_code() As Long
    ICommand_code = objSC.code
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
        .init "file"
        .ICommand_run
        
        objSC.code = .ICommand_code
        
        If objSC.code = 1 Then
            getCsvPath = strPath
            Exit Function
        End If
        
        Set objDialog = .dialog
    End With
    
    With objOD
        .init objDialog
        .clearFilters
        .addFilters "CSV", "*.csv"
        .ICommand_run
        
        objSC.code = .ICommand_code
        strPath = .path
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
        .init "file"
        .ICommand_run
        
        objSC.code = .ICommand_code
        
        If objSC.code = 1 Then
            getExcelPath = strPath
            Exit Function
        End If
        
        Set objDialog = .dialog
    End With
    
    With objOD
        .init objDialog
        .clearFilters
        .addFilters "Excel", "*.xls;*.xlsx;*.xlsm"
        .ICommand_run
        
        objSC.code = .ICommand_code
        strPath = .path
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
        .init "directory"
        .ICommand_run
        
        objSC.code = .ICommand_code
        
        If objSC.code = 1 Then
            getDirectoryPath = strPath
            Exit Function
        End If
        
        Set objDialog = .dialog
    End With
    
    With objOD
        .init objDialog
        .ICommand_run
        
        objSC.code = .ICommand_code
        strPath = .path
    End With
    
    getDirectoryPath = strPath
End Function

Private Sub ICommand_run()

End Sub
