VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IReturnObject

Private objSC As IStatus
Private objParams As IParameters
Private objFSO As FileSystemObject

Private propDirectory As Folder
Private propFiles As files

'Properties

Public Property Get IReturnObject_code() As Long
    IReturnObject_code = objSC.code
End Property

'Methods

Private Sub Class_Initialize()
    Set objSC = New Status
    Set objParams = New Parameters
    Set objFSO = New FileSystemObject
End Sub

'指定した Folder 配下に存在する Files を取得する。
'
'@return Files
Public Function IReturnObject_run() As Object
    objSC.initStatus
    
    Dim objDirectory As Folder
    Dim objFiles As files
    
    With objParams
        Set objDirectory = .getObject("directory")
        objSC.code = .code
    End With
    
    If objSC.code = 1 Then
        Exit Function
    End If
    
    Debug.Print objDirectory.path & " 配下のファイル一覧を取得しています......"
    Set objFiles = objDirectory.files
    Set IReturnObject_run = objFiles
    
    If objFiles.Count > 0 Then
        objSC.code = 2
    Else
        objSC.initStatus
    End If
End Function

'@param strKey 有効なキー: directory
'@param objInstance 操作対象とする Folder を指定する。
Public Sub IReturnObject_setParamObject(strKey As String, objInstance As Object)
    objSC.initStatus
    objParams.setObject strKey, objInstance
    objSC.code = objParams.code
End Sub

Private Sub IReturnObject_setParamValue(strKey As String, varValue As Variant)

End Sub
