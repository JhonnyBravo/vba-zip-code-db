VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetTextStream"
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

Private objStream As TextStream

'Properties

'@see IReturnObject#code
Public Property Get IReturnObject_code() As Long
    IReturnObject_code = objSC.code
End Property

'Methods

Private Sub Class_Initialize()
    Set objSC = New Status
    Set objParams = New Parameters
    Set objFSO = New FileSystemObject
End Sub

Private Sub Class_Terminate()
    objStream.Close
End Sub

'@return TextStream を取得して返す。
Public Function IReturnObject_run() As Object
    objSC.initStatus

    Dim objFile As File
    Dim strMode As String
    
    With objParams
        Set objFile = .getObject("file")
        objSC.code = .code
        
        If objSC.code = 1 Then
            Exit Function
        End If
        
        strMode = .getVariant("mode")
        objSC.code = .code
        
        If objSC.code = 1 Then
            Exit Function
        End If
    End With
    
    Debug.Print objFile.path & " を開いています......"

    Select Case strMode
        Case "read"
            Set objStream = objFile.OpenAsTextStream(ForReading)
            objSC.code = 2
        Case "write"
            Set objStream = objFile.OpenAsTextStream(ForWriting)
            objSC.code = 2
        Case "append"
            Set objStream = objFile.OpenAsTextStream(ForAppending)
            objSC.code = 2
        Case Else
            Set objStream = Nothing

            With objSC
                .message = "mode の値が不正です。"
                .errorTerminate
            End With
    End Select
    
    Set IReturnObject_run = objStream
End Function

'@param strKey 有効なキー: file
'@param objInstance 操作対象とする File を指定する。
Public Sub IReturnObject_setParamObject(strKey As String, objInstance As Object)
    objSC.initStatus
    objParams.setObject strKey, objInstance
    objSC.code = objParams.code
End Sub

'@param strKey 有効なキー: mode
'@param varValue
'   read: 読取専用で TextStream を取得する。
'   write: 上書きモードで TextStream を取得する。
'   append: 追記モードで TextStream を取得する。
Public Sub IReturnObject_setParamValue(strKey As String, varValue As Variant)
    objSC.initStatus
    objParams.setVariant strKey, varValue
    objSC.code = objParams.code
End Sub
