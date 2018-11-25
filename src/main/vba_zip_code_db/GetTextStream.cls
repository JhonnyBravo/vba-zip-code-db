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

Implements ICommand

Private objSC As IStatus
Private objFSO As FileSystemObject

Private propFile As File
Private propStream As TextStream
Private propMode As String

'Properties

'@see ICommand#code
Public Property Get ICommand_code() As Long
    ICommand_code = objSC.code
End Property

'@return stream 取得した TextStream を返す。
Public Property Get stream() As TextStream
    Set stream = propStream
End Property

'Methods

Private Sub Class_Initialize()
    Set objSC = New Status
    Set objFSO = New FileSystemObject
End Sub

Private Sub Class_Terminate()
    propStream.Close
End Sub

'@param objFile 操作対象とする File を指定する。
'@param mode TextStream のモードを指定する。
'   read: 読取専用で開く。
'   write: 上書きモードで開く。
'   append: 追記モードで開く。
Public Sub init(ByRef objFile As File, mode As String)
    Set propFile = objFile
    propMode = mode
End Sub

'TextStream を取得する。
Public Sub ICommand_run()
    objSC.initStatus
    
    If objFSO.FileExists(propFile.path) = False Then
        With objSC
            .message = propFile.path & " が見つかりません。"
            .errorTerminate
            Exit Sub
        End With
    End If
    
    Debug.Print propFile.path & " を開いています......"
    
    Select Case propMode
        Case "read"
            Set propStream = propFile.OpenAsTextStream(ForReading)
            objSC.code = 2
        Case "write"
            Set propStream = propFile.OpenAsTextStream(ForWriting)
            objSC.code = 2
        Case "append"
            Set propStream = propFile.OpenAsTextStream(ForAppending)
            objSC.code = 2
        Case Else
            Set propStream = Nothing
            
            With objSC
                .message = "mode の値が不正です。"
                .errorTerminate
            End With
    End Select
End Sub
