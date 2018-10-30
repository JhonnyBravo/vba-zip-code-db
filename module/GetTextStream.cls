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

'@param strPath TextStream 取得対象とするファイルのパスを指定する。
'@return TextStream TextStream を取得して返す。
Public Function runCommand(strPath As String) As TextStream
    objSC.initStatus
    
    Dim objFSO As New FileSystemObject
    Dim objTS As TextStream
    
    With objFSO
        If .FileExists(strPath) = False Then
            IStatus_message = strPath & " が見つかりません。"
            objSC.errorTerminate
            Exit Function
        Else
            Debug.Print strPath & " を開きます。"
            
            Set objTS = .GetFile(strPath).OpenAsTextStream(ForReading)
            IStatus_code = 2
        End If
        
        Set runCommand = objTS
    End With
End Function
