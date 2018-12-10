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

Implements IStream

Private objStatus As IStatus
Private objStream As TextStream

Private Sub Class_Initialize()
    Set objStatus = New Status
End Sub

Private Sub Class_Terminate()
    IStream_closeStream
End Sub

'@see IStream#openStream
Public Function IStream_openStream(objFile As Scripting.IFile, strMode As String) As IStream
    objStatus.initStatus
    Debug.Print objFile.path & " を開いています......"
    
    If strMode = "read" Then
        Set objStream = objFile.OpenAsTextStream(ForReading)
        objStatus.setCode 2
    ElseIf strMode = "write" Then
        Set objStream = objFile.OpenAsTextStream(ForWriting)
        objStatus.setCode 2
    ElseIf strMode = "append" Then
        Set objStream = objFile.OpenAsTextStream(ForAppending)
        objStatus.setCode 2
    Else
        objStatus.errorTerminate "strMode の値が不正です。"
    End If
    
    Set IStream_openStream = Me
End Function

'@see IStream#closeStream
Public Function IStream_closeStream() As IStream
    objStatus.initStatus
    
    If objStream Is Nothing = False Then
        objStream.Close
        Set objStream = Nothing
        objStatus.setCode 2
    End If
    
    Set IStream_closeStream = Me
End Function

'@see IStream#getStream
Public Function IStream_getStream() As Scripting.ITextStream
    Set IStream_getStream = objStream
End Function

'@see IStream#Status
Public Property Get IStream_Status() As IStatus
    Set IStream_Status = objStatus
End Property
