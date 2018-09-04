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

Private propCode As Long
Private objSC As New StatusController

'@return code メソッド実行直後の終了コードを返す。
'   0: エラーもなく、リソースの変更もなく終了した状態を表す。
'   1: エラー終了した状態を表す。
'   2: リソースを変更し、正常終了した状態を表す。
Public Property Get code() As Long
    code = propCode
End Property

'@param strPath TextStream 取得対象とするファイルのパスを指定する。
'@return TextStream TextStream を取得して返す。
Public Function runCommand(strPath As String) As TextStream
    objSC.initStatus
    
    Dim objFSO As New FileSystemObject
    Dim objTS As TextStream
    
    With objFSO
        If .FileExists(strPath) = False Then
            With objSC
                .message = strPath & " が見つかりません。"
                .errorTerminate
                propCode = .code
            End With
            
            Exit Function
        Else
            Debug.Print strPath & " を開きます。"
            
            Set objTS = .GetFile(strPath).OpenAsTextStream(ForReading)
            
            With objSC
                .code = 2
                .message = ""
                propCode = .code
            End With
        End If
        
        Set runCommand = objTS
    End With
End Function
