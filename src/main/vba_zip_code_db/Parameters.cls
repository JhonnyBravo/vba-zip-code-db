VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IParameters

Private objStatus As IStatus
Private objDict As Dictionary

Private Sub Class_Initialize()
    Set objStatus = New Status
    Set objDict = New Dictionary
End Sub

'@see IParameters#code
Public Property Get IParameters_code() As Long
    IParameters_code = objStatus.code
End Property

'@see IParameters#getObject
Private Function IParameters_getObject(strKey As String) As Object
    objStatus.initStatus
    
    With objDict
        If .Exists(strKey) = False Then
            With objStatus
                .message = strKey & " が見つかりません。"
                .errorTerminate
            End With
            
            Exit Function
        End If
        
        Set IParameters_getObject = objDict.Item(strKey)
    End With
    
    objStatus.code = 2
End Function

'@see IParameters#getVariant
Private Function IParameters_getVariant(strKey As String) As Variant
    objStatus.initStatus
    
    With objDict
        If .Exists(strKey) = False Then
            With objStatus
                .message = strKey & " が見つかりません。"
                .errorTerminate
            End With
            
            Exit Function
        End If
        
        IParameters_getVariant = objDict.Item(strKey)
    End With
    
    objStatus.code = 2
End Function

'@see IParameters#setObject
Public Sub IParameters_setObject(strKey As String, objInstance As Object)
    objStatus.initStatus
    
    With objDict
        If .Exists(strKey) Then
            .Remove strKey
        End If
        
        .Add strKey, objInstance
    End With
    
    objStatus.code = 2
End Sub

'@see IParameters#setVariant
Private Sub IParameters_setVariant(strKey As String, varValue As Variant)
    objStatus.initStatus
    
    With objDict
        If .Exists(strKey) Then
            .Remove strKey
        End If
        
        .Add strKey, varValue
    End With
    
    objStatus.code = 2
End Sub
