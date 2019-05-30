Attribute VB_Name = "TestDialogResource"
Option Compare Database
Option Explicit

Public Sub testDialog1()
    Dim objDR As DialogResource
    Dim strPath As String
    
    'ファイルピッカーを起動できること。
    Set objDR = NewDialogResource("file")
    
    With objDR
        '拡張子フィルターの設定ができること。
        .addFilter NewFilterResource("CSV", "*.csv")
        .addFilter NewFilterResource("Excel", "*.xls;*.xlsx;*.xlsm")
        
        'ファイルピッカーで選択したファイルのパスを取得できること。
        strPath = .getPath
        Debug.Print "code = " & .status.code
        Debug.Print "path = " & strPath
        
        Select Case .status.code
            '終了コードが 2 である場合にパスが空ではないこと。
            Case 2:
                If strPath <> "" Then
                    Debug.Print "OK"
                Else
                    Debug.Print "NG"
                End If
            '終了コードが 2 ではない場合にパスが空であること。
            Case Default:
                If strPath = "" Then
                    Debug.Print "OK"
                Else
                    Debug.Print "NG"
                End If
        End Select
    End With
End Sub

Public Sub testDialog2()
    Dim objDR As DialogResource
    Dim strPath As String
    
    'ディレクトリピッカーを起動できること。
    Set objDR = NewDialogResource("directory")
    
    With objDR
        'ディレクトリピッカーで選択したディレクトリのパスを取得できること。
        strPath = .getPath
        Debug.Print "code = " & .status.code
        Debug.Print "path = " & strPath
        
        Select Case .status.code
            '終了コードが 2 である場合にパスが空ではないこと。
            Case 2:
                If strPath <> "" Then
                    Debug.Print "OK"
                Else
                    Debug.Print "NG"
                End If
            '終了コードが 2 ではない場合にパスが空であること。
            Case Default:
                If strPath = "" Then
                    Debug.Print "OK"
                Else
                    Debug.Print "NG"
                End If
        End Select
    End With
End Sub

Public Sub testDialog3()
    Dim objDR As DialogResource
    '不正な dialogType を指定した場合にエラーとなること。
    Set objDR = NewDialogResource("test")
    
    Dim strPath As String
    strPath = objDR.getPath
    
    
    With objDR
        '終了コードが 1 であること。
        Debug.Print "code = " & .status.code
        
        If .status.code = 1 Then
            Debug.Print "OK"
        Else
            Debug.Print "NG"
        End If
    End With
    
    'パスが空であること。
    Debug.Print "path = " & strPath
    
    If strPath = "" Then
        Debug.Print "OK"
    Else
        Debug.Print "NG"
    End If
End Sub
