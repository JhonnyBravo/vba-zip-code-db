VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'@return zipCode 郵便番号を返す。
Public Property Get zipCode() As String

End Property

'@param strValue 郵便番号として設定する値を指定する。
Public Property Let zipCode(strValue As String)

End Property

'@return prefecturePhonetic 都道府県名-カナを返す。
Public Property Get prefecturePhonetic() As String

End Property

'@param strValue 都道府県名-カナとして設定する値を指定する。
Public Property Let prefecturePhonetic(strValue As String)

End Property

'@return cityPhonetic 市区町村名-カナを返す。
Public Property Get cityPhonetic() As String

End Property

'@param strValue 市区町村名-カナとして設定する値を指定する。
Public Property Let cityPhonetic(strValue As String)

End Property

'@return areaPhonetic 町域名-カナを返す。
Public Property Get areaPhonetic() As String

End Property

'@param strValue 町村名-カナとして設定する値を指定する。
Public Property Let areaPhonetic(strValue As String)

End Property

'@return prefecture 都道府県名を返す。
Public Property Get prefecture() As String

End Property

'@param strValue 都道府県名として設定する値を指定する。
Public Property Let prefecture(strValue As String)

End Property

'@return city 市区町村名を返す。
Public Property Get city() As String

End Property

'@param strValue 市区町村名として設定する値を指定する。
Public Property Let city(strValue As String)

End Property

'@return 町域名を返す。
Public Property Get area() As String

End Property

'@param strValue 町域名として設定する値を指定する。
Public Property Let area(strValue As String)

End Property

'@return updateFlag 更新の表示を返す。
Public Property Get updateFlag() As Long

End Property

'@param lngValue 更新の表示として設定する値を指定する。
Public Property Let updateFlag(lngValue As Long)

End Property

'@return 変更理由を返す。
Public Property Get reasonFlag() As Long

End Property

'@param lngValue 変更理由として設定する値を指定する。
Public Property Let reasonFlag(lngValue As Long)

End Property
