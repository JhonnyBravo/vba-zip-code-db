VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFindData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnRemoveFilter_Click()
    With Me
        .Filter = ""
        .FilterOn = False
        .txtFindAddress = Null
        .txtFindPhonetic = Null
        .txtFindZipCode = Null
    End With
End Sub

Private Sub Form_Load()
    Me.RecordsetType = 2
End Sub

Private Sub btnOpenUpdateData_Click()
    DoCmd.OpenForm "frmUpdateData"
End Sub

Private Sub txtFindAddress_AfterUpdate()
    With Me
        .Filter = "住所 Like '*" & .txtFindAddress & "*'"
        .FilterOn = True
    End With
End Sub

Private Sub txtFindPhonetic_AfterUpdate()
    With Me
        .Filter = "[住所-カナ] Like '*" & .txtFindPhonetic & "*'"
        .FilterOn = True
    End With
End Sub

Private Sub txtFindZipCode_AfterUpdate()
    With Me
        .Filter = "郵便番号 = '" & Replace(.txtFindZipCode, "-", "") & "'"
        .FilterOn = True
    End With
End Sub
