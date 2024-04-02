VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm11 
   Caption         =   "Simergy Selection"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   9075.001
   ClientWidth     =   18615
   OleObjectBlob   =   "UserForm11.frx":0000
End
Attribute VB_Name = "UserForm11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton6_Click()
    With Worksheets(1)
        ATUALIZA (True)
    End With
    UserForm11.Hide
    simergyPoints
End Sub
Private Sub CommandButton4_Click()
    With Worksheets(1)
        ATUALIZA (False)
    End With
    UserForm11.Hide
End Sub
Private Sub CommandButton5_Click()
    With Worksheets(1)
        ATUALIZA (True)
    End With
    simergyPoints
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Label19_Click()

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub TextBox340_Change()

End Sub

Private Sub UserForm_Click()

End Sub
