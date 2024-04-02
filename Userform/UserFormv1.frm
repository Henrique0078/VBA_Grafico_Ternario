VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Points Selection"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   9075.001
   ClientWidth     =   18345
   OleObjectBlob   =   "UserFormv1.frx":0000
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click() 'Botão de OK
    With Worksheets(1)
        ActiveWorkbook.Unprotect
        ATUALIZA (True)
        VerificaSimergia
        VerificaSensibilidade
        ATUALIZA (True)
        TamanhoPontos
        UserForm1.Hide
        
    End With
End Sub
Private Sub CommandButton2_Click() 'Botão de Cancela
    With Worksheets(1)
        ActiveWorkbook.Unprotect
        ATUALIZA (False)
        
    End With
    UserForm1.Hide

End Sub
Private Sub CommandButton3_Click() 'Botão de Aplica
    With Worksheets(1)
        ActiveWorkbook.Unprotect
        ATUALIZA (True)
        VerificaSimergia
        VerificaSensibilidade
        ATUALIZA (True)
        TamanhoPontos
    End With
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Label19_Click()

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub UserForm_Click()

End Sub
