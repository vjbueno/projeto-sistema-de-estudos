VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Usf_Login 
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Usf_Login.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Usf_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If TextBox1 = "1234" Then
    MsgBox "Usário logado com sucesso!"
    Application.Visible = True
    Usf_Login.Hide
    Exit Sub
    End If
    
    MsgBox "Senha incorreta"
    
   
    
End Sub

Private Sub TextBox1_Change()

End Sub
