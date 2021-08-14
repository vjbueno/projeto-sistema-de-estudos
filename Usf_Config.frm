VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Usf_Config 
   ClientHeight    =   3420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7290
   OleObjectBlob   =   "Usf_Config.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Usf_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

End Sub

'configurações do botão salvar
Private Sub btn_salvar_Click()
    Sheets("Configurações").Cells(2, 1).Value = Txt_índice.Value / 100
    Sheets("Configurações").Cells(2, 2).Value = Txt_revisão.Value
    
    MsgBox "Configurações salvas com sucesso!", vbInformation
    
End Sub


Private Sub btn_salvar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Lb_Salvar.Visible = True
End Sub


Private Sub btn_save_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_voltar.Visible = True
    btn_save.BackColor = &H4000&
    
End Sub

Private Sub btn_save_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Usf_Config.Hide
    Usf_Principal.Show
    
End Sub


Private Sub Label1_Click()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub TextBox3_Change()

End Sub


Private Sub Label3_Click()

End Sub

Private Sub Lb_fundo_preto_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_voltar.Visible = False
    btn_save.BackColor = &H4000&
    
End Sub

Private Sub Lb_fundo_preto_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_voltar.Visible = False
    btn_save.BackColor = &H80000012
    
End Sub


Private Sub Lb_voltar_Click()

End Sub

'configurações botão de indice (para diminuir valor)
Private Sub SpinButton5_SpinDown()
    If Txt_índice.Value = 0 Then
    
        Exit Sub
    End If
    
    Txt_índice.Value = Txt_índice.Value - 5

End Sub
'configurações botão de indice (para aumentar valor até 100)
Private Sub SpinButton5_SpinUp()
    If Txt_índice.Value = 100 Then
    
        Exit Sub
    End If
    
    Txt_índice.Value = Txt_índice.Value + 5

    End Sub
    
'Configurações do botão de diminuir dias de revisão
Private Sub SpinButton6_SpinDown()

    If Txt_revisão.Value = 1 Then
    
        Exit Sub
    End If
    
    Txt_revisão.Value = Txt_revisão.Value - 1

End Sub

'Configurações do botão de aumentar dias de revisão
Private Sub SpinButton6_SpinUp()
    
    Txt_revisão.Value = Txt_revisão.Value + 1
        
End Sub

Private Sub Txt_revisão_Change()

End Sub

Private Sub UserForm_Activate()

    Application.ScreenUpdating = False
    
    Txt_índice.Value = Sheets("Configurações").Cells(2, 1).Value * 100
    Txt_revisão.Value = Sheets("Configurações").Cells(2, 2).Value
    Lb_voltar.Visible = False
    Lb_Salvar.Visible = False
    btn_save.BackColor = &H808080
    
    
    Application.ScreenUpdating = True

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_voltar.Visible = False
    btn_save.BackColor = &H80000012
    Lb_Salvar.Visible = False
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
        MsgBox "Saia pelo botão de sair da tela principal", vbCritical, "ATENÇÂO"
        Cancel = True
    End If
    
End Sub


