VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Usf_CadDis 
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11670
   OleObjectBlob   =   "Usf_CadDis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Usf_CadDis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_save_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub btn_save_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub btn_save_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)

End Sub

Private Sub btn_save_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub btn_save_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_voltar.Visible = True
    btn_save.BackColor = &H4000&
End Sub

Private Sub btn_save_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'configurações do botão voltar
    Usf_CadDis.Hide
    ' levar para tela principal
    Usf_Principal.Show
    
End Sub


Private Sub Cb_Disc_Change()

    Application.ScreenUpdating = False
    
    'declrando variavel i do tipo inteiro
    Dim i As Integer
    
    'variavel i vale 2
    i = 2
    
    
    Do While Sheets("BD").Cells(i, 1) <> ""
        If Cb_Disc = Sheets("BD").Cells(i, 1) Then
            Cb_Sub.AddItem Sheets("BD").Cells(i, 2)
    
    End If
    
    i = i + 1
    Loop
    
    Application.ScreenUpdating = True
End Sub

Private Sub Cb_Sub_Change()

End Sub

Private Sub CommandButton1_Click()
    Application.ScreenUpdating = False
    
    'se a caixa de combinação de disciplina e sub disciplina estiverem vazias:
    If Cb_Disc = "" Or Cb_Sub = "" Then
    'exiba a seguinte mensagem:
    MsgBox "Há campos vazios, favor preenche-los", vbCritical
    Application.ScreenUpdating = True
    Exit Sub
    End If
    

    'Declarando variavel i
    Dim i As Integer
    
    'Atribuindo o valor 2 a variavel i
    i = 2
    
    'faça enquanto a aba BD (linha e coluna 1 for diferente(<>) de vazio
    'função serve para buscar uma linha em branco
    Do While Sheets("BD").Cells(i, 1) <> ""
    i = i + 1
    Loop
    
    Sheets("BD").Cells(i, 1) = Cb_Disc
    Sheets("BD").Cells(i, 2) = Cb_Sub
    
    MsgBox "Dados cadastrados com sucesso", vbInformation
    
    Cb_Disc = " "
    Cb_Sub = " "
    
    Cb_Disc.Clear
    
    
    'declarando variavel contador
    Dim contador As Integer
    
    i = 2
    Do While Sheets("BD").Cells(i, 1) <> ""
    
    contador = Application.WorksheetFunction.CountIf(Range("A2:A" & i), Sheets("BD").Cells(i, 1))
        If contador > 1 Then
        GoTo tocaopau
        End If
        Cb_Disc.AddItem Sheets("BD").Cells(i, 1)
tocaopau:
    i = i + 1
    Loop
    
    Application.ScreenUpdating = True
    
    
End Sub

Private Sub CommandButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_Cadastro.Visible = True
End Sub

Private Sub Lb_Cadastro_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_Cadastro.Visible = True
End Sub

Private Sub Lb_fundo_preto_AfterUpdate()

End Sub

Private Sub Lb_fundo_preto_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub


Private Sub Lb_fundo_preto_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub


Private Sub Lb_fundo_preto_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub


Private Sub Lb_fundo_preto_Change()

End Sub

Private Sub Lb_fundo_preto_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub Lb_fundo_preto_Enter()

End Sub


Private Sub Lb_fundo_preto_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub


Private Sub Lb_fundo_preto_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_voltar.Visible = False
    btn_save.BackColor = &H4000&
End Sub

Private Sub Lb_fundo_preto_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub


Private Sub Lb_voltar_Click()

End Sub

Private Sub UserForm_Activate()
    Application.ScreenUpdating = False
    
    'libar combo box de disciplina
    Cb_Disc.Clear
    
    
    Lb_voltar.Visible = False
    Lb_Cadastro.Visible = False
    btn_save.BackColor = &H808080
    
    'declarando variavel i do tipo inteiro
    Dim i As Integer
    'variavel i vale 2
    i = 2
    
    'declarando variável contador do tipo inteiro
    Dim contador As Integer
    
    'enquanto os dados do banco de dados da linha 2 coluna 1 por diferente de vazio
    Do While Sheets("BD").Cells(i, 1) <> ""
    
    contador = Application.WorksheetFunction.CountIf(Range("A2:A" & i), Sheets("BD").Cells(i, 1))
    If contador > 1 Then
    GoTo tocaopau
    
    End If
    
    Cb_Disc.AddItem Sheets("BD").Cells(i, 1)
tocaopau:
    
    i = i + 1
    Loop
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_voltar.Visible = False
    Lb_Cadastro.Visible = False
    btn_save.BackColor = &H80000012
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        MsgBox "Saia pelo botão de sair da tela principal", vbCritical, "ATENÇÂO"
        Cancel = True
    End If
End Sub
