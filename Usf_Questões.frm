VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Usf_Questões 
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12480
   OleObjectBlob   =   "Usf_Questões.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Usf_Questões"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_save_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub btn_save_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_voltar.Visible = True
    btn_save.BackColor = &H4000&
End Sub

Private Sub btn_save_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Usf_Questões.Hide
    Usf_Principal.Show
    
End Sub

Private Sub Cb_Disc_Change()
    'declrando variavel i do tipo inteiro
    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    Cb_Sub.Clear
    
    
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

Private Sub Im_Computar_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub


Private Sub Im_Computar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_Computar.Visible = True
    Im_Computar.BackColor = &H4000&
End Sub


Private Sub Im_Computar_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Application.ScreenUpdating = False
    'configurando validação do combo box de disciplina e subdisciplina para não cadastrar caso esteja vazio
    If Cb_Disc = "" Or Cb_Sub = "" Then
        MsgBox "Favor preencher todos os campos!", vbCritical
        Application.ScreenUpdating = False
    Exit Sub
    End If
    
    'configurando validação do combo box de data e questões feitas para não cadastrar caso esteja vazio
    If Txt_Data = "" Or Txt_QF = "" Then
        MsgBox "Favor preencher todos os campos!", vbCritical
        Application.ScreenUpdating = True
    Exit Sub
    End If
    
    'configurando validação do combo box de data e questões feitas para não cadastrar caso esteja vazio
    If Txt_QA = "" Then
        MsgBox "Favor preencher todos os campos!", vbCritical
        Application.ScreenUpdating = True
    Exit Sub
    End If

    'se no combo box de questões acertadas for maior que o combo box de questões feitas
    If Val(Txt_QA) > Val(Txt_QF) Then
        'exiba essa mensagem ao usuario:
        MsgBox "O número de questões acertadas não pode ser maior que o número de questões feitas", vbCritical
        
        'deixar os campos de qa e qf vazios se o numero de qa for maior que qf e fazer com que o cursos inicie em qa
        Txt_QA = ""
        Txt_QF = ""
        Txt_QF.SetFocus
        
    Application.ScreenUpdating = True
    Exit Sub
    End If
    

    
    'declarando variavel do tipo int
    Dim i As Integer
    
     i = 2
     
    'faça enquanto na linha 2 celula 4 for diferente de vazio = ""
    Do While Sheets("BD").Cells(i, 4) <> ""
    i = i + 1
    Loop
    'cadastrando dados na aba bd nas linhas e colunas
    Sheets("BD").Cells(i, 4) = Cb_Disc
    Sheets("BD").Cells(i, 5) = Cb_Sub
    Sheets("BD").Cells(i, 6) = Format(Txt_Data.Value, "mm/dd/yyyy")
    Sheets("BD").Cells(i, 7) = Txt_QF.Value
    Sheets("BD").Cells(i, 8) = Txt_QA.Value
    
    'mensagem a ser exibida ao usuário após o cadastro ter sido realizado:
    MsgBox "Questões registradas com sucesso!", vbInformation
    
    Cb_Disc = ""
    Cb_Sub = ""
    Txt_Data = ""
    Txt_QF = ""
    Txt_QA = ""
    
    'função para que o cursor da tela inicie na combo box disciplina
    Cb_Disc.SetFocus
    
    Application.ScreenUpdating = True
    
End Sub


Private Sub Label3_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Lb_fundo_preto_Change()
    
End Sub

Private Sub Lb_fundo_preto_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_voltar.Visible = False
    Lb_Computar.Visible = False
    btn_save.BackColor = &H808080
    Im_Computar.BackColor = &H808080
End Sub


Private Sub TextBox3_Change()

End Sub

Private Sub Txt_Data_Change()
    If Len(Txt_Data) = 2 Then
    Txt_Data = Txt_Data & "/"
    ElseIf Len(Txt_Data) = 5 Then
    Txt_Data = Txt_Data & "/"
    End If
    
End Sub

Private Sub UserForm_Activate()
    
    Application.ScreenUpdating = False
    
    Lb_voltar.Visible = False
    Lb_Computar.Visible = False
    btn_save.BackColor = &H808080
    Im_Computar.BackColor = &H808080
    
    Cb_Disc.Clear

    'Declarando variavel i
    Dim i As Integer
    
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

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_voltar.Visible = False
    Lb_Computar.Visible = False
    btn_save.BackColor = &H808080
    Im_Computar.BackColor = &H808080
End Sub


Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        MsgBox "Saia pelo botão de sair da tela principal", vbCritical, "ATENÇÂO"
        Cancel = True
    End If
End Sub


