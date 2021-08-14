VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Usf_Histórico 
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13020
   OleObjectBlob   =   "Usf_Histórico.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Usf_Histórico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_save_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    
End Sub

Private Sub btn_save_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_voltar.Visible = True
    btn_save.BackColor = &H4000&
End Sub

Private Sub btn_save_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Usf_Histórico.Hide
    'direcionando o usuário ao clicar no botão voltar para tela principal
    Usf_Principal.Show
    
End Sub

Private Sub Cb_Disc_Change()
    'declrando variavel i do tipo inteiro
    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    Ltb_Histórico.RowSource = ""
    
    Range("M2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    
    Cb_Sub.Clear
    
    
    'variavel i vale 2
    i = 2
    
    Do While Sheets("BD").Cells(i, 1) <> ""
        If Cb_Disc = Sheets("BD").Cells(i, 1) Then
            Cb_Sub.AddItem Sheets("BD").Cells(i, 2)
    
    End If
    
    i = i + 1
    Loop
    
    
    Dim k As Integer
    k = 2
    i = 2
    
    Do While Sheets("BD").Cells(i, 4) <> ""
        If Sheets("BD").Cells(i, 4) = Cb_Disc Then
            Range("D" & i & ":H" & i).Select
            Selection.Copy
            Range("M" & k).Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            k = k + 1
        
        End If
    
    i = i + 1
    Loop
    
    If Sheets("BD").Cells(2, 13) <> "" Then
    Ltb_Histórico.RowSource = "BD_Filtrada"
    End If
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub Cb_Sub_Change()
    Application.ScreenUpdating = False
    'declrando variaveis do tipo inteiro
    Dim i As Integer
    Dim k As Integer
    
    Ltb_Histórico.RowSource = ""
    
    Range("M2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    
    k = 2
    i = 2
    
    Do While Sheets("BD").Cells(i, 4) <> ""
        If Sheets("BD").Cells(i, 4) = Cb_Disc Then
            Range("D" & i & ":H" & i).Select
            Selection.Copy
            Range("M" & k).Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            k = k + 1
        
        End If
    
    i = i + 1
    Loop
    
    If Cb_Sub = "" Then
        GoTo alimenta
    End If
    
    i = 2
    
    Do While Sheets("BD").Cells(i, 13) <> ""
        If Sheets("BD").Cells(i, 14) = Cb_Sub Then
            GoTo proximo_i
        Else
        Range("M" & i & ":Q" & i).Select
        Selection.Delete Shift:=xlUp
        End If
proximo_i:
    
    i = i + 1
    Loop
    
alimenta:
    If Sheets("BD").Cells(2, 13) <> "" Then
        Ltb_Histórico.RowSource = "BD_Filtrada"
    End If



    'declrando variavel i do tipo inteiro
    'Dim i As Integer
    
    'Cb_Sub.Clear
    
    'variavel i vale 2
    
    'i = 2
    'Do While Sheets("BD").Cells(i, 1) <> ""
        'If Cb_Disc = Sheets("BD").Cells(i, 1) Then
            'Cb_Sub.AddItem Sheets("BD").Cells(i, 2)
    
    'End If
    
    'i = i + 1
    'Loop
    Application.ScreenUpdating = True
End Sub


Private Sub Label2_Click()

End Sub


Private Sub Lb_fundo_preto_Change()

End Sub

Private Sub Lb_fundo_preto_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
     Lb_voltar.Visible = False
    btn_save.BackColor = &H808080
End Sub


Private Sub Ltb_Histórico_Click()

End Sub

Private Sub UserForm_Activate()

    Application.ScreenUpdating = False

    
    Lb_voltar.Visible = False
    btn_save.BackColor = &H808080
    Cb_Disc.Clear
    
    
    Range("M2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Ltb_Histórico.RowSource = ""
    
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
    
    Ltb_Histórico.RowSource = "BD_Histórico"
    
    Application.ScreenUpdating = True
    
    
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_voltar.Visible = False
    btn_save.BackColor = &H808080
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        MsgBox "Saia pelo botão de sair da tela principal", vbCritical, "ATENÇÂO"
        Cancel = True
    End If
End Sub


