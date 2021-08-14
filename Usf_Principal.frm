VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Usf_Principal 
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18000
   OleObjectBlob   =   "Usf_Principal.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Usf_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Im_Back_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Im_Back_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_voltar.Visible = True
    Im_Back.BorderColor = &H4000&
End Sub


Private Sub Im_Back_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Application.DisplayAlerts = False
    
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    Application.Quit
    
End Sub


Private Sub Im_CadD_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    
End Sub

Private Sub Im_CadD_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_CadD.Visible = True
    Im_CadD.BorderColor = &H4000&
End Sub

Private Sub Im_CadD_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Usf_Principal.Hide
    Usf_CadDis.Show
End Sub


Private Sub Im_CadQ_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Im_CadQ_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_CadQ.Visible = True
    Im_CadQ.BorderColor = &H4000&
End Sub


Private Sub Im_CadQ_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Usf_Principal.Hide
    Usf_Questões.Show
    
End Sub

Private Sub Im_Config_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Im_Config_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_Config.Visible = True
    Im_Config.BorderColor = &H4000&
End Sub

Private Sub Im_Config_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Usf_Principal.Hide
    Usf_Config.Show
End Sub


Private Sub Im_Histórico_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Im_Histórico_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Lb_Histórico.Visible = True
    Im_Histórico.BorderColor = &H4000&
End Sub

Private Sub Im_Histórico_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Usf_Principal.Hide
    Usf_Histórico.Show
    
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Lb_Cadastro_Click()

End Sub

Private Sub Lb_CadQ_Click()

End Sub

Private Sub Lb_fundo_preto_Change()

End Sub

Private Sub Lb_fundo_preto_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Im_CadQ.BackColor = &H808080
    Im_CadD.BackColor = &H808080
    Im_Histórico.BackColor = &H808080
    Im_Back.BackColor = &H808080
    Im_Config.BackColor = &H808080
    
    Lb_CadQ.Visible = False
    Lb_CadD.Visible = False
    Lb_Histórico.Visible = False
    Lb_voltar.Visible = False
    Lb_Config.Visible = False
End Sub


Private Sub Lbt_Principal_Click()

End Sub

Private Sub UserForm_Activate()

    Application.ScreenUpdating = False
    
    Lb_CadQ.Visible = False
    Lb_CadD.Visible = False
    Lb_Histórico.Visible = False
    Lb_voltar.Visible = False
    Lb_Config.Visible = False



    Lbt_Principal.RowSource = ""
    
    Dim i As Integer
    Dim porcent As Double
    
    i = 2
    
    Do While Sheets("BD").Cells(i, 4) <> ""
    Sheets("BD").Cells(i, 22) = Sheets("BD").Cells(i, 4)
    Sheets("BD").Cells(i, 23) = Sheets("BD").Cells(i, 5)
    Sheets("BD").Cells(i, 21) = Format(Sheets("BD").Cells(i, 6), "mm/dd/yyyy")
    Sheets("BD").Cells(i, 24) = Sheets("BD").Cells(i, 7)
    Sheets("BD").Cells(i, 25) = Sheets("BD").Cells(i, 8)
    
    
    porcent = Sheets("BD").Cells(i, 25).Value / Sheets("BD").Cells(i, 24).Value
     
      If porcent >= Sheets("Configurações").Cells(2, 1).Value Then
        Sheets("BD").Cells(i, 26) = "OK"
        Else
        Sheets("BD").Cells(i, 26) = "Não suficiente"
        End If
        
        Sheets("BD").Cells(i, 27) = Format(Sheets("BD").Cells(i, 21).Value + Sheets("Configurações").Cells(2, 2).Value, "mm/dd/yyyy")
    
    i = i + 1
    Loop
    
    Dim k As Integer
    
    i = 2
    k = 3
    
    Do While Sheets("BD").Cells(i, 22) <> ""
    k = i + 1
    
    Do While Sheets("BD").Cells(k, 22) <> ""
    If Sheets("BD").Cells(i, 22) = Sheets("BD").Cells(k, 22) And _
    Sheets("BD").Cells(i, 23) = Sheets("BD").Cells(k, 23) Then
        If Sheets("BD").Cells(i, 21) > Sheets("BD").Cells(k, 21) Then
        Range("U" & k & ":AA" & k).Select
        Selection.Delete Shift:=xlUp
        Else
        
        Range("U" & i & ":AA" & i).Select
        Selection.Delete Shift:=xlUp
        
        End If
    
    End If
    k = k + 1
    Loop
    'k = k + 1
    i = i + 1
    Loop
   
    
    If Sheets("BD").Cells(2, 21) <> "" Then
    Lbt_Principal.RowSource = "BD_Principal"
    End If
    

    Application.ScreenUpdating = True
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Usf_Login.Show
    
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Im_CadQ.BackColor = &H808080
    Im_CadD.BackColor = &H808080
    Im_Histórico.BackColor = &H808080
    Im_Back.BackColor = &H808080
    Im_Config.BackColor = &H808080
    
    Lb_CadQ.Visible = False
    Lb_CadD.Visible = False
    Lb_Histórico.Visible = False
    Lb_voltar.Visible = False
    Lb_Config.Visible = False
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
End Sub
