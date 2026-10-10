Option Explicit

'====================================================================
' PAINEL - comportamento das caixas de data
'  - aceita só números e "/"  (barras inseridas automaticamente)
'  - fica vermelha se a data estiver inválida
'  - duplo clique = data de hoje
'====================================================================

Public WithEvents Caixa As MSForms.TextBox

Private Sub Caixa_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Dim k As Long

    k = KeyAscii

    If k < 32 Then Exit Sub          'teclas de controle (Backspace, Ctrl+V...)
    If k = 47 Then Exit Sub          '"/"

    If k >= 48 And k <= 57 Then
        With Caixa
            If .SelLength = 0 And .SelStart = Len(.Text) Then
                If Len(.Text) = 2 Or Len(.Text) = 5 Then
                    .Text = .Text & "/"
                    .SelStart = Len(.Text)
                End If
            End If
        End With
        Exit Sub
    End If

    KeyAscii = 0

End Sub

Private Sub Caixa_Change()

    Dim d As Date
    Dim vz As Boolean

    If PO_TextoParaData(Caixa.Text, d, vz) Then
        Caixa.BackColor = vbWhite
    ElseIf Len(Caixa.Text) >= 8 Then
        Caixa.BackColor = PO_COR_ERRO
    Else
        Caixa.BackColor = vbWhite
    End If

End Sub

Private Sub Caixa_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Caixa.Text = Format$(Date, "dd\/mm\/yyyy")
    Cancel = True

End Sub

