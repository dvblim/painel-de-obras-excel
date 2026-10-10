Option Explicit
'===============================================================================
' cEntregaPainel - área OBRA ENTREGUE dentro do Painel de Obras
'
' Ocupa a área de dados do painel: mostra a obra, a data de entrega e o
' que vai para ENTREGUES; pede confirmação ali mesmo e move.
' Ao terminar, avisa o painel pelo evento Fechar (moveu = True se a obra
' saiu do follow up).
'===============================================================================

Public Event Fechar(ByVal moveu As Boolean)

' etapas da área
Private Const E_INICIO As Long = 0
Private Const E_CONFIRMA As Long = 1
Private Const E_FEITO As Long = 2
Private Const E_ERRO As Long = 3

Private mFra As MSForms.Frame
Private mArq As cArquivoObras
Private mCliques As Collection
Private mFonte As String
Private mFonteTit As String
Private mCod As String
Private mNome As String
Private mEtapa As Long
Private mMoveu As Boolean

Private lblTitulo As MSForms.Label
Private lblObra As MSForms.Label
Private lblLocal As MSForms.Label
Private lblCaixa As MSForms.Label
Private lblIcone As MSForms.Label
Private lblTit As MSForms.Label
Private lblSub As MSForms.Label
Private lblOkF As MSForms.Label
Private lblOkT As MSForms.Label
Private lblNaoF As MSForms.Label
Private lblNaoT As MSForms.Label
Private lblStatus As MSForms.Label
Private mG(1 To 7) As MSForms.Label
Private mT(1 To 7) As MSForms.Label

'===============================================================================
' MONTAGEM
'===============================================================================
Private Function Rot(ByVal nome As String, ByVal texto As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                     ByVal cor As Long, ByVal tam As Single, ByVal negrito As Boolean, _
                     Optional ByVal fundo As Long = -1, Optional ByVal alinh As Long = 1) As MSForms.Label
    Dim c As MSForms.Label
    Set c = mFra.Controls.Add("Forms.Label.1", nome, True)
    c.Left = l
    c.Top = t
    c.Width = w
    c.Height = h
    c.Font.Name = mFonte
    c.Font.Size = tam
    c.Font.Bold = negrito
    c.ForeColor = cor
    If fundo >= 0 Then
        c.BackStyle = 1
        c.BackColor = fundo
    Else
        c.BackStyle = 0
    End If
    c.TextAlign = alinh
    c.WordWrap = False
    c.Caption = texto
    Set Rot = c
End Function

Private Sub Clique(ByVal rotulo As MSForms.Label, ByVal chave As String)
    Dim ev As cEntregaClique
    Set ev = New cEntregaClique
    ev.Conectar rotulo, chave, Me
    mCliques.Add ev
End Sub

Private Sub Botao(ByVal nome As String, ByVal chave As String, ByVal texto As String, _
                  ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                  ByVal fundo As Long, ByVal cor As Long, ByVal borda As Long, _
                  ByRef rotFundo As MSForms.Label, ByRef rotTexto As MSForms.Label)
    Set rotFundo = Rot(nome & "F", "", l, t, w, h, cor, 8, False, fundo)
    If borda >= 0 Then
        rotFundo.BorderStyle = 1
        rotFundo.BorderColor = borda
    End If
    Set rotTexto = Rot(nome & "T", texto, l, t + (h - 12) / 2 - 0.5, w, 12, cor, 8.5, True, -1, 2)
    Clique rotFundo, chave
    Clique rotTexto, chave
End Sub

Public Sub Criar(ByVal fra As MSForms.Frame)
    Dim c As MSForms.Label
    Dim i As Long
    Dim w As Single
    Dim xd As Single
    Dim wd As Single

    Set mFra = fra
    Set mArq = New cArquivoObras
    Set mCliques = New Collection
    mFonte = FonteIA()
    mFonteTit = FonteTituloIA()
    w = mFra.InsideWidth
    If w < 600 Then w = mFra.Width - 4

    Rot "iaFaixa", "", 0, 0, w, 28, IA_BRANCO, 8, False, IA_NAVY
    Rot "iaFaixaVd", "", 0, 28, w, 3, IA_BRANCO, 8, False, IA_VERDE
    Set lblTitulo = Rot("iaTitulo", "OBRA ENTREGUE", 14, 6, w - 180, 18, IA_BRANCO, 11.5, False)
    lblTitulo.Font.Name = mFonteTit
    Set c = Rot("iaFechar", "Voltar ao painel  " & ChrW(215), w - 150, 8, 136, 14, IA_BRANCO, 9, True, -1, 3)
    Clique c, "voltar"

    Rot "iaIntro", "A linha inteira vai para ENTREGUES (tabela tb_Entregues), com as colunas casadas pelo nome, e a obra sai do follow up.", 16, 40, w - 32, 14, IA_CINZA, 8.5, False

    ' esquerda: obra e situação
    Rot "iaObraRot", "OBRA", 16, 64, 100, 12, IA_CINZA, 7.5, True
    Set lblObra = Rot("iaObra", "", 16, 78, 360, 20, IA_NAVY, 13, True)
    lblObra.Font.Name = mFonteTit
    lblObra.Font.Bold = False
    Set lblLocal = Rot("iaLocal", "", 16, 100, 360, 13, IA_CINZA, 8.5, False)

    Set lblCaixa = Rot("iaCaixa", "", 16, 120, 360, 170, IA_NAVY, 8, False, IA_BRANCO)
    lblCaixa.BorderStyle = 1
    lblCaixa.BorderColor = IA_BORDA
    Set lblIcone = Rot("iaIcone", "", 26, 130, 18, 18, IA_VERDE, 13, False)
    lblIcone.Font.Name = "Wingdings"
    lblIcone.Font.Charset = 2
    Set lblTit = Rot("iaTit", "", 46, 130, 320, 17, IA_NAVY, 11, True)
    Set lblSub = Rot("iaSub", "", 46, 152, 320, 130, IA_CINZA, 8.5, False)
    lblSub.WordWrap = True

    Botao "iaOk", "ok", "Mover para Entregues", 16, 300, 360, 30, IA_NAVY, IA_BRANCO, -1, lblOkF, lblOkT
    Botao "iaNao", "voltar", "Cancelar e voltar ao painel", 16, 338, 360, 26, IA_BRANCO, IA_NAVY, IA_BORDA, lblNaoF, lblNaoT
    Set lblStatus = Rot("iaStatus", "", 16, 372, 360, 13, IA_CINZA, 8, False)

    ' direita: o que vai para o arquivo e o que acontece depois
    xd = 396
    wd = w - xd - 16
    Rot "iaSep", "", xd - 10, 66, 1, 300, IA_NAVY, 6, False, IA_BORDA
    Rot "iaDirTit", "O QUE VAI PARA ENTREGUES", xd, 66, wd, 12, IA_CINZA, 7.5, True
    For i = 1 To 7
        Set mG(i) = Rot("iaG" & i, "", xd, 84 + (i - 1) * 24 + IIf(i >= 4, 16, 0), 16, 16, IA_VERDE, 11, False)
        mG(i).Font.Name = "Wingdings"
        mG(i).Font.Charset = 2
        Set mT(i) = Rot("iaT" & i, "", xd + 20, 86 + (i - 1) * 24 + IIf(i >= 4, 16, 0), wd - 20, 14, IA_NAVY, 9, False)
    Next i
    Rot "iaDepois", "DEPOIS", xd, 160, wd, 12, IA_CINZA, 7.5, True
End Sub

Public Sub Desligar()
    Dim ev As cEntregaClique
    If mCliques Is Nothing Then Exit Sub
    For Each ev In mCliques
        ev.Desconectar
    Next ev
    Set mCliques = Nothing
End Sub

'===============================================================================
' CONTEÚDO
'===============================================================================
' glifo "*" = marcador simples (bolinha); os demais são da fonte Wingdings
Private Sub Linha(ByVal i As Long, ByVal glifo As String, ByVal cor As Long, ByVal texto As String)
    If glifo = "*" Then
        mG(i).Font.Name = mFonte
        mG(i).Font.Charset = 0
        glifo = ChrW(8226)
    Else
        mG(i).Font.Name = "Wingdings"
        mG(i).Font.Charset = 2
    End If
    mG(i).Caption = glifo
    mG(i).ForeColor = cor
    mT(i).Caption = texto
    mT(i).ControlTipText = texto
End Sub

Private Sub Caixa(ByVal glifo As String, ByVal cor As Long, ByVal fundo As Long, ByVal titulo As String, ByVal detalhe As String)
    lblIcone.Caption = glifo
    lblIcone.ForeColor = cor
    lblCaixa.BackColor = fundo
    lblCaixa.BorderColor = cor
    lblTit.Caption = titulo
    lblSub.Caption = detalhe
End Sub

Private Sub Botoes(ByVal textoOk As String, ByVal fundoOk As Long, ByVal textoNao As String)
    lblOkT.Caption = textoOk
    lblOkF.BackColor = fundoOk
    lblOkT.ForeColor = IA_BRANCO
    lblNaoF.Visible = (Len(textoNao) > 0)
    lblNaoT.Visible = (Len(textoNao) > 0)
    lblNaoT.Caption = textoNao
End Sub

' Abre a área para a obra do painel.
Public Sub Preparar(ByVal codObra As String)
    Dim d As Variant
    Dim fora As String
    Dim qtd As Long
    Dim lugar As String
    Dim v As String
    Dim dias As Long

    mCod = Trim$(codObra)
    mMoveu = False
    mEtapa = E_INICIO
    lblStatus.Caption = ""
    mNome = mArq.Texto(mCod, "NOME")
    If Len(mNome) = 0 Then mNome = "CÓDIGO " & mCod
    lblTitulo.Caption = "OBRA ENTREGUE  " & ChrW(183) & "  " & UCase$(mNome)
    lblObra.Caption = mNome
    lugar = "CÓDIGO " & mCod
    v = mArq.Texto(mCod, "CIDADE")
    If Len(v) > 0 Then lugar = lugar & "  " & ChrW(183) & "  " & v
    v = mArq.Texto(mCod, "UF")
    If Len(v) > 0 Then lugar = lugar & " - " & v
    v = mArq.Texto(mCod, "ETAPA")
    If Len(v) > 0 Then lugar = lugar & "  " & ChrW(183) & "  Etapa: " & v
    lblLocal.Caption = lugar

    d = mArq.DataEntrega(mCod)
    If Not mArq.TemArquivo() Then
        Caixa Chr$(251), IA_CORAL, IA_CORAL_FUNDO, "Tabela de entregues não encontrada", "A tabela tb_Entregues (aba ENTREGUES) não está nesta planilha; nada pode ser movido."
        Botoes "Mover para Entregues (bloqueado)", IA_BORDA, "Voltar ao painel"
        mEtapa = E_ERRO
    ElseIf IsEmpty(d) Then
        Caixa Chr$(105), IA_AMARELO, IA_DESTAQUE, "Sem data de entrega", "A coluna ENTREGA desta obra está vazia (ela vem do Cronograma de Obras). Confirme que a obra foi entregue antes de mover."
        Botoes "Mover " & Curto(mNome) & " para Entregues", IA_NAVY, "Cancelar e voltar ao painel"
    ElseIf CDate(d) <= Date Then
        Caixa Chr$(252), IA_VERDE, IA_VERDE_FUNDO, "Entregue em " & Format$(d, "dd\/mm\/yyyy"), "Data de entrega do follow up (vem do Cronograma de Obras). A obra pode ir para o arquivo."
        Botoes "Mover " & Curto(mNome) & " para Entregues", IA_NAVY, "Cancelar e voltar ao painel"
    Else
        dias = CLng(CDate(d) - Date)
        Caixa Chr$(105), IA_AMARELO, IA_DESTAQUE, "Entrega prevista para " & Format$(d, "dd\/mm\/yyyy"), "A data do follow up ainda não chegou (faltam " & dias & " dia(s)). Se a obra já foi entregue, pode mover; a data no arquivo será a que está no follow up."
        Botoes "Mover " & Curto(mNome) & " para Entregues", IA_NAVY, "Cancelar e voltar ao painel"
    End If
    If lblIcone.Caption = Chr$(105) Then
        lblIcone.Font.Name = "Webdings"
    Else
        lblIcone.Font.Name = "Wingdings"
    End If

    qtd = mArq.QtdObservacoes(mCod)
    fora = mArq.ColunasFora()
    Linha 1, Chr$(252), IA_VERDE, mArq.ColunasQueVao() & " colunas casadas pelo nome (fórmulas entram como valor)"
    Linha 2, Chr$(252), IA_VERDE, "HISTÓRICO (" & qtd & " registro(s)) e ÚLTIMO REGISTRO vão junto"
    If Len(fora) > 0 Then
        Linha 3, Chr$(251), IA_CINZA, "Ficam de fora (não existem no arquivo): " & fora
    Else
        Linha 3, Chr$(252), IA_VERDE, "Todas as colunas do follow up existem no arquivo"
    End If
    Linha 4, "*", IA_AZUL, "A obra sai do follow up e as de baixo sobem"
    Linha 5, "*", IA_AZUL, "Fica no histórico de alterações quem moveu e quando"
    Linha 6, "*", IA_AZUL, "Filtros da aba OBRAS são guardados e reaplicados no fim"
    If mArq.TemMarcador(mCod) Then
        Linha 7, "*", IA_AZUL, "O marcador " & ChrW(8220) & "para analisar" & ChrW(8221) & " da obra é concluído"
    Else
        Linha 7, "*", IA_AZUL, ""
        mG(7).Caption = ""
    End If
End Sub

Private Function Curto(ByVal s As String) As String
    If Len(s) > 28 Then s = Left$(s, 26) & "..."
    Curto = s
End Function

'===============================================================================
' CLIQUES
'===============================================================================
Public Sub TratarClique(ByVal chave As String)
    Select Case chave
        Case "voltar"
            RaiseEvent Fechar(mMoveu)
        Case "ok"
            Select Case mEtapa
                Case E_INICIO
                    mEtapa = E_CONFIRMA
                    Caixa Chr$(105), IA_AMARELO, IA_DESTAQUE, "Confirma a movimentação?", _
                        Curto(mNome) & " (CÓDIGO " & mCod & ") vai para ENTREGUES e sai do follow up." & vbLf & vbLf & "Não dá para desfazer com Ctrl+Z."
                    lblIcone.Font.Name = "Webdings"
                    Botoes "Confirmar e mover para Entregues", IA_VERDE, "Cancelar e voltar ao painel"
                Case E_CONFIRMA
                    Mover
                Case E_FEITO, E_ERRO
                    RaiseEvent Fechar(mMoveu)
            End Select
    End Select
End Sub

Private Sub Mover()
    Dim ok As Boolean
    Dim aviso As String

    lblStatus.ForeColor = IA_CINZA
    lblStatus.Caption = "Movendo..."
    DoEvents
    ok = mArq.Entregar(mCod)
    aviso = mArq.AvisoFiltro
    lblStatus.Caption = ""
    lblIcone.Font.Name = "Wingdings"
    If ok Then
        mMoveu = True
        mEtapa = E_FEITO
        Caixa Chr$(252), IA_VERDE, IA_VERDE_FUNDO, "Movida para Entregues", _
            Curto(mNome) & " está em ENTREGUES e saiu do follow up (conferido)." & vbLf & _
            "Registrado no histórico de alterações." & IIf(Len(aviso) > 0, vbLf & vbLf & aviso, "")
        Botoes "Voltar ao painel", IA_NAVY, ""
    Else
        mMoveu = (Len(mArq.UltimoErro) > 0 And InStr(1, mArq.UltimoErro, "não saiu do follow up", vbTextCompare) > 0)
        mEtapa = E_ERRO
        Caixa Chr$(251), IA_CORAL, IA_CORAL_FUNDO, "Não foi possível concluir", mArq.UltimoErro & IIf(Len(aviso) > 0, vbLf & vbLf & aviso, "")
        Botoes "Voltar ao painel", IA_NAVY, ""
    End If
End Sub
