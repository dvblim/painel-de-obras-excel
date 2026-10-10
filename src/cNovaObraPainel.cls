Option Explicit
'===============================================================================
' cNovaObraPainel - área NOVA OBRA dentro do Painel de Obras
'
' Ocupa a área de dados do painel enquanto a obra nova é conferida. O CÓDIGO
' é a chave: a área procura a obra no Cronograma de Obras e, se não
' achar, na Carteira - as mesmas bases das fórmulas da
' tb_Obras - e mostra o que as fórmulas vão trazer antes de incluir.
' Ao terminar, avisa o painel pelo evento Fechar.
'===============================================================================

Public Event Fechar(ByVal linhaObras As Long)

Private Const T_Obras As String = "tb_Obras"
Private Const T_CRONO As String = "tb_Cronograma"
Private Const T_AUX As String = "tb_Carteira"

Private Const S_VAZIO As Long = 0
Private Const S_CRONO As Long = 1
Private Const S_AUX As Long = 2
Private Const S_NADA As Long = 3
Private Const S_NA_Obras As Long = 4

Private WithEvents txtCod As MSForms.TextBox

Private mFra As MSForms.Frame
Private mCliques As Collection
Private mFonte As String
Private mFonteTit As String
Private mSituacao As Long
Private mCod As String
Private mLinhaObras As Long

Private lblIcone As MSForms.Label
Private lblTit As MSForms.Label
Private lblSub As MSForms.Label
Private lblCaixa As MSForms.Label
Private lblDesc As MSForms.Label
Private mRot(1 To 8) As MSForms.Label
Private mVal(1 To 8) As MSForms.Label
Private lblPrevTit As MSForms.Label
Private lblOkF As MSForms.Label
Private lblOkT As MSForms.Label
Private lblStatus As MSForms.Label

'===============================================================================
' MONTAGEM (dentro do Frame que o painel cria)
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
    Dim ev As cNovaClique
    Set ev = New cNovaClique
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
    Dim f As MSForms.Label
    Dim t As MSForms.Label
    Dim i As Long
    Dim nomes As Variant
    Dim w As Single
    Dim xd As Single
    Dim wd As Single
    Dim y As Single

    Set mFra = fra
    Set mCliques = New Collection
    mFonte = FonteNL()
    mFonteTit = FonteTituloNL()
    w = mFra.InsideWidth
    If w < 600 Then w = mFra.Width - 4

    Rot "nlFaixa", "", 0, 0, w, 28, NL_BRANCO, 8, False, NL_NAVY
    Rot "nlFaixaAm", "", 0, 28, w, 3, NL_BRANCO, 8, False, NL_AMARELO
    Set c = Rot("nlTitulo", "NOVA OBRA NO FOLLOW UP", 14, 6, 300, 18, NL_BRANCO, 11.5, False)
    c.Font.Name = mFonteTit
    Set c = Rot("nlFechar", "Voltar ao painel  " & ChrW(215), w - 150, 8, 136, 14, NL_BRANCO, 9, True, -1, 3)
    Clique c, "cancelar"

    Set c = Rot("nlIntro", "O CÓDIGO é a chave: o resto da linha vem das fórmulas, que buscam a obra no Cronograma de Obras e, se não acharem, na Carteira. Confira antes de incluir.", 16, 40, w - 32, 14, NL_CINZA, 8.5, False)

    ' coluna da esquerda: CÓDIGO e situação
    Rot "nlCodRot", "CÓDIGO", 16, 66, 100, 12, NL_CINZA, 7.5, True
    Set txtCod = mFra.Controls.Add("Forms.TextBox.1", "nlCod", True)
    txtCod.Left = 16
    txtCod.Top = 80
    txtCod.Width = 160
    txtCod.Height = 28
    txtCod.MaxLength = 12
    txtCod.SpecialEffect = 0
    txtCod.BorderStyle = 1
    txtCod.BorderColor = NL_AMARELO
    txtCod.BackColor = NL_EDICAO
    txtCod.ForeColor = NL_NAVY
    txtCod.Font.Name = mFonte
    txtCod.Font.Size = 15
    txtCod.Font.Bold = True
    Rot "nlCodDica", "a conferência é feita enquanto você digita", 184, 90, 200, 12, NL_CINZA, 8, False

    Set lblCaixa = Rot("nlCaixa", "", 16, 120, 360, 170, NL_NAVY, 8, False, NL_BRANCO)
    lblCaixa.BorderStyle = 1
    lblCaixa.BorderColor = NL_BORDA
    Set lblIcone = Rot("nlIcone", "", 26, 130, 18, 18, NL_VERDE, 13, False)
    lblIcone.Font.Name = "Wingdings"
    lblIcone.Font.Charset = 2
    Set lblTit = Rot("nlTit", "", 46, 130, 320, 17, NL_NAVY, 11, True)
    Set lblSub = Rot("nlSub", "", 46, 152, 320, 130, NL_CINZA, 8.5, False)
    lblSub.WordWrap = True

    Botao "nlOk", "ok", "Incluir no follow up", 16, 300, 360, 30, NL_NAVY, NL_BRANCO, -1, lblOkF, lblOkT
    Botao "nlCancelar", "cancelar", "Cancelar e voltar ao painel", 16, 338, 360, 26, NL_BRANCO, NL_NAVY, NL_BORDA, f, t
    Set lblStatus = Rot("nlStatus", "", 16, 372, 360, 13, NL_CINZA, 8, False)

    ' coluna da direita: o que as fórmulas vão trazer
    xd = 396
    wd = w - xd - 16
    Rot "nlSep", "", xd - 10, 66, 1, 300, NL_NAVY, 6, False, NL_BORDA
    Set lblPrevTit = Rot("nlPrevTit", "O QUE AS FÓRMULAS DO FOLLOW UP VÃO TRAZER", xd, 66, wd, 12, NL_CINZA, 7.5, True)
    Rot "nlDescRot", "NOME", xd, 84, 120, 12, NL_CINZA, 7.5, True
    Set lblDesc = Rot("nlDesc", "", xd + 124, 81, wd - 124, 20, NL_NAVY, 10, True, NL_FUNDO)
    lblDesc.BorderStyle = 1
    lblDesc.BorderColor = NL_BORDA
    nomes = Array("CLIENTE", "CRITICIDADE", "CIDADE - UF", "TIPO DE OBRA", "ÁREA (m²)", "INÍCIO DA OBRA", "FIM DA OBRA", "ENTREGA")
    For i = 1 To 8
        y = 110 + (i - 1) * 28
        Set mRot(i) = Rot("nlPR" & i, CStr(nomes(i - 1)), xd, y + 3, 120, 12, NL_CINZA, 7.5, True)
        Set mVal(i) = Rot("nlPV" & i, "", xd + 124, y, wd - 124, 20, NL_NAVY, 10, False, NL_FUNDO)
        mVal(i).BorderStyle = 1
        mVal(i).BorderColor = NL_BORDA
    Next i
    Rot "nlNota", "Somente o CÓDIGO é gravado no follow up (como texto, igual às bases). Estes dados não são digitados: aparecem pelas fórmulas da tabela.", xd, 340, wd, 26, NL_CINZA, 8, False
    Me.Limpar
End Sub

Public Sub Limpar()
    If txtCod Is Nothing Then Exit Sub
    txtCod.Text = ""
    Conferir
End Sub

Public Sub Focar()
    On Error Resume Next
    txtCod.SetFocus
End Sub

Public Sub Desligar()
    Dim ev As cNovaClique
    If mCliques Is Nothing Then Exit Sub
    For Each ev In mCliques
        ev.Desconectar
    Next ev
    Set mCliques = Nothing
End Sub

'===============================================================================
' CONFERÊNCIA
'===============================================================================
Private Sub txtCod_Change()
    Conferir
End Sub

Private Sub txtCod_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        TratarClique "ok"
    ElseIf KeyCode = 27 Then
        KeyCode = 0
        TratarClique "cancelar"
    End If
End Sub

Private Function Valor(ByVal iCrono As Long, ByVal iAux As Long, ByVal colCrono As String, ByVal colAux As String) As String
    If iCrono > 0 And Len(colCrono) > 0 Then Valor = ValorNL(T_CRONO, iCrono, colCrono)
    If Len(Valor) = 0 And iAux > 0 And Len(colAux) > 0 Then Valor = ValorNL(T_AUX, iAux, colAux)
End Function

Private Sub Conferir()
    Dim iCrono As Long
    Dim iAux As Long
    Dim desc As String
    Dim cid As String
    Dim uf As String
    Dim i As Long

    mCod = Trim$(Replace(txtCod.Text, Chr$(160), ""))
    mLinhaObras = 0
    lblStatus.Caption = ""
    For i = 1 To 8
        mVal(i).Caption = ""
    Next i
    lblDesc.Caption = ""

    If Len(mCod) = 0 Then
        Pintar S_VAZIO, "Digite o CÓDIGO da obra", "Somente o CÓDIGO é digitado; as demais colunas do follow up são preenchidas pelas fórmulas."
        Exit Sub
    End If

    iCrono = LinhaNaTabela(T_CRONO, mCod)
    iAux = LinhaNaTabela(T_AUX, mCod)
    mLinhaObras = LinhaNaTabela(T_Obras, mCod)

    desc = Valor(iCrono, iAux, "NOME", "NOME")
    If iCrono > 0 Or iAux > 0 Then
        mVal(1).Caption = " " & Valor(iCrono, iAux, "CLIENTE", "CLIENTE")
        mVal(2).Caption = " " & Valor(0, iAux, "", "CRITICIDADE")
        cid = Valor(iCrono, iAux, "CIDADE", "CIDADE")
        uf = Valor(iCrono, iAux, "UF", "UF")
        If Len(uf) > 0 Then cid = cid & " - " & uf
        mVal(3).Caption = " " & cid
        mVal(4).Caption = " " & Valor(iCrono, iAux, "TIPO DE OBRA", "TIPO DE OBRA")
        mVal(5).Caption = " " & Valor(0, iAux, "", "ÁREA (m²)")
        mVal(6).Caption = " " & Valor(iCrono, iAux, "INÍCIO DA OBRA", "INÍCIO PREVISTO")
        mVal(7).Caption = " " & Valor(iCrono, iAux, "FIM DA OBRA", "FIM PREVISTO")
        mVal(8).Caption = " " & Valor(iCrono, iAux, "ENTREGA PREVISTA", "ENTREGA PREVISTA")
    End If
    For i = 1 To 8
        If Len(Trim$(mVal(i).Caption)) = 0 Then mVal(i).Caption = " " & ChrW(8212)
    Next i
    If iCrono > 0 Or iAux > 0 Then lblDesc.Caption = " " & desc
    If Len(desc) = 0 Then desc = "CÓDIGO " & mCod

    If mLinhaObras > 0 Then
        Pintar S_NA_Obras, desc, "Este CÓDIGO já está no follow up. Nada será incluído para não duplicar a obra; abra a obra no painel."
    ElseIf iCrono > 0 Then
        Pintar S_CRONO, desc, "Cadastrada no Cronograma de Obras. As fórmulas do follow up vão trazer os dados ao lado."
    ElseIf iAux > 0 Then
        Pintar S_AUX, desc, "Está na Carteira, mas ainda não no Cronograma de Obras: os dados ao lado vêm da Carteira e as datas podem mudar quando entrar no cronograma."
    Else
        Pintar S_NADA, "CÓDIGO " & mCod & " não encontrado", "Não está no Cronograma de Obras nem na Carteira. Confira o número; se a obra é nova, cadastre-a na base e use ATUALIZAR DADOS no painel antes de incluir."
    End If
End Sub

Private Sub Pintar(ByVal situacao As Long, ByVal titulo As String, ByVal detalhe As String)
    Dim i As Long
    Dim mostrar As Boolean

    mSituacao = situacao
    lblTit.Caption = titulo
    lblSub.Caption = detalhe
    mostrar = (situacao = S_CRONO Or situacao = S_AUX Or situacao = S_NA_Obras)
    For i = 1 To 8
        If Not mostrar Then mVal(i).Caption = ""
    Next i
    If Not mostrar Then lblDesc.Caption = ""

    Select Case situacao
        Case S_CRONO
            Icone Chr$(252), NL_VERDE, NL_VERDE_FUNDO, NL_VERDE
            Primario "Incluir " & Curto(titulo) & " no follow up", True
        Case S_AUX
            Icone Chr$(252), NL_AMARELO, NL_DESTAQUE, NL_AMARELO
            Primario "Incluir " & Curto(titulo) & " no follow up", True
        Case S_NA_Obras
            Icone Chr$(105), NL_AZUL, NL_AZUL_FUNDO, NL_AZUL, "Webdings"
            Primario "Abrir a obra no painel", True
        Case S_NADA
            Icone Chr$(251), NL_CORAL, NL_CORAL_FUNDO, NL_CORAL
            Primario "Incluir no follow up (bloqueado)", False
        Case Else
            Icone "", NL_CINZA, NL_BRANCO, NL_BORDA
            Primario "Incluir no follow up", False
    End Select
End Sub

Private Function Curto(ByVal s As String) As String
    If Len(s) > 30 Then s = Left$(s, 28) & "..."
    Curto = s
End Function

Private Sub Icone(ByVal glifo As String, ByVal cor As Long, ByVal fundo As Long, ByVal borda As Long, _
                  Optional ByVal fonte As String = "Wingdings")
    lblIcone.Font.Name = fonte
    lblIcone.Font.Charset = 2
    lblIcone.Caption = glifo
    lblIcone.ForeColor = cor
    lblCaixa.BackColor = fundo
    lblCaixa.BorderColor = borda
End Sub

Private Sub Primario(ByVal texto As String, ByVal ativo As Boolean)
    lblOkT.Caption = texto
    If ativo Then
        lblOkF.BackColor = NL_NAVY
        lblOkT.ForeColor = NL_BRANCO
    Else
        lblOkF.BackColor = NL_DESLIGADO
        lblOkT.ForeColor = NL_BRANCO
    End If
End Sub

'===============================================================================
' CLIQUES
'===============================================================================
Public Sub TratarClique(ByVal chave As String)
    Select Case chave
        Case "cancelar"
            RaiseEvent Fechar(0)
        Case "ok"
            Select Case mSituacao
                Case S_CRONO, S_AUX
                    Incluir
                Case S_NA_Obras
                    RaiseEvent Fechar(mLinhaObras)
                Case S_NADA
                    lblStatus.ForeColor = NL_CORAL
                    lblStatus.Caption = "Obra não encontrada na base: cadastre-a e atualize os dados antes de incluir."
                Case Else
                    lblStatus.ForeColor = NL_CORAL
                    lblStatus.Caption = "Digite o CÓDIGO."
            End Select
    End Select
End Sub

Private Sub Incluir()
    Dim tb As ListObject
    Dim i As Long
    Dim cA As Long
    Dim cD As Long
    Dim cel As Range
    Dim ev As Boolean
    Dim desc As String

    On Error GoTo Falha
    Set tb = TabelaNL(T_Obras)
    If tb Is Nothing Then
        lblStatus.ForeColor = NL_CORAL
        lblStatus.Caption = "A tabela tb_Obras não foi encontrada."
        Exit Sub
    End If
    If LinhaNaTabela(T_Obras, mCod) > 0 Then
        Conferir
        Exit Sub
    End If
    cA = ColNL(tb, "CÓDIGO")
    cD = ColNL(tb, "NOME")
    ev = Application.EnableEvents
    Application.EnableEvents = False

    i = LinhaNova(tb, cA)
    If i = 0 Then
        Application.EnableEvents = ev
        lblStatus.ForeColor = NL_CORAL
        lblStatus.Caption = "O Excel recusou a nova linha: confira se a aba OBRAS não está protegida e limpe o filtro da tabela."
        Exit Sub
    End If

    Set cel = tb.DataBodyRange.Cells(i, cA)
    cel.NumberFormat = "@"
    cel.Value = CStr(mCod)
    Application.EnableEvents = ev
    On Error Resume Next
    Application.Calculate
    On Error GoTo Falha

    If cD > 0 Then desc = TextoNL(tb.DataBodyRange.Cells(i, cD).Value)
    lblStatus.ForeColor = NL_CINZA
    lblStatus.Caption = ""
    RaiseEvent Fechar(i)
    Exit Sub
Falha:
    Application.EnableEvents = True
    lblStatus.ForeColor = NL_CORAL
    lblStatus.Caption = "Não foi possível incluir: " & Err.Description
End Sub

' Linha para a obra nova: reaproveita uma linha sem CÓDIGO ou cria outra no
' fim, repetindo as fórmulas da linha de cima (0 = o Excel recusou).
Private Function LinhaNova(ByVal tb As ListObject, ByVal cA As Long) As Long
    Dim i As Long
    Dim antes As Long
    Dim c As Long
    Dim ref As Range
    Dim nova As Range

    If Not tb.DataBodyRange Is Nothing Then
        For i = 1 To tb.ListRows.Count
            If Len(TextoNL(tb.DataBodyRange.Cells(i, cA).Value)) = 0 Then
                If Not tb.DataBodyRange.Cells(i, cA).HasFormula Then
                    LinhaNova = i
                    Exit Function
                End If
            End If
        Next i
    End If

    antes = tb.ListRows.Count
    If antes > 0 Then Set ref = tb.ListRows(antes).Range
    On Error Resume Next
    tb.ListRows.Add
    If tb.ListRows.Count <= antes Then
        Err.Clear
        tb.Resize tb.Range.Resize(tb.Range.Rows.Count + 1)
    End If
    On Error GoTo 0
    If tb.ListRows.Count <= antes Then Exit Function

    Set nova = tb.ListRows(tb.ListRows.Count).Range
    If Not ref Is Nothing Then
        For c = 1 To tb.ListColumns.Count
            If c <> cA Then
                If ref.Cells(1, c).HasFormula And Not nova.Cells(1, c).HasFormula Then
                    On Error Resume Next
                    nova.Cells(1, c).Formula2 = ref.Cells(1, c).Formula2
                    If Err.Number <> 0 Then
                        Err.Clear
                        nova.Cells(1, c).Formula = ref.Cells(1, c).Formula
                    End If
                    On Error GoTo 0
                End If
            End If
        Next c
    End If
    LinhaNova = tb.ListRows.Count
End Function
