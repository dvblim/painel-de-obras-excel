Option Explicit

'---------------- API: tamanho real da tela (para ajustar o painel)
Private Type RECT_TELA
    Esq As Long
    Topo As Long
    Direita As Long
    Baixo As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function SystemParametersInfoA Lib "user32" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function CreateFontA Lib "gdi32" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As LongPtr
    Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function GetTextFaceA Lib "gdi32" (ByVal hdc As LongPtr, ByVal nCount As Long, ByVal lpFaceName As String) As Long
#Else
    Private Declare Function SystemParametersInfoA Lib "user32" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare Function CreateFontA Lib "gdi32" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function GetTextFaceA Lib "gdi32" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFaceName As String) As Long
#End If

'====================================================================
' PAINEL DE OBRAS - FORMULÁRIO PRINCIPAL
' Todos os controles são criados por código (não precisa desenhar nada).
'
' Trabalha sobre a tabela tb_Obras (aba OBRAS) e usa as rotinas do módulo
' mPainelBase e das classes clsObra e clsCampoData.
'
'   SALVAR ............ grava as datas alteradas e a nova observação
'   GERAR ATA ......... Word, com o texto inteiro da ÚLTIMO REGISTRO
'   ATUALIZAR DADOS ... atualiza as consultas e recarrega o painel
'   NOVA OBRA ......... inclui uma linha pelo CÓDIGO; as fórmulas da
'                       tabela puxam os demais dados
'====================================================================

Private Const TB_OBRAS As String = "tb_Obras"

'---------------- IDENTIDADE VISUAL (paleta institucional, sem o rosa)
Private Const COR_ESCURO As Long = 3877150       'RGB(30, 41, 59)
Private Const COR_AZUL As Long = 8950797        'RGB(13, 148, 136)
Private Const COR_AZUL_CLARO As Long = 15858636  'RGB(204, 251, 241)
Private Const COR_AMARELO As Long = 2408443      'RGB(251, 191, 36)
Private Const COR_VERMELHO As Long = 2500316     'RGB(220, 38, 38)
Private Const COR_CINZA As Long = 9139300        'RGB(100, 116, 139)
'tons neutros de apoio (fundos e bordas)
Private Const COR_FUNDO As Long = 16579320       'RGB(248, 250, 252)
Private Const COR_LEITURA As Long = 16381425     'RGB(241, 245, 249)
Private Const COR_EDICAO As Long = 15465471      'RGB(255, 251, 235)
Private Const COR_BORDA As Long = 14800331       'RGB(203, 213, 225)

'Fontes da marca. Se não estiverem instaladas no computador, o painel
'usa a fonte reserva automaticamente.
Private Const FONTE_MARCA As String = "Segoe UI"
Private Const FONTE_MARCA_TITULO As String = "Segoe UI Black"
Private Const FONTE_RESERVA As String = "Segoe UI"

Private mFonte As String             'fonte de textos e subtítulos
Private mFonteTitulo As String       'fonte do nome da obra
Private mTituloNegrito As Boolean    'True quando o título usa a fonte comum

Private Const LARG As Single = 1092   '880 + 212 do calendário
Private Const ALT As Single = 522   '496 + 26 da faixa de avisos

Private tb As ListObject

Private mMapa() As Long          'posição na lista -> índice ListRow
Private mQtd As Long
Private mAtual As Long           'índice ListRow da obra aberta
Private mObra As clsObra   'objeto da obra aberta
Private mCarregando As Boolean
Private mEditandoHist As Boolean
Private mHistOriginal As String
Private mPendenteIr As Long

'Campos da coluna da esquerda (montados em MontarCamposInfo)
Private mQtdCampos As Long
Private mCampoIdx() As Long          'índice da coluna na tabela
Private mCampoRot() As String        'rótulo mostrado no painel
Private mCampoTipo() As String       '"D" = data editável | "N" = só consulta
Private mEditavel() As Boolean       'editável NESTA obra (sem fórmula na célula)
Private mOrig() As String            'texto ao carregar a obra
Private mLbl() As MSForms.label
Private mBoxEd() As MSForms.TextBox  'caixa de edição (datas digitadas)
Private mBoxRo() As MSForms.TextBox  'caixa de consulta (vem de fórmula)
Private mCamposData As Collection    'clsCampoData de cada caixa de edição

Private WithEvents txtBusca As MSForms.TextBox
Private WithEvents cboObra As MSForms.ComboBox
Private WithEvents chkVisiveis As MSForms.CheckBox
Private WithEvents btnAnterior As MSForms.CommandButton
Private WithEvents btnProxima As MSForms.CommandButton
Private WithEvents btnEditarHist As MSForms.CommandButton
Private WithEvents btnSalvar As MSForms.CommandButton
Private WithEvents btnAta As MSForms.CommandButton
Private WithEvents btnAtualizar As MSForms.CommandButton
Private WithEvents btnNova As MSForms.CommandButton
Private WithEvents btnIrPlanilha As MSForms.CommandButton
Private WithEvents btnMarcar As MSForms.CommandButton
Private WithEvents chkMarcadas As MSForms.CheckBox
Private WithEvents lblMcClique As MSForms.Label
Private mMarcas As cMarcadores
Private mMcIcone As cMarcaIcone
Private lblMcFundo As MSForms.Label
Private lblMcBarra As MSForms.Label
Private lblMcTit As MSForms.Label
Private lblMcPrazo As MSForms.Label
Private lblMcTexto As MSForms.Label
Private lblMcPor As MSForms.Label
Private lblMcMedida As MSForms.Label
Private mCal As cCalendario
Private mFerCampos As Collection
Private mFerTag() As MSForms.Label
Private mFerQtd As Long
Private mFerIdx As Long
Private mFerUltimo As Long
Private mFerSug1 As Date
Private mFerSug2 As Date
Private mFerUF As String
Private mFerCid As String
Private WithEvents mNova As cNovaObraPainel
Private fraNova As MSForms.Frame
Private WithEvents btnEntrega As MSForms.CommandButton
Private WithEvents mEntrega As cEntregaPainel
Private fraEntrega As MSForms.Frame
Private WithEvents btnFechar As MSForms.CommandButton

Private lblNome As MSForms.label
Private lblSub As MSForms.label
Private lblContador As MSForms.label
Private txtHist As MSForms.TextBox
Private txtNova As MSForms.TextBox
Private fraInfo As MSForms.Frame


'====================================================================
' INICIALIZAÇÃO
'====================================================================
Private Sub UserForm_Initialize()

    Set tb = PO_TabelaPorNome(TB_OBRAS)

    DefinirFontes

    Me.Caption = "PAINEL DE OBRAS - DADOS DA OBRA"
    Me.BackColor = COR_FUNDO
    Me.Width = LARG + 6
    Me.Height = ALT + 28

    mCarregando = True
    MontarTela
    chkVisiveis.value = True          'começa pelas obras visíveis no filtro da planilha
    mCarregando = False

    If Not tb Is Nothing Then

        MontarCamposInfo
        FerPrepararCampos
        FiltrarLista

        If mQtd = 0 Then
            'nenhuma obra visível no filtro: mostra todas
            mCarregando = True
            chkVisiveis.value = False
            mCarregando = False
            FiltrarLista
        End If

        If mQtd > 0 Then CarregarObra mMapa(1)

    End If

    AjustarZoom

End Sub

Private Sub UserForm_Activate()
    RolarHistoricoFim
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not ConfirmarPendencias() Then Cancel = True
    If Cancel = 0 Then FerDesligar
End Sub

'Chamado pelo módulo: abre o painel numa obra (0 = mantém a atual)
Public Sub IrParaLinha(ByVal idx As Long)

    If tb Is Nothing Then Exit Sub
    If idx < 1 Or idx > tb.ListRows.Count Then Exit Sub
    If idx = mAtual Then Exit Sub
    If Not ConfirmarPendencias() Then Exit Sub

    If PosicaoNaLista(idx) = 0 Then
        'obra fora do filtro atual: limpa a busca
        mCarregando = True
        txtBusca.Text = ""
        chkVisiveis.value = False
        mCarregando = False
        FiltrarLista
    End If

    CarregarObra idx

End Sub


'====================================================================
' CRIAÇÃO DOS CONTROLES
'====================================================================
Private Function NovoLabel(ByVal pai As Object, ByVal nome As String, ByVal cap As String, _
                           ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                           Optional ByVal tam As Single = 9, Optional ByVal negrito As Boolean = False, _
                           Optional ByVal cor As Long = -1) As MSForms.label

    Dim c As MSForms.label

    Set c = pai.Controls.Add("Forms.Label.1", nome, True)

    With c
        .Caption = cap
        .Left = l: .Top = t: .Width = w: .Height = h
        .BackStyle = fmBackStyleTransparent
        .Font.Name = mFonte
        .Font.Size = tam
        .Font.Bold = negrito
        If cor <> -1 Then .ForeColor = cor Else .ForeColor = COR_ESCURO
        .WordWrap = True
    End With

    Set NovoLabel = c

End Function

Private Function NovoTexto(ByVal pai As Object, ByVal nome As String, _
                           ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                           Optional ByVal multi As Boolean = False, _
                           Optional ByVal bloqueado As Boolean = False) As MSForms.TextBox

    Dim c As MSForms.TextBox

    Set c = pai.Controls.Add("Forms.TextBox.1", nome, True)

    With c
        .Left = l: .Top = t: .Width = w: .Height = h
        .Font.Name = mFonte
        .Font.Size = 10
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COR_BORDA
        .Multiline = multi
        .WordWrap = multi
        .EnterKeyBehavior = multi
        If multi Then .ScrollBars = fmScrollBarsVertical
        .Locked = bloqueado
        If bloqueado Then .BackColor = COR_LEITURA Else .BackColor = vbWhite
    End With

    Set NovoTexto = c

End Function

Private Function NovoBotao(ByVal nome As String, ByVal cap As String, _
                           ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                           Optional ByVal corFundo As Long = -1, Optional ByVal corTexto As Long = -1) As MSForms.CommandButton

    Dim c As MSForms.CommandButton

    Set c = Me.Controls.Add("Forms.CommandButton.1", nome, True)

    With c
        .Caption = cap
        .Left = l: .Top = t: .Width = w: .Height = h
        .Font.Name = mFonte
        .Font.Size = 10
        .Font.Bold = True
        .TakeFocusOnClick = False
        If corFundo <> -1 Then .BackColor = corFundo
        If corTexto <> -1 Then .ForeColor = corTexto
    End With

    Set NovoBotao = c

End Function

Private Sub MontarTela()

    Dim l As MSForms.label
    Dim y As Single

    '---------------- faixa de título
    Set l = NovoLabel(Me, "lblFaixa", "", 0, 0, LARG + 10, 60)
    l.BackStyle = fmBackStyleOpaque
    l.BackColor = COR_ESCURO

    Set l = NovoLabel(Me, "lblFaixaAmarela", "", 0, 60, LARG + 10, 3)
    l.BackStyle = fmBackStyleOpaque
    l.BackColor = COR_AMARELO

    NovoLabel Me, "lblTitulo", "FOLLOW UP DE OBRAS  |  DADOS DA OBRA", 14, 5, 520, 13, 9, True, COR_AZUL_CLARO
    Set lblNome = NovoLabel(Me, "lblNome", "", 14, 16, LARG - 28, 26, 17, True, vbWhite)
    lblNome.WordWrap = False
    lblNome.Font.Name = mFonteTitulo
    lblNome.Font.Bold = mTituloNegrito
    Set lblSub = NovoLabel(Me, "lblSub", "", 14, 42, LARG - 28, 16, 10, False, RGB(203, 213, 225))
    lblSub.WordWrap = False

    '---------------- navegação
    y = 72
    NovoLabel Me, "lblBusca", "Buscar:", 12, y + 3, 40, 16, 10, True
    Set txtBusca = NovoTexto(Me, "txtBusca", 52, y, 132, 21)
    txtBusca.ControlTipText = "Filtra por CÓDIGO, nome ou cidade"

    NovoLabel Me, "lblObra", "Obra:", 192, y + 3, 32, 16, 10, True
    Set cboObra = Me.Controls.Add("Forms.ComboBox.1", "cboObra", True)
    With cboObra
        .Left = 224: .Top = y: .Width = 292: .Height = 21
        .Style = fmStyleDropDownList
        .Font.Name = mFonte: .Font.Size = 10
        .ListRows = 20
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COR_BORDA
    End With

    Set chkVisiveis = Me.Controls.Add("Forms.CheckBox.1", "chkVisiveis", True)
    With chkVisiveis
        .Left = 524: .Top = y + 2: .Width = 116: .Height = 18
        .Caption = "Só visíveis no filtro"
        .Font.Name = mFonte: .Font.Size = 9
        .BackStyle = fmBackStyleTransparent
        .ControlTipText = "Respeita os filtros da aba OBRAS"
    End With

    Set btnAnterior = NovoBotao("cmdAnterior", "< Anterior", 644, y, 70, 22)
    Set btnProxima = NovoBotao("cmdProxima", "Próxima >", 718, y, 70, 22)
    Set lblContador = NovoLabel(Me, "lblContador", "", 792, y + 3, LARG - 800, 16, 9.5, False, COR_CINZA)
    lblContador.TextAlign = fmTextAlignRight

    '---------------- coluna esquerda: dados da obra
    NovoLabel Me, "lblSecInfo", "DADOS DA OBRA", 12, 100, 200, 17, 10.5, True, COR_ESCURO

    Set fraInfo = Me.Controls.Add("Forms.Frame.1", "fraInfo", True)
    With fraInfo
        .Left = 12: .Top = 118: .Width = 330: .Height = 324
        .Caption = ""
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COR_BORDA
        .BackColor = vbWhite
    End With

    '---------------- coluna direita: observações
    NovoLabel Me, "lblSecHist", "HISTÓRICO DE REGISTROS", 354, 100, 300, 17, 10.5, True, COR_ESCURO
    Set btnEditarHist = NovoBotao("cmdEditarHist", "Editar histórico", LARG - 132, 97, 120, 20)
    btnEditarHist.Font.Size = 9
    Set txtHist = NovoTexto(Me, "txtHist", 354, 118, LARG - 366, 228, True, True)

    NovoLabel Me, "lblSecNova", "NOVO REGISTRO", 354, 352, 150, 17, 10.5, True, COR_ESCURO
    NovoLabel Me, "lblDicaNova", "(data e hora são inseridas automaticamente ao salvar)", 504, 354, 340, 16, 9, False, COR_CINZA
    Set txtNova = NovoTexto(Me, "txtNova", 354, 370, LARG - 366, 72, True, False)
    txtNova.BackColor = COR_EDICAO

    '---------------- botões
    y = 454
    Set btnSalvar = NovoBotao("cmdSalvar", "SALVAR", 12, y, 100, 30, COR_AZUL, vbWhite)
    btnSalvar.ControlTipText = "Grava as datas alteradas e a nova observação"
    Set btnAta = NovoBotao("cmdAta", "GERAR ATA (WORD)", 120, y, 136, 30, COR_ESCURO, vbWhite)
    btnAta.ControlTipText = "Ata com a ÚLTIMO REGISTRO das obras da lista do painel"
    Set btnAtualizar = NovoBotao("cmdAtualizar", "ATUALIZAR DADOS", 264, y, 130, 30, COR_AZUL_CLARO, COR_ESCURO)
    btnAtualizar.ControlTipText = "Atualiza as consultas da planilha e recarrega o painel"
    Set btnNova = NovoBotao("cmdNovaObra", "NOVA OBRA", 402, y, 104, 30, COR_AMARELO, COR_ESCURO)
    btnNova.ControlTipText = "Inclui uma obra no follow up pelo CÓDIGO"
    Set btnIrPlanilha = NovoBotao("cmdIrPlanilha", "IR PARA A LINHA", 514, y, 120, 30)
    Set btnFechar = NovoBotao("cmdFechar", "FECHAR", LARG - 112, y, 100, 30)

    '---------------- calendário, feriados e marcador (instalado por mInstaladorPainelObras)
    ExtrasMontar

    '---------------- OBRA ENTREGUE (instalado por mInstaladorEntregueObras)
    Set btnEntrega = NovoBotao("cmdEntrega", "OBRA ENTREGUE", btnIrPlanilha.Left + btnIrPlanilha.Width + 8, btnIrPlanilha.Top, 140, btnIrPlanilha.Height, RGB(22, 163, 74), vbWhite)
    If btnEntrega.Left + btnEntrega.Width > LARG - 120 Then btnEntrega.Width = LARG - 120 - btnEntrega.Left
    btnEntrega.ControlTipText = "Leva esta obra para ENTREGUES (confirmação dentro do painel)"

End Sub

Private Sub MontarCamposInfo()

    Dim lista As Variant
    Dim i As Long, c As Long, n As Long
    Dim y As Single
    Dim campo As clsCampoData

    'Coluna da tabela tb_Obras, rótulo no painel e tipo:
    '   D = data (editável quando a célula não tem fórmula)
    '   N = só consulta
    'Para incluir ou tirar um campo do painel, basta mexer nesta lista.
    lista = Array( _
        Array("ALVARÁ", "ALVARÁ", "D"), _
        Array("PROJETO EXECUTIVO", "PROJETO EXECUTIVO", "D"), _
        Array("INÍCIO DA OBRA", "INÍCIO DA OBRA", "D"), _
        Array("PRAZO (DIAS)", "PRAZO (DIAS)", "N"), _
        Array("FIM DA OBRA", "FIM DA OBRA", "D"), _
        Array("ENTREGA", "ENTREGA", "D"))

    n = UBound(lista) - LBound(lista) + 1

    ReDim mCampoIdx(1 To n)
    ReDim mCampoRot(1 To n)
    ReDim mCampoTipo(1 To n)
    ReDim mEditavel(1 To n)
    ReDim mOrig(1 To n)
    ReDim mLbl(1 To n)
    ReDim mBoxEd(1 To n)
    ReDim mBoxRo(1 To n)

    Set mCamposData = New Collection
    mQtdCampos = 0
    y = 12

    For i = LBound(lista) To UBound(lista)

        c = PO_ColIdx(tb, CStr(lista(i)(0)))

        If c > 0 Then

            mQtdCampos = mQtdCampos + 1
            n = mQtdCampos

            mCampoIdx(n) = c
            mCampoRot(n) = CStr(lista(i)(1))
            mCampoTipo(n) = CStr(lista(i)(2))

            Set mLbl(n) = NovoLabel(fraInfo, "lblCampo" & n, mCampoRot(n), 6, y + 3, 120, 16, 8, True, COR_ESCURO)
            mLbl(n).WordWrap = False
            mLbl(n).ControlTipText = tb.ListColumns(c).Name

            'duas caixas no mesmo lugar: a de consulta (cinza) e a de edição;
            'o painel mostra uma ou outra conforme a célula da obra
            Set mBoxRo(n) = NovoTexto(fraInfo, "txtRo" & n, 128, y, 180, 21, False, True)
            mBoxRo(n).TabStop = False

            Set mBoxEd(n) = NovoTexto(fraInfo, "txtEd" & n, 128, y, 180, 21, False, False)
            mBoxEd(n).BorderColor = COR_AMARELO
            mBoxEd(n).MaxLength = 10

            'máscara de data, vermelho se inválida, duplo clique = hoje
            Set campo = New clsCampoData
            Set campo.Caixa = mBoxEd(n)
            mCamposData.Add campo

            y = y + 46

        End If

    Next i

    NovoLabel fraInfo, "lblCampoDica", "Borda amarela = pode alterar (15112026 ou dd/mm/aaaa; duplo clique = hoje). " & _
              "Cinza = vem da fórmula da tabela.", 6, y + 2, 302, 26, 8, False, COR_CINZA

End Sub

Private Sub AjustarZoom()

    'Ajusta o painel ao tamanho REAL da tela (área sem a barra de tarefas)
    Dim telaW As Single, telaH As Single
    Dim f As Single

    On Error GoTo Fim

    If Not AreaTelaPt(telaW, telaH) Then Exit Sub

    f = (telaH - 16) / Me.Height
    If (telaW - 16) / Me.Width < f Then f = (telaW - 16) / Me.Width
    If f > 1.3 Then f = 1.3
    If f < 0.75 Then f = 0.75
    If Abs(f - 1) < 0.02 Then Exit Sub

    Me.Zoom = Int(f * 100)
    Me.Width = Me.Width * f
    Me.Height = (Me.Height - 28) * f + 28

Fim:
End Sub


'====================================================================
' LISTA DE OBRAS / NAVEGAÇÃO
'====================================================================

Private Sub FiltrarLista()

    Dim i As Long
    Dim f As String
    Dim txt As String
    Dim lj As New clsObra
    Dim ok As Boolean

    If tb Is Nothing Then Exit Sub

    f = PO_Norm(txtBusca.Text)
    mQtd = 0

    If tb.ListRows.Count = 0 Then Exit Sub
    ReDim mMapa(1 To tb.ListRows.Count)

    mCarregando = True
    cboObra.Clear

    For i = 1 To tb.ListRows.Count

        lj.Carregar tb, i
        txt = lj.rotulo
        ok = True

        If f <> "" Then
            ok = (InStr(1, PO_Norm(txt & " " & lj.texto(PO_COL_CIDADE)), f) > 0)
        End If

        If ok And Not chkMarcadas Is Nothing Then
            If chkMarcadas.value = True Then ok = MarcasPainel().Marcada(lj.CodObra)
        End If
        If ok And chkVisiveis.value = True Then
            ok = lj.Visivel
        End If

        If ok Then
            mQtd = mQtd + 1
            mMapa(mQtd) = i
            cboObra.AddItem txt
        End If

    Next i

    SelecionarNoCombo mAtual
    mCarregando = False

End Sub

Private Function PosicaoNaLista(ByVal idx As Long) As Long

    Dim i As Long

    For i = 1 To mQtd
        If mMapa(i) = idx Then
            PosicaoNaLista = i
            Exit Function
        End If
    Next i

End Function

Private Sub SelecionarNoCombo(ByVal idx As Long)

    Dim pos As Long
    Dim antes As Boolean

    antes = mCarregando
    mCarregando = True

    pos = PosicaoNaLista(idx)

    If pos > 0 Then
        cboObra.ListIndex = pos - 1
        lblContador.Caption = pos & " de " & mQtd
    Else
        cboObra.ListIndex = -1
        lblContador.Caption = "- de " & mQtd
    End If

    mCarregando = antes

End Sub

Private Sub txtBusca_Change()

    If mCarregando Then Exit Sub
    FiltrarLista

    'Se sobrou só uma obra, abre direto
    If mQtd = 1 Then
        If mMapa(1) <> mAtual Then
            If ConfirmarPendencias() Then CarregarObra mMapa(1)
        End If
    End If

End Sub

Private Sub chkVisiveis_Click()
    If mCarregando Then Exit Sub
    FiltrarLista
End Sub

Private Sub cboObra_Change()

    Dim novo As Long

    If mCarregando Then Exit Sub
    If cboObra.ListIndex < 0 Then Exit Sub

    novo = mMapa(cboObra.ListIndex + 1)
    If novo = mAtual Then Exit Sub

    If Not ConfirmarPendencias() Then
        SelecionarNoCombo mAtual
        Exit Sub
    End If

    CarregarObra novo

End Sub

Private Sub btnAnterior_Click()
    Navegar -1
End Sub

Private Sub btnProxima_Click()
    Navegar 1
End Sub

Private Sub Navegar(ByVal passo As Long)

    Dim pos As Long

    If mQtd = 0 Then Exit Sub

    pos = PosicaoNaLista(mAtual)

    If pos = 0 Then
        pos = 1
    Else
        pos = pos + passo
    End If

    If pos < 1 Or pos > mQtd Then Exit Sub
    If Not ConfirmarPendencias() Then Exit Sub

    CarregarObra mMapa(pos)

End Sub


'====================================================================
' CARREGAR DADOS
'====================================================================
Private Sub CarregarObra(ByVal idx As Long)

    If tb Is Nothing Then Exit Sub
    If idx < 1 Or idx > tb.ListRows.Count Then Exit Sub

    mCarregando = True
    mAtual = idx
    Set mObra = New clsObra
    mObra.Carregar tb, idx

    CarregarCampos
    CarregarObservacoes

    SelecionarNoCombo mAtual
    mCarregando = False

    RolarHistoricoFim

End Sub

Private Sub CarregarCampos()

    Dim i As Long
    Dim cel As Range
    Dim txt As String

    lblNome.Caption = mObra.nome
    lblSub.Caption = mObra.Resumo

    If mQtdCampos = 0 Then Exit Sub

    'tira o cursor das caixas de data antes de mexer nelas
    On Error Resume Next
    If Me.Visible Then txtNova.SetFocus
    On Error GoTo 0

    For i = 1 To mQtdCampos

        Set cel = mObra.Linha.Range.Cells(1, mCampoIdx(i))
        txt = mObra.TextoColuna(mCampoIdx(i))

        'data digitada = editável; célula com fórmula = só consulta
        mEditavel(i) = (mCampoTipo(i) = "D") And Not CBool(cel.HasFormula)

        mBoxRo(i).Text = txt

        If mEditavel(i) Then
            mBoxEd(i).Text = txt
        Else
            mBoxEd(i).Text = ""
        End If

        mOrig(i) = mBoxEd(i).Text
        mBoxEd(i).BackColor = vbWhite

        mBoxEd(i).Visible = mEditavel(i)
        mBoxRo(i).Visible = Not mEditavel(i)

        If mEditavel(i) Then
            mBoxEd(i).ControlTipText = tb.ListColumns(mCampoIdx(i)).Name & "  |  duplo clique = hoje"
        ElseIf cel.HasFormula Then
            mBoxRo(i).ControlTipText = "Preenchido pela fórmula da tabela"
        Else
            mBoxRo(i).ControlTipText = ""
        End If

    Next i

    ExtrasAtualizar
End Sub

'A data i foi alterada pelo usuário?
Private Function DataAlterada(ByVal i As Long) As Boolean
    If mEditavel(i) Then DataAlterada = (mBoxEd(i).Text <> mOrig(i))
End Function

Private Function DatasAlteradas() As Boolean

    Dim i As Long

    If mObra Is Nothing Then Exit Function

    For i = 1 To mQtdCampos
        If DataAlterada(i) Then
            DatasAlteradas = True
            Exit Function
        End If
    Next i

End Function

Private Sub CarregarObservacoes()

    txtHist.Text = PO_ParaCaixa(mObra.Historico)
    mHistOriginal = txtHist.Text
    txtNova.Text = ""

    ModoEdicaoHistorico False

End Sub

Private Sub RolarHistoricoFim()

    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    txtHist.SetFocus
    txtHist.SelStart = Len(txtHist.Text)
    txtHist.SelLength = 0
    txtNova.SetFocus

End Sub


'====================================================================
' OBSERVAÇÕES
'====================================================================
Private Sub ModoEdicaoHistorico(ByVal ligado As Boolean)

    mEditandoHist = ligado
    txtHist.Locked = Not ligado

    If ligado Then
        txtHist.BackColor = COR_EDICAO
        btnEditarHist.Caption = "Cancelar edição"
        btnEditarHist.ForeColor = COR_VERMELHO
    Else
        txtHist.BackColor = RGB(248, 249, 251)
        btnEditarHist.Caption = "Editar histórico"
        btnEditarHist.ForeColor = COR_ESCURO
    End If

End Sub

Private Function HistoricoAlterado() As Boolean
    HistoricoAlterado = mEditandoHist And (PO_ParaCaixa(txtHist.Text) <> PO_ParaCaixa(mHistOriginal))
End Function

Private Function TemPendencias() As Boolean
    If mAtual = 0 Then Exit Function
    TemPendencias = (Trim$(Replace(Replace(txtNova.Text, vbCr, ""), vbLf, "")) <> "") _
                    Or HistoricoAlterado() Or DatasAlteradas()
End Function

'True = pode seguir | False = usuário cancelou
Private Function ConfirmarPendencias() As Boolean

    Dim r As VbMsgBoxResult

    If Not TemPendencias() Then
        ConfirmarPendencias = True
        Exit Function
    End If

    r = MsgBox("Há alterações não salvas em '" & lblNome.Caption & "'." & vbCrLf & vbCrLf & _
               "Deseja salvar antes de continuar?", vbYesNoCancel + vbQuestion, "Painel de Obras")

    Select Case r
        Case vbYes:  ConfirmarPendencias = SalvarObra(True)
        Case vbNo:   ConfirmarPendencias = True
        Case Else:   ConfirmarPendencias = False
    End Select

End Function

Private Sub btnEditarHist_Click()

    If mAtual = 0 Then Exit Sub

    If mEditandoHist Then
        If HistoricoAlterado() Then
            If MsgBox("Descartar as alterações feitas no histórico?", vbYesNo + vbQuestion, "Painel de Obras") = vbNo Then Exit Sub
        End If
        txtHist.Text = mHistOriginal
        ModoEdicaoHistorico False
    Else
        ModoEdicaoHistorico True
        On Error Resume Next
        txtHist.SetFocus
    End If

End Sub

Private Sub btnSalvar_Click()
    SalvarObra False
End Sub

Private Function SalvarObra(ByVal silencioso As Boolean) As Boolean

    Dim nova As String
    Dim histMudou As Boolean
    Dim datasMudaram As Boolean
    Dim obsMudou As Boolean
    Dim ok As Boolean
    Dim i As Long, r As Long
    Dim qtdDatas As Long, falhas As Long
    Dim d As Date, vz As Boolean
    Dim erros As String, msg As String

    If mObra Is Nothing Then Exit Function
    If Not FerPodeSalvar() Then Exit Function

    nova = Trim$(PO_TrimFim(PO_ParaCaixa(txtNova.Text)))
    histMudou = HistoricoAlterado()
    datasMudaram = DatasAlteradas()
    obsMudou = (nova <> "" Or histMudou)

    If Not obsMudou And Not datasMudaram Then
        If Not silencioso Then MsgBox "Não há alterações para salvar.", vbInformation, "Painel de Obras"
        SalvarObra = True
        Exit Function
    End If

    'valida todas as datas antes de gravar qualquer coisa
    If datasMudaram Then

        For i = 1 To mQtdCampos
            If DataAlterada(i) Then
                If Not PO_TextoParaData(mBoxEd(i).Text, d, vz) Then
                    erros = erros & vbCrLf & "  - " & mCampoRot(i) & ": " & mBoxEd(i).Text
                    mBoxEd(i).BackColor = PO_COR_ERRO
                End If
            End If
        Next i

        If erros <> "" Then
            MsgBox "Datas inválidas (use dd/mm/aaaa):" & erros & vbCrLf & vbCrLf & _
                   "Nada foi gravado. Corrija e salve de novo.", vbExclamation, "Painel de Obras"
            Exit Function
        End If

    End If

    Application.ScreenUpdating = False

    If datasMudaram Then
        For i = 1 To mQtdCampos
            If DataAlterada(i) Then
                r = mObra.GravarData(mCampoIdx(i), mBoxEd(i).Text)
                If r = 1 Then qtdDatas = qtdDatas + 1
                If r < 0 Then falhas = falhas + 1
            End If
        Next i
    End If

    ok = True

    If nova <> "" Then
        If histMudou Then
            ok = mObra.AdicionarObservacao(nova, PO_ParaCaixa(txtHist.Text))
        Else
            ok = mObra.AdicionarObservacao(nova)
        End If
    ElseIf histMudou Then
        ok = mObra.SubstituirHistorico(txtHist.Text)
    End If

    Application.ScreenUpdating = True

    'recarrega o que foi gravado (as datas voltam a ser o "original")
    mCarregando = True
    If datasMudaram Then CarregarCampos
    If obsMudou And ok Then CarregarObservacoes
    mCarregando = False

    If obsMudou And ok Then RolarHistoricoFim

    If falhas > 0 Then
        MsgBox qtdDatas & " data(s) gravada(s); " & falhas & " não puderam ser gravadas.", _
               vbExclamation, "Painel de Obras"
        Exit Function
    End If

    If Not ok Then Exit Function

    If Not silencioso Then

        If qtdDatas > 0 Then msg = qtdDatas & " data(s) atualizada(s) na planilha."

        If obsMudou Then
            If msg <> "" Then msg = msg & vbCrLf
            msg = msg & "Observação salva com sucesso!"
        End If

        If msg = "" Then msg = "Nenhuma alteração: os valores já eram os mesmos."

        MsgBox msg, vbInformation, "Painel de Obras"

    End If

    SalvarObra = True

End Function


'====================================================================
' BOTÕES
'====================================================================
Private Sub btnAta_Click()

    If Not ConfirmarPendencias() Then Exit Sub
    PO_GerarAta tb, mMapa, mQtd

End Sub

'---------------------------------------------------------- ATUALIZAR DADOS
Private Sub btnAtualizar_Click()

    Dim antes As Long

    If Not ConfirmarPendencias() Then Exit Sub

    If MsgBox("Atualizar agora todas as consultas da planilha?" & vbCrLf & vbCrLf & _
              "Pode levar alguns minutos. O painel é recarregado no final.", _
              vbYesNo + vbQuestion, "Painel de Obras") = vbNo Then Exit Sub

    antes = mAtual

    'usa a macro de atualização que já existe na planilha; se ela não
    'estiver disponível, atualiza tudo direto
    On Error Resume Next
    Application.Run "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AtualizarTodasAsQueries"

    If Err.Number <> 0 Then
        Err.Clear
        Application.StatusBar = "Atualizando consultas..."
        ThisWorkbook.RefreshAll
        Application.CalculateUntilAsyncQueriesDone
        Application.StatusBar = False
    End If

    Err.Clear
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    On Error GoTo 0

    Recarregar antes

End Sub

'Lê a tabela de novo e reabre a obra indicada (ou a primeira da lista)
Private Sub Recarregar(ByVal idx As Long)

    Set tb = PO_TabelaPorNome(TB_OBRAS)
    If tb Is Nothing Then Exit Sub

    mAtual = 0
    FiltrarLista

    If idx >= 1 And idx <= tb.ListRows.Count Then
        If PosicaoNaLista(idx) > 0 Then
            CarregarObra idx
            Exit Sub
        End If
    End If

    If mQtd > 0 Then CarregarObra mMapa(1)

End Sub

'---------------------------------------------------------- NOVA OBRA

'Abre a obra idx, tirando busca e filtro se ela não estiver na lista
Private Sub MostrarObra(ByVal idx As Long)

    mAtual = 0
    FiltrarLista

    If PosicaoNaLista(idx) = 0 Then
        mCarregando = True
        txtBusca.Text = ""
        chkVisiveis.value = False
        mCarregando = False
        FiltrarLista
    End If

    CarregarObra idx

End Sub

'Acrescenta uma linha no fim da tabela e devolve o índice dela (0 = falhou).
'Com a tabela filtrada o Excel pode recusar a inclusão normal; nesse caso
'a tabela é ampliada em uma linha.
Private Function NovaLinhaTabela() As Long

    Dim antes As Long
    Dim rng As Range
    Dim evt As Boolean

    antes = tb.ListRows.Count
    evt = Application.EnableEvents

    On Error Resume Next

    Application.EnableEvents = False

    tb.ListRows.Add

    If tb.ListRows.Count <= antes Then
        Err.Clear
        Set rng = tb.Range
        tb.Resize rng.Resize(rng.Rows.Count + 1)
    End If

    Application.EnableEvents = evt
    On Error GoTo 0

    If tb.ListRows.Count <= antes Then
        MsgBox "O Excel recusou a nova linha na tabela '" & TB_OBRAS & "'." & vbCrLf & _
               "Confira se a planilha não está protegida e tente limpar o filtro da tabela.", _
               vbExclamation, "Painel de Obras"
        Exit Function
    End If

    NovaLinhaTabela = tb.ListRows.Count

End Function

'Colunas que não são "calculadas" pela tabela, mas têm fórmula na linha
'de cima: a fórmula é repetida na nova linha, para a obra nova puxar os
'dados do mesmo jeito que as demais.
Private Sub CopiarFormulasDaLinhaAnterior(ByVal nova As Long, ByVal idxCod As Long)

    Dim c As Long
    Dim celNova As Range, celAnt As Range
    Dim evt As Boolean

    If nova < 2 Then Exit Sub

    evt = Application.EnableEvents

    On Error Resume Next
    Application.EnableEvents = False

    For c = 1 To tb.ListColumns.Count

        If c <> idxCod Then

            Set celNova = tb.ListRows(nova).Range.Cells(1, c)
            Set celAnt = tb.ListRows(nova - 1).Range.Cells(1, c)

            If Not celNova.HasFormula Then
                If IsEmpty(celNova.value) And celAnt.HasFormula Then
                    celNova.Formula = celAnt.Formula
                End If
            End If

        End If

    Next c

    Application.EnableEvents = evt
    On Error GoTo 0

End Sub

Private Sub btnIrPlanilha_Click()
    If Not mObra Is Nothing Then mObra.IrParaPlanilha
End Sub

Private Sub btnFechar_Click()
    Unload Me
End Sub


'====================================================================
' FONTES DA MARCA (com reserva se não estiverem instaladas)
'====================================================================
Private Sub DefinirFontes()

    If FonteInstalada(FONTE_MARCA) Then
        mFonte = FONTE_MARCA
    Else
        mFonte = FONTE_RESERVA
    End If

    If FonteInstalada(FONTE_MARCA_TITULO) Then
        mFonteTitulo = FONTE_MARCA_TITULO
        mTituloNegrito = False
    Else
        mFonteTitulo = mFonte
        mTituloNegrito = True
    End If

End Sub

'Pede a fonte ao Windows e confere se ele entregou a mesma (se não
'existir, o Windows devolve outra no lugar).
Private Function FonteInstalada(ByVal nome As String) As Boolean

#If VBA7 Then
    Dim hdc As LongPtr
    Dim hFonte As LongPtr
    Dim hAntes As LongPtr
#Else
    Dim hdc As Long
    Dim hFonte As Long
    Dim hAntes As Long
#End If
    Dim buf As String
    Dim n As Long

    On Error GoTo Fim

    hdc = GetDC(0)
    If hdc = 0 Then Exit Function

    hFonte = CreateFontA(-12, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 0, 0, nome)

    If hFonte <> 0 Then

        hAntes = SelectObject(hdc, hFonte)

        buf = String$(64, vbNullChar)
        n = GetTextFaceA(hdc, 64, buf)

        If n > 1 Then
            FonteInstalada = (StrComp(Left$(buf, n - 1), nome, vbTextCompare) = 0)
        End If

        SelectObject hdc, hAntes
        DeleteObject hFonte

    End If

    ReleaseDC 0, hdc

Fim:
End Function


'====================================================================
' ÁREA ÚTIL DA TELA EM PONTOS (sem a barra de tarefas)
'====================================================================
Private Function AreaTelaPt(ByRef largura As Single, ByRef altura As Single) As Boolean

    Dim r As RECT_TELA
    Dim dpi As Long
#If VBA7 Then
    Dim hdc As LongPtr
#Else
    Dim hdc As Long
#End If

    On Error GoTo Fim

    If SystemParametersInfoA(48, 0, r, 0) = 0 Then Exit Function    'SPI_GETWORKAREA

    hdc = GetDC(0)
    dpi = GetDeviceCaps(hdc, 90)                                     'LOGPIXELSY
    ReleaseDC 0, hdc
    If dpi <= 0 Then dpi = 96

    largura = (r.Direita - r.Esq) * 72 / dpi
    altura = (r.Baixo - r.Topo) * 72 / dpi
    AreaTelaPt = (largura > 200 And altura > 200)

Fim:
End Function


'====================================================================
' ABREVIA CABEÇALHOS LONGOS (nome completo fica no ControlTipText)
'====================================================================
Private Function RotuloCurto(ByVal cab As String) As String

    Const LIMITE As Long = 22

    Dim s As String
    Dim abrev As Variant
    Dim i As Long
    Dim p1 As Long, p2 As Long

    s = UCase$(Trim$(Replace(Replace(cab, vbCr, " "), vbLf, " ")))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    If Len(s) <= LIMITE Then RotuloCurto = s: Exit Function

    '0) parênteses com texto longo: mantém só o que vem antes
    p1 = InStr(s, "(")
    If p1 > 1 Then
        p2 = InStr(p1, s, ")")
        If p2 = 0 Then p2 = Len(s)
        If (p2 - p1) > 16 And Len(Trim$(Left$(s, p1 - 1))) <= LIMITE Then
            RotuloCurto = Trim$(Left$(s, p1 - 1))
            Exit Function
        End If
    End If

    '1) abreviações comuns, uma por vez, até caber
    abrev = Array("PROJETO BÁSICO", "PROJ. BÁS.", "PROJETO EXECUTIVO", "PROJ. EXEC.", _
                  "CONTRATO ASSINADO", "CONTRATO", "CONTRATO", "CONTR.", _
                  "TIPO DE OBRA", "TIPO", "APROVAÇÃO", "APROV.", _
                  "PREVISTO", "PREV.", "PREVISTA", "PREV.", "ENTREGA", "ENTR.", "INÍCIO", "INÍC.")

    For i = LBound(abrev) To UBound(abrev) - 1 Step 2
        s = Replace(s, abrev(i), abrev(i + 1))
        If Len(s) <= LIMITE Then RotuloCurto = s: Exit Function
    Next i

    '2) remove o texto entre parênteses
    p1 = InStr(s, "(")
    If p1 > 0 Then
        p2 = InStr(p1, s, ")")
        If p2 = 0 Then p2 = Len(s)
        s = Trim$(Left$(s, p1 - 1) & " " & Mid$(s, p2 + 1))
        If Len(s) <= LIMITE Then RotuloCurto = s: Exit Function
    End If

    '3) corta
    RotuloCurto = Left$(s, LIMITE - 1) & "."

End Function


'====================================================================
' CALENDÁRIO, FERIADOS E MARCADOR (instalado por mInstaladorPainelObras)
'   - data digitada em fim de semana ou feriado não é aceita ao salvar
'   - datas que vêm da fórmula (cronograma) só recebem alerta
'====================================================================
Private Sub ExtrasMontar()
    Const LARG_CAL As Single = 212
    Dim nm As Variant
    Dim c As Object

    'abre espaço para o calendário: a coluna da direita anda para a direita
    On Error Resume Next
    For Each nm In Array("lblSecHist", "lblSecNova", "lblDicaNova")
        Me.Controls(CStr(nm)).Left = Me.Controls(CStr(nm)).Left + LARG_CAL
    Next nm
    For Each nm In Array("txtHist", "txtNova")
        Me.Controls(CStr(nm)).Left = Me.Controls(CStr(nm)).Left + LARG_CAL
        Me.Controls(CStr(nm)).Width = Me.Controls(CStr(nm)).Width - LARG_CAL
    Next nm
    'botões de baixo descem para a faixa de avisos
    For Each c In Me.Controls
        If c.Parent.Name = Me.Name Then
            If c.Top >= 448 Then c.Top = c.Top + 26
        End If
    Next c
    On Error GoTo 0

    NovoLabel Me, "lblSecCal", "CALENDÁRIO", 354, 100, 200, 17, 10.5, True, COR_ESCURO
    Set mCal = New cCalendario
    mCal.Criar Me, 354, 118, 202, 12, 452, LARG - 24

    '---- marcador
    Set btnMarcar = NovoBotao("cmdMarcar", "MARCAR P/ ANALISAR", LARG - 264, 97, 126, 20)
    btnMarcar.Font.Size = 9
    btnMarcar.ControlTipText = "Marca esta obra para analisar depois (marca-página)"

    chkVisiveis.Caption = "Visíveis"
    chkVisiveis.Width = 58
    Set chkMarcadas = Me.Controls.Add("Forms.CheckBox.1", "chkMarcadas", True)
    chkMarcadas.Left = chkVisiveis.Left + 60
    chkMarcadas.Top = chkVisiveis.Top
    chkMarcadas.Width = 58
    chkMarcadas.Height = chkVisiveis.Height
    chkMarcadas.Caption = "Marcadas"
    chkMarcadas.Font.Name = mFonte
    chkMarcadas.Font.Size = 9
    chkMarcadas.Font.Bold = True
    chkMarcadas.BackStyle = fmBackStyleTransparent
    chkMarcadas.ControlTipText = "Mostra só as obras marcadas para analisar"

    Set lblMcMedida = NovoLabel(Me, "lblMcMedida", "", -3000, 16, 10, 26, 17, True, vbWhite)
    lblMcMedida.Font.Name = mFonteTitulo
    lblMcMedida.Font.Bold = mTituloNegrito
    lblMcMedida.WordWrap = False
    lblMcMedida.AutoSize = True

    Set lblMcFundo = NovoLabel(Me, "lblMcFundo", "", LARG - 404, 5, 396, 51)
    lblMcFundo.BackStyle = fmBackStyleOpaque
    lblMcFundo.BackColor = MC_NAVY2
    Set lblMcBarra = NovoLabel(Me, "lblMcBarra", "", LARG - 404, 5, 4, 51)
    lblMcBarra.BackStyle = fmBackStyleOpaque
    lblMcBarra.BackColor = MC_AMARELO
    Set lblMcTit = NovoLabel(Me, "lblMcTit", "", LARG - 392, 8, 250, 14, 9, True, MC_AMARELO)
    lblMcTit.WordWrap = False
    Set lblMcPrazo = NovoLabel(Me, "lblMcPrazo", "", LARG - 140, 8, 128, 14, 8, True, COR_ESCURO)
    lblMcPrazo.BackStyle = fmBackStyleOpaque
    lblMcPrazo.BackColor = MC_AMARELO
    lblMcPrazo.TextAlign = fmTextAlignCenter
    lblMcPrazo.WordWrap = False
    Set lblMcTexto = NovoLabel(Me, "lblMcTexto", "", LARG - 392, 23, 380, 15, 9.5, False, vbWhite)
    lblMcTexto.WordWrap = False
    Set lblMcPor = NovoLabel(Me, "lblMcPor", "", LARG - 392, 39, 380, 13, 8, False, MC_TEXTO_CLARO)
    lblMcPor.WordWrap = False
    Set lblMcClique = NovoLabel(Me, "lblMcClique", "", LARG - 404, 5, 396, 51)
    lblMcClique.ControlTipText = "Clique para ver todas as obras marcadas"
    Set mMcIcone = New cMarcaIcone
    mMcIcone.Criar Me, "icoMarca", 0, 20, 1.2, MC_AMARELO
    mMcIcone.Mostrar False
    MostrarCardMarcador False
End Sub

'Chamado ao fim de CarregarCampos
Private Sub ExtrasAtualizar()
    FerAtualizarObra
    AtualizarMarcadorPainel
End Sub

'-------------------------------------------------------------------- marcador
Private Function MarcasPainel() As cMarcadores
    If mMarcas Is Nothing Then Set mMarcas = New cMarcadores
    Set MarcasPainel = mMarcas
End Function

Private Sub btnMarcar_Click()
    If mObra Is Nothing Then Exit Sub
    If AbrirMarcador(mObra.CodObra, mObra.nome) <> 0 Then
        MarcasPainel().Recarregar
        If chkMarcadas.value = True Then FiltrarLista
        AtualizarMarcadorPainel
    End If
End Sub

Private Sub chkMarcadas_Click()
    If mCarregando Then Exit Sub
    MarcasPainel().Recarregar
    FiltrarLista
    If chkMarcadas.value = True And mQtd > 0 Then
        If PosicaoNaLista(mAtual) = 0 Then
            If ConfirmarPendencias() Then CarregarObra mMapa(1)
        End If
    End If
End Sub

'Clique no cartão do topo: lista das obras marcadas
Private Sub lblMcClique_Click()
    Dim cod As String
    Dim idx As Long
    cod = AbrirListaMarcadas()
    MarcasPainel().Recarregar
    If chkMarcadas.value = True Then FiltrarLista
    If Len(cod) > 0 Then
        idx = MarcasPainel().LinhaNoFollowUp(cod)
        If idx > 0 Then
            IrParaLinha idx
        Else
            AbrirObraMarcada cod
        End If
    End If
    AtualizarMarcadorPainel
End Sub

Private Sub MostrarCardMarcador(ByVal ligado As Boolean)
    lblMcFundo.Visible = ligado
    lblMcBarra.Visible = ligado
    lblMcTit.Visible = ligado
    lblMcPrazo.Visible = ligado
    lblMcTexto.Visible = ligado
    lblMcPor.Visible = ligado
    lblMcClique.Visible = ligado
    If ligado Then lblMcClique.ZOrder 0
End Sub

Private Sub AtualizarMarcadorPainel()
    Dim texto As String
    Dim revisar As Variant
    Dim soEu As Boolean
    Dim por As String
    Dim em As Variant
    Dim total As Long
    Dim hoje As Long
    Dim atrasadas As Long
    Dim marcada As Boolean
    Dim atrasada As Boolean
    Dim s As String

    On Error GoTo Fim
    If lblMcFundo Is Nothing Then Exit Sub
    If Not mObra Is Nothing Then marcada = MarcasPainel().Ler(mObra.CodObra, texto, revisar, soEu, por, em)
    atrasada = (TipoPrazo(revisar) = PZ_ATRASADA)

    If marcada Then
        lblMcMedida.Font.Size = lblNome.Font.Size
        lblMcMedida.Caption = lblNome.Caption
        mMcIcone.Mover lblNome.Left + lblMcMedida.Width + 8, lblNome.Top + 5
        If atrasada Then
            mMcIcone.Cor MC_CORAL
        Else
            mMcIcone.Cor MC_AMARELO
        End If
        mMcIcone.Dica "Marcada para analisar: " & Left$(texto, 200)
        mMcIcone.Frente
        mMcIcone.Mostrar True
        btnMarcar.Caption = "EDITAR MARCADOR"
    Else
        mMcIcone.Mostrar False
        btnMarcar.Caption = "MARCAR P/ ANALISAR"
    End If

    MarcasPainel().Resumo total, hoje, atrasadas
    If marcada Then
        lblMcTit.Caption = "MARCADA PARA ANALISAR"
        lblMcPrazo.Caption = TextoPrazo(revisar)
        If atrasada Then
            lblMcPrazo.BackColor = MC_CORAL
            lblMcPrazo.ForeColor = vbWhite
        Else
            lblMcPrazo.BackColor = MC_AMARELO
            lblMcPrazo.ForeColor = COR_ESCURO
        End If
        lblMcTexto.Caption = ChrW(8220) & Replace(Replace(texto, vbCr, " "), vbLf, " ") & ChrW(8221)
        lblMcTexto.ControlTipText = Left$(texto, 250)
        s = "marcada por " & por
        If IsDate(em) Then s = s & " em " & Format$(em, "dd\/mm\/yyyy hh:mm")
        If soEu Then s = s & "  " & ChrW(183) & "  só você vê"
        s = s & "   " & ChrW(183) & "   clique para ver as " & total & " marcada(s)"
        lblMcPor.Caption = s
        MostrarCardMarcador True
    ElseIf total > 0 Then
        lblMcTit.Caption = "OBRAS MARCADAS PARA ANALISAR: " & total
        If atrasadas > 0 Then
            lblMcPrazo.Caption = atrasadas & " atrasada(s)"
            lblMcPrazo.BackColor = MC_CORAL
            lblMcPrazo.ForeColor = vbWhite
        ElseIf hoje > 0 Then
            lblMcPrazo.Caption = hoje & " para hoje"
            lblMcPrazo.BackColor = MC_AMARELO
            lblMcPrazo.ForeColor = COR_ESCURO
        Else
            lblMcPrazo.Caption = "nenhuma para hoje"
            lblMcPrazo.BackColor = MC_CINZA_CLARO
            lblMcPrazo.ForeColor = COR_ESCURO
        End If
        lblMcTexto.Caption = "Esta obra não está marcada."
        lblMcTexto.ControlTipText = ""
        lblMcPor.Caption = "Clique aqui para ver a lista e abrir uma obra marcada."
        MostrarCardMarcador True
    Else
        MostrarCardMarcador False
    End If
Fim:
End Sub

'-------------------------------------------------------------------- feriados
Private Sub FerPrepararCampos()
    Dim i As Long
    Dim ev As cFerCampo
    Dim t As MSForms.label

    FeriadosRecarregar
    Set mFerCampos = New Collection
    mFerQtd = mQtdCampos
    If mFerQtd = 0 Then Exit Sub
    ReDim mFerTag(1 To mFerQtd)
    For i = 1 To mFerQtd
        Set t = fraInfo.Controls.Add("Forms.Label.1", "lblFerTag" & i, True)
        t.Left = mBoxEd(i).Left + mBoxEd(i).Width - 58
        t.Top = mBoxEd(i).Top + 4
        t.Width = 54
        t.Height = 13
        t.Font.Name = mFonte
        t.Font.Size = 7
        t.Font.Bold = True
        t.TextAlign = fmTextAlignCenter
        t.BackStyle = fmBackStyleOpaque
        t.Caption = ""
        t.Visible = False
        Set mFerTag(i) = t
        Set ev = New cFerCampo
        ev.Conectar mBoxEd(i), i, Me
        mFerCampos.Add ev
    Next i
End Sub

Private Sub FerDesligar()
    Dim ev As cFerCampo
    On Error Resume Next
    If Not mFerCampos Is Nothing Then
        For Each ev In mFerCampos
            ev.Desconectar
        Next ev
    End If
    Set mFerCampos = Nothing
    If Not mCal Is Nothing Then mCal.Desligar
    Set mCal = Nothing
End Sub

'Texto mostrado no campo i (caixa de edição ou de consulta)
Private Function FerTexto(ByVal i As Long) As String
    If mEditavel(i) Then
        FerTexto = mBoxEd(i).Text
    Else
        FerTexto = mBoxRo(i).Text
    End If
End Function

Private Function FerData(ByVal i As Long, ByRef d As Date) As Boolean
    If i < 1 Or i > mFerQtd Then Exit Function
    If mCampoTipo(i) <> "D" Then Exit Function
    FerData = LerDataFer(FerTexto(i), d)
End Function

Private Function FerNome(ByVal i As Long) As String
    FerNome = mCampoRot(i)
End Function

Private Function FerUsar(ByVal d As Date) As String
    FerUsar = "Usar " & Format$(d, "dd\/mm") & " (" & DiaSemanaCurto(d) & ")"
End Function

Private Function FerAlterado(ByVal i As Long) As Boolean
    If mEditavel(i) Then FerAlterado = (mBoxEd(i).Text <> mOrig(i))
End Function

Private Function FerDatasDaObra() As Collection
    Dim col As New Collection
    Dim i As Long
    Dim d As Date
    For i = 1 To mFerQtd
        If FerData(i, d) Then col.Add Array(d, FerNome(i))
    Next i
    Set FerDatasDaObra = col
End Function

Private Sub FerBorda(ByVal i As Long, ByVal cor As Long)
    If cor < 0 Then
        mBoxEd(i).BorderColor = COR_AMARELO
        mBoxRo(i).BorderColor = COR_BORDA
    Else
        mBoxEd(i).BorderColor = cor
        mBoxRo(i).BorderColor = cor
    End If
End Sub

'Etiqueta e borda do campo i
Private Sub FerMarcar(ByVal i As Long, ByVal info As Long, ByVal alterado As Boolean)
    If info = DIA_UTIL Then
        mFerTag(i).Visible = False
        FerBorda i, -1
        Exit Sub
    End If
    mFerTag(i).Caption = EtiquetaDia(info)
    If Bloqueia(info) And alterado Then
        mFerTag(i).BackColor = FR_CORAL
        mFerTag(i).ForeColor = vbWhite
        FerBorda i, FR_CORAL
        mFerTag(i).ControlTipText = "Data alterada em dia não útil: não será aceita ao salvar"
    Else
        mFerTag(i).BackColor = FR_AMARELO
        mFerTag(i).ForeColor = COR_ESCURO
        FerBorda i, -1
        If Not mEditavel(i) And Bloqueia(info) Then
            mFerTag(i).ControlTipText = "Vem do Cronograma de Obras: ajuste a data lá"
        ElseIf Bloqueia(info) Then
            mFerTag(i).ControlTipText = "Data já gravada em dia não útil: ajuste quando puder"
        Else
            mFerTag(i).ControlTipText = "Ponto facultativo: a data é aceita"
        End If
    End If
    mFerTag(i).Visible = True
    mFerTag(i).ZOrder 0
End Sub

'Confere o campo i; digitando = mostra o mês no calendário e o aviso
Private Sub FerAvaliar(ByVal i As Long, ByVal digitando As Boolean)
    Dim d As Date
    Dim info As Long
    Dim nome As String
    Dim abr As String
    Dim alterado As Boolean
    Dim txt As String

    If i < 1 Or i > mFerQtd Then Exit Sub
    If Not FerData(i, d) Then
        mFerTag(i).Visible = False
        FerBorda i, -1
        Exit Sub
    End If
    info = InfoDia(d, mFerUF, mFerCid, nome, abr)
    alterado = FerAlterado(i)
    FerMarcar i, info, alterado
    If Not digitando Then Exit Sub

    mFerIdx = i
    mFerUltimo = i
    mCal.Datas FerDatasDaObra()
    mCal.Mostrar Year(d), Month(d), d
    txt = FerNome(i) & " em " & Format$(d, "dd\/mm\/yyyy") & " cai em " & MotivoDia(info, nome, abr) & "."
    If Bloqueia(info) Then
        mFerSug1 = DiaUtilAnterior(d, mFerUF, mFerCid)
        mFerSug2 = DiaUtilSeguinte(d, mFerUF, mFerCid)
        If alterado Then
            mCal.Aviso txt & " A data não pode cair em fim de semana nem feriado.", 1, FerUsar(mFerSug1), FerUsar(mFerSug2)
        Else
            mCal.Aviso txt & " Já estava gravada assim; ajuste quando puder.", 2, FerUsar(mFerSug1), FerUsar(mFerSug2)
        End If
    ElseIf info = DIA_FACULTATIVO Then
        mCal.Aviso txt & " A data é aceita; confira se haverá expediente.", 2
    Else
        mCal.Aviso FerNome(i) & ": " & Format$(d, "dd\/mm\/yyyy") & " (" & DiaSemanaCurto(d) & ") é dia útil.", 3
    End If
End Sub

Public Sub FerCampoMudou(ByVal i As Long)
    If mCarregando Then Exit Sub
    If mCal Is Nothing Then Exit Sub
    If Not mEditavel(i) Then Exit Sub
    FerAvaliar i, True
End Sub

Public Sub FerCampoFoco(ByVal i As Long)
    Dim d As Date
    If mCal Is Nothing Then Exit Sub
    If i = mFerUltimo Then Exit Sub
    mFerUltimo = i
    If FerData(i, d) Then
        mCal.Datas FerDatasDaObra()
        mCal.Mostrar Year(d), Month(d), d
    End If
End Sub

'Campo de data editável com o cursor (0 = nenhum)
Private Function FerCampoAtivo() As Long
    Dim c As Object
    Dim nm As String
    Dim i As Long
    On Error Resume Next
    Set c = Me.ActiveControl
    If Not c Is Nothing Then
        If TypeName(c) = "Frame" Then Set c = c.ActiveControl
    End If
    If Not c Is Nothing Then nm = c.Name
    On Error GoTo 0
    If Len(nm) > 0 Then
        For i = 1 To mFerQtd
            If mEditavel(i) And mBoxEd(i).Name = nm Then
                FerCampoAtivo = i
                Exit Function
            End If
        Next i
    End If
    If mFerUltimo >= 1 And mFerUltimo <= mFerQtd Then
        If mEditavel(mFerUltimo) Then FerCampoAtivo = mFerUltimo
    End If
End Function

Public Sub FerClique(ByVal chave As String)
    Dim d As Date
    Dim i As Long
    Dim info As Long
    Dim nome As String
    Dim abr As String

    If mCal Is Nothing Then Exit Sub
    Select Case chave
        Case "cal:prev"
            mCal.Mover -1
        Case "cal:next"
            mCal.Mover 1
        Case "aviso:1", "aviso:2"
            If mFerIdx < 1 Or mFerIdx > mFerQtd Then Exit Sub
            If Not mEditavel(mFerIdx) Then
                mCal.Aviso FerNome(mFerIdx) & " vem do Cronograma de Obras: ajuste a data lá e atualize os dados.", 2
                Exit Sub
            End If
            If chave = "aviso:1" Then
                d = mFerSug1
            Else
                d = mFerSug2
            End If
            If d = 0 Then Exit Sub
            mBoxEd(mFerIdx).Text = Format$(d, "dd\/mm\/yyyy")
        Case Else
            If Left$(chave, 6) <> "cal:c:" Then Exit Sub
            d = mCal.DataDaCelula(CLng(Mid$(chave, 7)))
            If d = 0 Then Exit Sub
            info = InfoDia(d, mFerUF, mFerCid, nome, abr)
            If Bloqueia(info) Then
                mCal.Aviso Format$(d, "dd\/mm\/yyyy") & " é " & MotivoDia(info, nome, abr) & ": escolha um dia útil.", 1
                Exit Sub
            End If
            i = FerCampoAtivo()
            If i = 0 Then
                mCal.Aviso "Clique primeiro num campo de data com borda amarela (por exemplo ALVARÁ) e depois no dia do calendário.", 2
                Exit Sub
            End If
            mBoxEd(i).Text = Format$(d, "dd\/mm\/yyyy")
    End Select
End Sub

Private Sub FerAtualizarObra()
    Dim i As Long
    If mCal Is Nothing Then Exit Sub
    mFerUF = ""
    mFerCid = ""
    If Not mObra Is Nothing Then
        mFerUF = mObra.texto(PO_COL_UF)
        mFerCid = mObra.texto(PO_COL_CIDADE)
    End If
    mCal.Obra mFerUF, mFerCid
    For i = 1 To mFerQtd
        FerAvaliar i, False
    Next i
    mFerIdx = 0
    mFerUltimo = 0
    mFerSug1 = 0
    mFerSug2 = 0
    mCal.Datas FerDatasDaObra()
    mCal.Mostrar Year(Date), Month(Date)
    mCal.Dica ""
End Sub

'Antes de salvar: data digitada em fim de semana ou feriado bloqueia
Private Function FerPodeSalvar() As Boolean
    Dim i As Long
    Dim d As Date
    Dim info As Long
    Dim nome As String
    Dim abr As String
    Dim lista As String
    Dim primeiro As Long

    FerPodeSalvar = True
    If mCal Is Nothing Then Exit Function
    For i = 1 To mFerQtd
        If FerAlterado(i) Then
            If FerData(i, d) Then
                info = InfoDia(d, mFerUF, mFerCid, nome, abr)
                If Bloqueia(info) Then
                    If Len(lista) > 0 Then lista = lista & "; "
                    lista = lista & FerNome(i) & " " & Format$(d, "dd\/mm\/yyyy") & " (" & MotivoDia(info, nome, abr) & ")"
                    If primeiro = 0 Then primeiro = i
                End If
            End If
        End If
    Next i
    If primeiro = 0 Then Exit Function

    FerPodeSalvar = False
    FerAvaliar primeiro, True
    mCal.Aviso "Não salvei: " & lista & ". Ajuste para um dia útil e salve de novo.", 1, FerUsar(mFerSug1), FerUsar(mFerSug2)
End Function

'====================================================================
' NOVA OBRA dentro do painel, conferida na base (mInstaladorNovaObraObras)
'====================================================================
Private Sub btnNova_Click()
    If tb Is Nothing Then Exit Sub
    If Not ConfirmarPendencias() Then Exit Sub
    If fraNova Is Nothing Then
        Set fraNova = Me.Controls.Add("Forms.Frame.1", "fraNova", True)
        fraNova.Left = 8
        fraNova.Top = 66
        fraNova.Width = LARG - 4
        fraNova.Height = ALT - 70
        fraNova.Caption = ""
        fraNova.SpecialEffect = fmSpecialEffectFlat
        fraNova.BorderStyle = fmBorderStyleSingle
        fraNova.BorderColor = COR_BORDA
        fraNova.BackColor = COR_FUNDO
        Set mNova = New cNovaObraPainel
        mNova.Criar fraNova
    Else
        mNova.Limpar
    End If
    fraNova.Visible = True
    fraNova.ZOrder 0
    mNova.Focar
End Sub

'A área NOVA OBRA terminou: mostra a obra incluída
Private Sub mNova_Fechar(ByVal linhaObras As Long)
    Dim idx As Long
    fraNova.Visible = False
    idx = linhaObras
    If idx > 0 And idx <= tb.ListRows.Count Then MostrarObra idx
End Sub

'====================================================================
' OBRA ENTREGUE dentro do painel (instalado por mInstaladorEntregueObras)
'====================================================================
Private Sub btnEntrega_Click()
    If tb Is Nothing Then Exit Sub
    If mObra Is Nothing Then Exit Sub
    If Not ConfirmarPendencias() Then Exit Sub
    If fraEntrega Is Nothing Then
        Set fraEntrega = Me.Controls.Add("Forms.Frame.1", "fraEntrega", True)
        fraEntrega.Left = 8
        fraEntrega.Top = 66
        fraEntrega.Width = LARG - 4
        fraEntrega.Height = ALT - 70
        fraEntrega.Caption = ""
        fraEntrega.SpecialEffect = fmSpecialEffectFlat
        fraEntrega.BorderStyle = fmBorderStyleSingle
        fraEntrega.BorderColor = COR_BORDA
        fraEntrega.BackColor = COR_FUNDO
        Set mEntrega = New cEntregaPainel
        mEntrega.Criar fraEntrega
    End If
    mEntrega.Preparar mObra.CodObra
    fraEntrega.Visible = True
    fraEntrega.ZOrder 0
End Sub

'A área OBRA ENTREGUE terminou: se a obra saiu do follow up, recarrega o painel
Private Sub mEntrega_Fechar(ByVal moveu As Boolean)
    Dim idx As Long
    fraEntrega.Visible = False
    If moveu Then
        idx = mAtual
        If idx > tb.ListRows.Count Then idx = tb.ListRows.Count
        Recarregar idx
    End If
End Sub
