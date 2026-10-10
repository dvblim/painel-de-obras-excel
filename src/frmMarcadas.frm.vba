Option Explicit
'===============================================================================
' frmMarcadas - obras marcadas para analisar (atrasadas primeiro)
' UserForm VAZIO: os controles são criados por código. Sem caixas de mensagem.
' Clique seleciona; duplo clique (ou Abrir obra) abre a obra no painel.
'===============================================================================

Private Const DW As Single = 720
Private Const DH As Single = 424
Private Const ALT_LINHA As Single = 36

Private mMarcas As cMarcadores
Private mCliques As Collection
Private mCliquesLinhas As Collection
Private mFonte As String
Private mFonteTit As String
Private mDados As Variant
Private mQtd As Long
Private mSel As Long
Private mSoMinhas As Boolean
Private mEscolhida As String
Private mFundos() As MSForms.Label

Private fraLista As MSForms.Frame
Private lblTitulo As MSForms.Label
Private lblMinhasF As MSForms.Label
Private lblMinhasT As MSForms.Label
Private lblStatus As MSForms.Label

'===============================================================================
' CICLO DE VIDA
'===============================================================================
Private Sub UserForm_Initialize()
    mFonte = FonteMarcador()
    mFonteTit = FonteTituloMarcador()
    Set mCliques = New Collection
    Set mCliquesLinhas = New Collection
    Set mMarcas = New cMarcadores

    Me.Caption = "Obras marcadas para analisar"
    Me.BackColor = MC_FUNDO
    Me.Width = DW + 12
    Me.Height = DH + 30
    Me.Width = DW + (Me.Width - Me.InsideWidth)
    Me.Height = DH + (Me.Height - Me.InsideHeight)
    MontarTela
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        Me.Hide
        Exit Sub
    End If
    Desligar mCliques
    Desligar mCliquesLinhas
    Set mCliques = Nothing
    Set mCliquesLinhas = Nothing
End Sub

Private Sub Desligar(ByVal col As Collection)
    Dim ev As cMarcaClique
    If col Is Nothing Then Exit Sub
    For Each ev In col
        ev.Desconectar
    Next ev
End Sub

Public Property Get Escolhida() As String
    Escolhida = mEscolhida
End Property

Public Sub Preparar()
    mEscolhida = ""
    mSoMinhas = False
    PintarMinhas
    Carregar
End Sub

'===============================================================================
' TELA
'===============================================================================
Private Function Rot(ByVal pai As Object, ByVal nome As String, ByVal texto As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                     ByVal cor As Long, ByVal tam As Single, ByVal negrito As Boolean, _
                     Optional ByVal fundo As Long = -1, Optional ByVal alinh As Long = 1) As MSForms.Label
    Dim c As MSForms.Label
    Set c = pai.Controls.Add("Forms.Label.1", nome, True)
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

Private Sub Clique(ByVal col As Collection, ByVal rotulo As MSForms.Label, ByVal chave As String)
    Dim ev As cMarcaClique
    Set ev = New cMarcaClique
    ev.Conectar rotulo, chave, Me
    col.Add ev
End Sub

Private Sub Botao(ByVal nome As String, ByVal chave As String, ByVal texto As String, _
                  ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                  ByVal fundo As Long, ByVal cor As Long, ByVal borda As Long, _
                  ByRef rotFundo As MSForms.Label, ByRef rotTexto As MSForms.Label)
    Set rotFundo = Rot(Me, nome & "F", "", l, t, w, h, cor, 8, False, fundo)
    If borda >= 0 Then
        rotFundo.BorderStyle = 1
        rotFundo.BorderColor = borda
    End If
    Set rotTexto = Rot(Me, nome & "T", texto, l, t + (h - 12) / 2 - 0.5, w, 12, cor, 8.5, True, -1, 2)
    Clique mCliques, rotFundo, chave
    Clique mCliques, rotTexto, chave
End Sub

Private Sub Legenda(ByVal nome As String, ByVal texto As String, ByVal x As Single, ByVal cor As Long)
    Dim ic As cMarcaIcone
    Set ic = New cMarcaIcone
    ic.Criar Me, nome, x, 357, 0.7, cor
    Rot Me, nome & "T", texto, x + 11, 356, 110, 12, MC_CINZA, 8, False
End Sub

Private Sub MontarTela()
    Dim ic As cMarcaIcone
    Dim f As MSForms.Label
    Dim t As MSForms.Label

    Rot Me, "lblFaixa", "", 0, 0, DW, 30, MC_BRANCO, 8, False, MC_NAVY
    Rot Me, "lblFaixaAm", "", 0, 30, DW, 3, MC_BRANCO, 8, False, MC_AMARELO
    Set ic = New cMarcaIcone
    ic.Criar Me, "icoTopo", 14, 8, 1, MC_AMARELO
    Set lblTitulo = Rot(Me, "lblTitulo", "OBRAS MARCADAS", 32, 8, 400, 16, MC_BRANCO, 10.5, False)
    lblTitulo.Font.Name = mFonteTit
    Rot Me, "lblDica", "clique seleciona  " & ChrW(183) & "  duplo clique abre a obra", DW - 260, 10, 248, 12, MC_TEXTO_CLARO, 8, False, -1, 3

    Rot Me, "lblC1", "OBRA", 40, 42, 180, 12, MC_CINZA, 7.5, True
    Rot Me, "lblC2", "O QUE ANALISAR", 232, 42, 260, 12, MC_CINZA, 7.5, True
    Rot Me, "lblC3", "REVISAR", 500, 42, 100, 12, MC_CINZA, 7.5, True
    Rot Me, "lblC4", "MARCADA POR", 600, 42, 100, 12, MC_CINZA, 7.5, True

    Set fraLista = Me.Controls.Add("Forms.Frame.1", "fraLista", True)
    fraLista.Left = 10
    fraLista.Top = 56
    fraLista.Width = DW - 20
    fraLista.Height = 292
    fraLista.Caption = ""
    fraLista.SpecialEffect = 0
    fraLista.BorderStyle = 1
    fraLista.BorderColor = MC_BORDA
    fraLista.BackColor = MC_BRANCO
    fraLista.ScrollBars = 2
    fraLista.KeepScrollBarsVisible = 0

    Legenda "legA", "atrasada", 16, MC_CORAL
    Legenda "legH", "revisar hoje", 86, MC_AMARELO
    Legenda "legF", "futura ou sem data", 170, MC_CINZA_CLARO

    Botao "btnAbrir", "abrir", "Abrir obra", 10, 376, 170, 26, MC_NAVY, MC_BRANCO, -1, f, t
    Botao "btnConcluir", "concluir", "Concluir análise", 188, 376, 150, 26, MC_BRANCO, MC_CORAL, MC_CORAL, f, t
    Botao "btnMinhas", "minhas", "Só as minhas", 346, 376, 130, 26, MC_BRANCO, MC_NAVY, MC_BORDA, lblMinhasF, lblMinhasT
    Botao "btnFechar", "fechar", "Fechar", DW - 110, 376, 100, 26, MC_BRANCO, MC_NAVY, MC_BORDA, f, t
    Set lblStatus = Rot(Me, "lblStatus", "", 12, 406, DW - 24, 14, MC_CINZA, 8, False)
End Sub

Private Sub PintarMinhas()
    If mSoMinhas Then
        lblMinhasF.BackColor = MC_AZUL
        lblMinhasF.BorderColor = MC_AZUL
        lblMinhasT.ForeColor = MC_BRANCO
    Else
        lblMinhasF.BackColor = MC_BRANCO
        lblMinhasF.BorderColor = MC_BORDA
        lblMinhasT.ForeColor = MC_NAVY
    End If
End Sub

'===============================================================================
' LISTA
'===============================================================================
Private Sub Carregar()
    Dim i As Long
    Dim y As Single
    Dim fundo As Long
    Dim c As MSForms.Label
    Dim ic As cMarcaIcone
    Dim chave As String
    Dim por As String

    Desligar mCliquesLinhas
    Set mCliquesLinhas = New Collection
    fraLista.Controls.Clear
    mSel = 0

    mDados = mMarcas.Lista(mSoMinhas)
    If IsArray(mDados) Then
        mQtd = UBound(mDados, 1) + 1
    Else
        mQtd = 0
    End If
    lblTitulo.Caption = "OBRAS MARCADAS (" & mQtd & ")"

    If mQtd = 0 Then
        Rot fraLista, "lblVazio", IIf(mSoMinhas, "Você não tem obras marcadas.", "Nenhuma obra marcada para analisar."), 12, 14, 600, 14, MC_CINZA, 9, False
        fraLista.ScrollHeight = 0
        Status "", 0
        Exit Sub
    End If

    ReDim mFundos(1 To mQtd)
    For i = 1 To mQtd
        y = (i - 1) * ALT_LINHA
        chave = "linha:" & i
        fundo = FundoDaLinha(i)
        Set mFundos(i) = Rot(fraLista, "lnF" & i, "", 0, y, DW - 40, ALT_LINHA, MC_NAVY, 8, False, fundo)
        Clique mCliquesLinhas, mFundos(i), chave
        Set c = Rot(fraLista, "lnS" & i, "", 0, y + ALT_LINHA - 1, DW - 40, 1, MC_NAVY, 8, False, MC_BORDA)

        Set ic = New cMarcaIcone
        ic.Criar fraLista, "lnI" & i, 10, y + 10, 0.9, CorPrazo(mDados(i - 1, 3))
        Clique mCliquesLinhas, ic.Corpo, chave

        Set c = Rot(fraLista, "lnA" & i, mDados(i - 1, 0), 30, y + 5, 190, 12, MC_NAVY, 9, True)
        Clique mCliquesLinhas, c, chave
        Set c = Rot(fraLista, "lnL" & i, mDados(i - 1, 1), 30, y + 18, 190, 13, MC_NAVY, 8.5, False)
        Clique mCliquesLinhas, c, chave
        Set c = Rot(fraLista, "lnT" & i, mDados(i - 1, 2), 222, y + 4, 264, 28, MC_NAVY, 8.5, False)
        c.WordWrap = True
        c.ControlTipText = Left$(mDados(i - 1, 2), 250)
        Clique mCliquesLinhas, c, chave
        Set c = Rot(fraLista, "lnR" & i, TextoPrazo(mDados(i - 1, 3)), 490, y + 11, 98, 13, MC_NAVY, 8.5, False)
        Select Case TipoPrazo(mDados(i - 1, 3))
            Case PZ_ATRASADA
                c.ForeColor = MC_CORAL
                c.Font.Bold = True
            Case PZ_HOJE
                c.Font.Bold = True
        End Select
        Clique mCliquesLinhas, c, chave
        por = mDados(i - 1, 4)
        If mDados(i - 1, 6) Then por = por & " (só você)"
        Set c = Rot(fraLista, "lnP" & i, por, 590, y + 11, 90, 13, MC_NAVY, 8.5, False)
        If IsDate(mDados(i - 1, 5)) Then c.ControlTipText = "Marcada em " & Format$(mDados(i - 1, 5), "dd/mm/yyyy hh:mm")
        Clique mCliquesLinhas, c, chave
    Next i
    fraLista.ScrollHeight = mQtd * ALT_LINHA
    fraLista.ScrollTop = 0
    Status "", 0
End Sub

Private Function FundoDaLinha(ByVal i As Long) As Long
    If i = mSel Then
        FundoDaLinha = MC_SELECAO
        Exit Function
    End If
    Select Case TipoPrazo(mDados(i - 1, 3))
        Case PZ_ATRASADA
            FundoDaLinha = MC_CORAL_FUNDO
        Case PZ_HOJE
            FundoDaLinha = MC_DESTAQUE
        Case Else
            FundoDaLinha = MC_BRANCO
    End Select
End Function

Private Sub Selecionar(ByVal i As Long)
    Dim antes As Long
    If i < 1 Or i > mQtd Then Exit Sub
    antes = mSel
    mSel = i
    If antes >= 1 And antes <= mQtd Then mFundos(antes).BackColor = FundoDaLinha(antes)
    mFundos(i).BackColor = FundoDaLinha(i)
    Status mDados(i - 1, 0) & "  -  " & mDados(i - 1, 1), 0
End Sub

Private Sub Status(ByVal texto As String, ByVal tipo As Long)
    If tipo = 2 Then
        lblStatus.ForeColor = MC_CORAL
    Else
        lblStatus.ForeColor = MC_CINZA
    End If
    lblStatus.Caption = texto
End Sub

'===============================================================================
' CLIQUES
'===============================================================================
Public Sub TratarClique(ByVal chave As String)
    If Left$(chave, 6) = "linha:" Then
        Selecionar CLng(Mid$(chave, 7))
        Exit Sub
    End If
    Select Case chave
        Case "fechar"
            Me.Hide
        Case "abrir"
            AbrirSelecionada
        Case "minhas"
            mSoMinhas = Not mSoMinhas
            PintarMinhas
            Carregar
        Case "concluir"
            If mSel = 0 Then
                Status "Selecione uma obra na lista.", 2
                Exit Sub
            End If
            If mMarcas.Concluir(CStr(mDados(mSel - 1, 0))) Then
                Carregar
                Status "Análise concluída; a obra saiu da lista.", 0
            Else
                Status mMarcas.UltimoErro, 2
            End If
    End Select
End Sub

Public Sub TratarDuploClique(ByVal chave As String)
    If Left$(chave, 6) = "linha:" Then
        Selecionar CLng(Mid$(chave, 7))
        AbrirSelecionada
    End If
End Sub

Private Sub AbrirSelecionada()
    If mSel = 0 Then
        Status "Selecione uma obra na lista.", 2
        Exit Sub
    End If
    If mMarcas.LinhaNoFollowUp(CStr(mDados(mSel - 1, 0))) = 0 Then
        Status "Esta obra não está mais no follow up (foi entregue ou declinada). Conclua a análise se não precisar mais dela.", 2
        Exit Sub
    End If
    mEscolhida = CStr(mDados(mSel - 1, 0))
    Me.Hide
End Sub
