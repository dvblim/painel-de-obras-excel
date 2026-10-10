Option Explicit
'===============================================================================
' frmMarcador - marcar a obra para analisar depois (marca-página)
' UserForm VAZIO: os controles são criados por código. Sem caixas de mensagem.
'===============================================================================

Private Const DW As Single = 440
Private Const DH As Single = 292

Private WithEvents txtData As MSForms.TextBox

Private mMarcas As cMarcadores
Private mCliques As Collection
Private mFonte As String
Private mFonteTit As String
Private mCarregando As Boolean
Private mCod As String
Private mObra As String
Private mExiste As Boolean
Private mResultado As Long
Private mPreset As Long
Private mDataOrig As String

Private lblTitulo As MSForms.Label
Private txtTexto As MSForms.TextBox
Private mChipF(0 To 3) As MSForms.Label
Private mChipT(0 To 3) As MSForms.Label
Private lblDataInfo As MSForms.Label
Private optTime As MSForms.OptionButton
Private optEu As MSForms.OptionButton
Private lblQuem As MSForms.Label
Private lblSalvarT As MSForms.Label
Private lblConcF As MSForms.Label
Private lblConcT As MSForms.Label
Private lblStatus As MSForms.Label

'===============================================================================
' CICLO DE VIDA
'===============================================================================
Private Sub UserForm_Initialize()
    mFonte = FonteMarcador()
    mFonteTit = FonteTituloMarcador()
    Set mCliques = New Collection
    Set mMarcas = New cMarcadores
    mPreset = 0

    Me.Caption = "Marcar para analisar"
    Me.BackColor = MC_BRANCO
    Me.Width = DW + 12
    Me.Height = DH + 30
    Me.Width = DW + (Me.Width - Me.InsideWidth)
    Me.Height = DH + (Me.Height - Me.InsideHeight)
    mCarregando = True
    MontarTela
    mCarregando = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim ev As cMarcaClique
    If CloseMode = 0 Then
        Cancel = True
        Me.Hide
        Exit Sub
    End If
    If mCliques Is Nothing Then Exit Sub
    For Each ev In mCliques
        ev.Desconectar
    Next ev
    Set mCliques = Nothing
End Sub

Public Property Get Resultado() As Long
    Resultado = mResultado
End Property

Public Sub Preparar(ByVal codObra As String, ByVal obra As String)
    Dim texto As String
    Dim revisar As Variant
    Dim soEu As Boolean
    Dim por As String
    Dim em As Variant

    mCod = Trim$(codObra)
    mObra = Trim$(obra)
    mResultado = 0
    lblTitulo.Caption = "MARCAR PARA ANALISAR  " & ChrW(183) & "  " & UCase$(mObra)
    optEu.Caption = "Só eu (" & mMarcas.Usuario() & ")"

    mCarregando = True
    mExiste = mMarcas.Ler(mCod, texto, revisar, soEu, por, em)
    If mExiste Then
        txtTexto.Text = texto
        If IsDate(revisar) Then
            txtData.Text = Format$(CDate(revisar), "dd/mm/yyyy")
            mPreset = -1
        Else
            txtData.Text = ""
            mPreset = 0
        End If
        optEu.Value = soEu
        optTime.Value = Not soEu
        lblQuem.Caption = "Marcada por " & por & IIf(IsDate(em), " em " & Format$(em, "dd/mm/yyyy hh:mm"), "")
        lblSalvarT.Caption = "Salvar alterações"
    Else
        txtTexto.Text = ""
        txtData.Text = ""
        mPreset = 0
        optTime.Value = True
        lblQuem.Caption = ""
        lblSalvarT.Caption = "Salvar marcador"
    End If
    lblConcF.Visible = mExiste
    lblConcT.Visible = mExiste
    mDataOrig = txtData.Text
    mCarregando = False
    PintarChips
    AtualizarDataInfo
End Sub

'===============================================================================
' TELA
'===============================================================================
Private Function Rot(ByVal nome As String, ByVal texto As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                     ByVal cor As Long, ByVal tam As Single, ByVal negrito As Boolean, _
                     Optional ByVal fundo As Long = -1, Optional ByVal alinh As Long = 1) As MSForms.Label
    Dim c As MSForms.Label
    Set c = Me.Controls.Add("Forms.Label.1", nome, True)
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
    Dim ev As cMarcaClique
    Set ev = New cMarcaClique
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

Private Sub MontarTela()
    Dim c As MSForms.Label
    Dim f As MSForms.Label
    Dim t As MSForms.Label
    Dim ic As cMarcaIcone
    Dim nomes As Variant
    Dim larg As Variant
    Dim i As Long
    Dim x As Single

    Rot "lblFaixa", "", 0, 0, DW, 30, MC_BRANCO, 8, False, MC_NAVY
    Rot "lblFaixaAm", "", 0, 30, DW, 3, MC_BRANCO, 8, False, MC_AMARELO
    Set ic = New cMarcaIcone
    ic.Criar Me, "icoTopo", 14, 8, 1, MC_AMARELO
    Set lblTitulo = Rot("lblTitulo", "", 32, 8, DW - 44, 16, MC_BRANCO, 10, False)
    lblTitulo.Font.Name = mFonteTit

    Rot "lblLab1", "O QUE PRECISA SER ANALISADO", 16, 44, 300, 12, MC_CINZA, 7.5, True
    Set txtTexto = Me.Controls.Add("Forms.TextBox.1", "txtTexto", True)
    txtTexto.Left = 16
    txtTexto.Top = 58
    txtTexto.Width = DW - 32
    txtTexto.Height = 52
    txtTexto.MultiLine = True
    txtTexto.WordWrap = True
    txtTexto.EnterKeyBehavior = True
    txtTexto.MaxLength = 400
    txtTexto.SpecialEffect = 0
    txtTexto.BorderStyle = 1
    txtTexto.BorderColor = MC_BORDA
    txtTexto.BackColor = MC_EDICAO
    txtTexto.ForeColor = MC_NAVY
    txtTexto.Font.Name = mFonte
    txtTexto.Font.Size = 9.5

    Rot "lblLab2", "REVISAR EM", 16, 120, 200, 12, MC_CINZA, 7.5, True
    nomes = Array("Sem data", "Amanhã", "Em 2 dias", "Em 1 semana")
    larg = Array(58, 58, 62, 74)
    x = 16
    For i = 0 To 3
        Botao "chip" & i, "chip:" & i, CStr(nomes(i)), x, 134, CSng(larg(i)), 20, MC_BRANCO, MC_NAVY, MC_BORDA, f, t
        Set mChipF(i) = f
        Set mChipT(i) = t
        x = x + CSng(larg(i)) + 6
    Next i
    Set txtData = Me.Controls.Add("Forms.TextBox.1", "txtData", True)
    txtData.Left = x
    txtData.Top = 134
    txtData.Width = 82
    txtData.Height = 20
    txtData.MaxLength = 10
    txtData.SpecialEffect = 0
    txtData.BorderStyle = 1
    txtData.BorderColor = MC_BORDA
    txtData.BackColor = MC_EDICAO
    txtData.ForeColor = MC_NAVY
    txtData.Font.Name = mFonte
    txtData.Font.Size = 9
    txtData.ControlTipText = "Outra data: dd/mm/aaaa (ou só os números, 15102026)"
    Set lblDataInfo = Rot("lblDataInfo", "", 16, 158, DW - 32, 12, MC_CINZA, 8, False)

    Rot "lblLab3", "QUEM VÊ", 16, 180, 200, 12, MC_CINZA, 7.5, True
    Set optTime = NovaOpcao("optTime", "Todo o time", 16, 193, 110)
    Set optEu = NovaOpcao("optEu", "Só eu", 130, 193, 220)
    Set lblQuem = Rot("lblQuem", "", 16, 216, DW - 32, 12, MC_CINZA, 7.5, False)

    Botao "btnSalvar", "salvar", "Salvar marcador", 16, 236, 196, 26, MC_NAVY, MC_BRANCO, -1, f, lblSalvarT
    Botao "btnCancelar", "cancelar", "Cancelar", 218, 236, 96, 26, MC_BRANCO, MC_NAVY, MC_BORDA, f, t
    Botao "btnConcluir", "concluir", "Concluir análise", 320, 236, 104, 26, MC_BRANCO, MC_CORAL, MC_CORAL, lblConcF, lblConcT
    lblConcF.ControlTipText = "Tira a marcação; fica guardado quem concluiu e quando"
    lblConcT.ControlTipText = lblConcF.ControlTipText
    Set lblStatus = Rot("lblStatus", "", 16, 270, DW - 32, 14, MC_CORAL, 8, False)
End Sub

Private Function NovaOpcao(ByVal nome As String, ByVal texto As String, ByVal l As Single, ByVal t As Single, ByVal w As Single) As MSForms.OptionButton
    Dim o As MSForms.OptionButton
    Set o = Me.Controls.Add("Forms.OptionButton.1", nome, True)
    o.Left = l
    o.Top = t
    o.Width = w
    o.Height = 18
    o.Caption = texto
    o.BackStyle = 0
    o.ForeColor = MC_NAVY
    o.Font.Name = mFonte
    o.Font.Size = 9
    o.GroupName = "quem"
    Set NovaOpcao = o
End Function

Private Sub PintarChips()
    Dim i As Long
    For i = 0 To 3
        If i = mPreset Then
            mChipF(i).BackColor = MC_AZUL
            mChipF(i).BorderColor = MC_AZUL
            mChipT(i).ForeColor = MC_BRANCO
        Else
            mChipF(i).BackColor = MC_BRANCO
            mChipF(i).BorderColor = MC_BORDA
            mChipT(i).ForeColor = MC_NAVY
        End If
    Next i
End Sub

Private Sub AtualizarDataInfo()
    Dim d As Date
    If Len(Trim$(txtData.Text)) = 0 Then
        lblDataInfo.ForeColor = MC_CINZA
        lblDataInfo.Caption = "Sem data de revisão: a obra fica marcada até alguém concluir a análise."
    ElseIf LerData(txtData.Text, d) Then
        lblDataInfo.ForeColor = MC_CINZA
        lblDataInfo.Caption = "= " & Format$(d, "dd/mm/yyyy") & " (" & DiaDaSemana(d) & ")"
        If d < Date Then
            lblDataInfo.ForeColor = MC_CORAL
            lblDataInfo.Caption = lblDataInfo.Caption & "  -  esta data já passou"
        End If
    Else
        lblDataInfo.ForeColor = MC_CORAL
        lblDataInfo.Caption = "Data inválida. Use dd/mm/aaaa."
    End If
End Sub

'===============================================================================
' EVENTOS
'===============================================================================
Private Sub txtData_Change()
    If mCarregando Then Exit Sub
    mPreset = -1
    If Len(Trim$(txtData.Text)) = 0 Then mPreset = 0
    PintarChips
    AtualizarDataInfo
End Sub

Public Sub TratarDuploClique(ByVal chave As String)
End Sub

Public Sub TratarClique(ByVal chave As String)
    Dim n As Long
    If Left$(chave, 5) = "chip:" Then
        n = CLng(Mid$(chave, 6))
        mCarregando = True
        Select Case n
            Case 0
                txtData.Text = ""
            Case 1
                txtData.Text = Format$(Date + 1, "dd/mm/yyyy")
            Case 2
                txtData.Text = Format$(Date + 2, "dd/mm/yyyy")
            Case 3
                txtData.Text = Format$(Date + 7, "dd/mm/yyyy")
        End Select
        mCarregando = False
        mPreset = n
        PintarChips
        AtualizarDataInfo
        Exit Sub
    End If

    Select Case chave
        Case "cancelar"
            Me.Hide
        Case "salvar"
            Salvar
        Case "concluir"
            If mMarcas.Concluir(mCod) Then
                mResultado = 2
                Me.Hide
            Else
                lblStatus.Caption = mMarcas.UltimoErro
            End If
    End Select
End Sub

Private Sub Salvar()
    Dim d As Date
    Dim revisar As Variant
    Dim texto As String

    texto = Trim$(txtTexto.Text)
    If Len(texto) = 0 Then
        lblStatus.Caption = "Escreva o que precisa ser analisado."
        Exit Sub
    End If
    revisar = Empty
    If Len(Trim$(txtData.Text)) > 0 Then
        If Not LerData(txtData.Text, d) Then
            lblStatus.Caption = "A data de revisão não é válida. Use dd/mm/aaaa ou deixe em branco."
            Exit Sub
        End If
        If d < Date And txtData.Text <> mDataOrig Then
            lblStatus.Caption = "A data de revisão já passou. Escolha hoje ou uma data futura."
            Exit Sub
        End If
        revisar = d
    End If
    If mMarcas.Salvar(mCod, mObra, texto, revisar, optEu.Value) Then
        mResultado = 1
        Me.Hide
    Else
        lblStatus.Caption = mMarcas.UltimoErro
    End If
End Sub
