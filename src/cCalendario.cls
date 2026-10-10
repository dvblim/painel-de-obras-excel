Option Explicit
'===============================================================================
' cCalendario - calendário do mês e faixa de avisos do painel
'
' Feriado = vermelho; ponto facultativo = amarelo claro; fim de semana =
' cinza; hoje = contorno escuro; data selecionada = contorno azul (vermelho
' se não for dia útil); datas da obra = pontinho embaixo do dia.
' Cliques: dono.FerClique "cal:prev" | "cal:next" | "cal:c:<n>" | "aviso:1|2"
'===============================================================================

Private Const CEL_W As Single = 26
Private Const CEL_H As Single = 21
Private Const CEL_G As Single = 2

Private mDono As Object
Private mForm As Object
Private mCliques As Collection
Private mFonte As String
Private mFonteTit As String

Private mL As Single
Private mT As Single
Private mW As Single

Private mAno As Long
Private mMes As Long
Private mUF As String
Private mCidade As String
Private mSel As Date
Private mDatas As Collection

Private lblMes As MSForms.Label
Private lblLocal As MSForms.Label
Private lblLista As MSForms.Label
Private mCel(1 To 42) As MSForms.Label
Private mPonto(1 To 42) As MSForms.Label
Private mDataCel(1 To 42) As Date

' faixa de avisos
Private lblFxFundo As MSForms.Label
Private lblFxBarra As MSForms.Label
Private lblFxTexto As MSForms.Label
Private lblFxB1 As MSForms.Label
Private lblFxB2 As MSForms.Label

'-------------------------------------------------------------------------------
' Montagem
'-------------------------------------------------------------------------------
Private Function Rot(ByVal nome As String, ByVal texto As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                     ByVal cor As Long, ByVal tam As Single, ByVal negrito As Boolean, _
                     Optional ByVal fundo As Long = -1, Optional ByVal alinh As Long = 1) As MSForms.Label
    Dim c As MSForms.Label
    Set c = mForm.Controls.Add("Forms.Label.1", nome, True)
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
    Dim ev As cFerClique
    Set ev = New cFerClique
    ev.Conectar rotulo, chave, mDono
    mCliques.Add ev
End Sub

' l, t, w: cartão do calendário (altura 324); fl, ft, fw: faixa de avisos (altura 24)
Public Sub Criar(ByVal dono As Object, ByVal l As Single, ByVal t As Single, ByVal w As Single, _
                 ByVal fl As Single, ByVal ft As Single, ByVal fw As Single)
    Dim c As MSForms.Label
    Dim i As Long
    Dim x0 As Single
    Dim x As Single
    Dim y As Single
    Dim dias As Variant
    Dim leg As Variant
    Dim corL As Variant

    Set mDono = dono
    Set mForm = dono
    Set mCliques = New Collection
    Set mDatas = New Collection
    mFonte = FonteFeriados()
    mFonteTit = FonteTituloFeriados()
    mL = l
    mT = t
    mW = w

    Set c = Rot("calFundo", "", l, t, w, 324, FR_NAVY, 8, False, FR_BRANCO)
    c.BorderStyle = 1
    c.BorderColor = FR_BORDA
    Rot "calTopo", "", l, t, w, 22, FR_BRANCO, 8, False, FR_NAVY
    Rot "calTopoAm", "", l, t + 22, w, 2, FR_BRANCO, 8, False, FR_AMARELO
    Set c = Rot("calAnt", "<", l + 2, t + 3, 22, 16, FR_BRANCO, 10, True, -1, 2)
    c.ControlTipText = "Mês anterior"
    Clique c, "cal:prev"
    Set c = Rot("calProx", ">", l + w - 24, t + 3, 22, 16, FR_BRANCO, 10, True, -1, 2)
    c.ControlTipText = "Próximo mês"
    Clique c, "cal:next"
    Set lblMes = Rot("calMes", "", l + 26, t + 4, w - 52, 15, FR_BRANCO, 9.5, False, -1, 2)
    lblMes.Font.Name = mFonteTit
    Set lblLocal = Rot("calLocal", "", l + 4, t + 27, w - 8, 11, FR_CINZA, 7, False, -1, 2)

    x0 = l + (w - (7 * CEL_W + 6 * CEL_G)) / 2
    dias = Array("D", "S", "T", "Q", "Q", "S", "S")
    For i = 0 To 6
        Rot "calSem" & i, CStr(dias(i)), x0 + i * (CEL_W + CEL_G), t + 40, CEL_W, 11, FR_CINZA, 7, True, -1, 2
    Next i
    For i = 1 To 42
        x = x0 + ((i - 1) Mod 7) * (CEL_W + CEL_G)
        y = t + 53 + ((i - 1) \ 7) * (CEL_H + CEL_G)
        Set mCel(i) = Rot("calCel" & i, "", x, y, CEL_W, CEL_H, FR_NAVY, 8.5, False, FR_BRANCO, 2)
        mCel(i).BorderStyle = 0
        Clique mCel(i), "cal:c:" & i
        Set mPonto(i) = Rot("calPto" & i, "", x + CEL_W / 2 - 2, y + CEL_H - 5, 4, 3, FR_AZUL, 6, False, FR_AZUL)
        mPonto(i).Visible = False
        Clique mPonto(i), "cal:c:" & i
    Next i

    ' legenda
    y = t + 53 + 6 * (CEL_H + CEL_G) + 2
    leg = Array("feriado", "ponto facult.", "fim de semana", "data da obra")
    corL = Array(FR_CORAL, FR_DESTAQUE, FR_FDS, FR_AZUL)
    For i = 0 To 3
        x = l + 8 + (i Mod 2) * 96
        Set c = Rot("calLegC" & i, "", x, y + 3 + (i \ 2) * 12, 7, 7, FR_NAVY, 6, False, CLng(corL(i)))
        If i = 1 Then
            c.BorderStyle = 1
            c.BorderColor = FR_AMARELO
        End If
        Rot "calLegT" & i, CStr(leg(i)), x + 10, y + (i \ 2) * 12, 84, 11, FR_CINZA, 7, False
    Next i

    Set c = Rot("calSep", "", l + 6, y + 27, w - 12, 1, FR_NAVY, 6, False, FR_BORDA)
    Set lblLista = Rot("calLista", "", l + 8, y + 31, w - 16, t + 320 - (y + 31), FR_NAVY, 7.5, False)
    lblLista.WordWrap = True

    ' faixa de avisos
    Set lblFxFundo = Rot("fxFundo", "", fl, ft, fw, 24, FR_NAVY, 8, False, FR_FUNDO)
    lblFxFundo.BorderStyle = 1
    lblFxFundo.BorderColor = FR_BORDA
    Set lblFxBarra = Rot("fxBarra", "", fl, ft, 4, 24, FR_NAVY, 8, False, FR_BORDA)
    Set lblFxTexto = Rot("fxTexto", "", fl + 12, ft + 5, fw - 300, 14, FR_CINZA, 8.5, False)
    Set lblFxB1 = Rot("fxB1", "", fl + fw - 276, ft + 3, 132, 18, FR_NAVY, 8.5, True, FR_BRANCO, 2)
    lblFxB1.BorderStyle = 1
    lblFxB1.BorderColor = FR_BORDA
    Clique lblFxB1, "aviso:1"
    Set lblFxB2 = Rot("fxB2", "", fl + fw - 138, ft + 3, 132, 18, FR_BRANCO, 8.5, True, FR_NAVY, 2)
    Clique lblFxB2, "aviso:2"
    Dica ""
End Sub

Public Sub Desligar()
    Dim ev As cFerClique
    If mCliques Is Nothing Then Exit Sub
    For Each ev In mCliques
        ev.Desconectar
    Next ev
    Set mCliques = Nothing
    Set mDono = Nothing
    Set mForm = Nothing
End Sub

'-------------------------------------------------------------------------------
' Dados
'-------------------------------------------------------------------------------
Public Sub Obra(ByVal uf As String, ByVal cidade As String)
    mUF = uf
    mCidade = cidade
End Sub

' datas da obra: cada item = Array(data, rótulo)
Public Sub Datas(ByVal itens As Collection)
    If itens Is Nothing Then
        Set mDatas = New Collection
    Else
        Set mDatas = itens
    End If
End Sub

Public Property Get DataDaCelula(ByVal n As Long) As Date
    If n >= 1 And n <= 42 Then DataDaCelula = mDataCel(n)
End Property

Public Sub Mover(ByVal passo As Long)
    Dim d As Date
    d = DateSerial(mAno, mMes + passo, 1)
    Mostrar Year(d), Month(d), mSel
End Sub

'-------------------------------------------------------------------------------
' Desenho
'-------------------------------------------------------------------------------
Public Sub Mostrar(ByVal ano As Long, ByVal mes As Long, Optional ByVal selecionada As Date = 0)
    Dim primeiro As Date
    Dim d As Date
    Dim i As Long
    Dim info As Long
    Dim nome As String
    Dim abr As String
    Dim lista As String
    Dim prox As String
    Dim k As Long
    Dim ultimo As Long

    If lblMes Is Nothing Then Exit Sub
    mAno = ano
    mMes = mes
    mSel = selecionada
    lblMes.Caption = NomeMes(mes) & " " & ano
    If Len(mCidade) > 0 Then
        lblLocal.Caption = "feriados de " & UCase$(mCidade) & IIf(Len(mUF) > 0, "-" & UCase$(mUF), "")
    Else
        lblLocal.Caption = "feriados nacionais" & IIf(Len(mUF) > 0, " e de " & UCase$(mUF), "")
    End If

    primeiro = DateSerial(ano, mes, 1)
    ultimo = Day(DateSerial(ano, mes + 1, 0))
    For i = 1 To 42
        d = primeiro + (i - Weekday(primeiro, vbSunday))
        mPonto(i).Visible = False
        mCel(i).BorderStyle = 0
        If Month(d) <> mes Then
            mDataCel(i) = 0
            mCel(i).Caption = ""
            mCel(i).BackColor = FR_BRANCO
            mCel(i).ControlTipText = ""
        Else
            mDataCel(i) = d
            mCel(i).Caption = CStr(Day(d))
            info = InfoDia(d, mUF, mCidade, nome, abr)
            mCel(i).Font.Bold = False
            Select Case info
                Case DIA_FERIADO
                    mCel(i).BackColor = FR_CORAL
                    mCel(i).ForeColor = FR_BRANCO
                    mCel(i).Font.Bold = True
                    mCel(i).ControlTipText = Format$(d, "dd\/mm") & " - " & nome & " (" & abr & ")"
                    lista = lista & Format$(d, "dd\/mm") & " " & DiaSemanaCurto(d) & " " & ChrW(183) & " " & nome & " (" & abr & ")" & vbLf
                Case DIA_FACULTATIVO
                    mCel(i).BackColor = FR_DESTAQUE
                    mCel(i).ForeColor = FR_NAVY
                    mCel(i).BorderStyle = 1
                    mCel(i).BorderColor = FR_AMARELO
                    mCel(i).ControlTipText = Format$(d, "dd\/mm") & " - " & nome & " (ponto facultativo)"
                    lista = lista & Format$(d, "dd\/mm") & " " & DiaSemanaCurto(d) & " " & ChrW(183) & " " & nome & " (ponto facultativo)" & vbLf
                Case DIA_SABADO, DIA_DOMINGO
                    mCel(i).BackColor = FR_FDS
                    mCel(i).ForeColor = FR_FDS_TEXTO
                    mCel(i).ControlTipText = Format$(d, "dd\/mm") & " - fim de semana"
                Case Else
                    mCel(i).BackColor = FR_BRANCO
                    mCel(i).ForeColor = FR_NAVY
                    mCel(i).ControlTipText = Format$(d, "dd\/mm") & " - dia útil"
            End Select
            If d = Date Then
                mCel(i).BorderStyle = 1
                mCel(i).BorderColor = FR_NAVY
                mCel(i).Font.Bold = True
            End If
            If d = selecionada Then
                mCel(i).BorderStyle = 1
                mCel(i).Font.Bold = True
                If Bloqueia(info) Then
                    mCel(i).BorderColor = FR_CORAL
                    If info <> DIA_FERIADO Then
                        mCel(i).BackColor = FR_CORAL_FUNDO
                        mCel(i).ForeColor = FR_CORAL
                    End If
                Else
                    mCel(i).BorderColor = FR_AZUL
                    mCel(i).BackColor = FR_AZUL_FUNDO
                End If
            End If
            MarcarPonto i, d, info
        End If
    Next i

    If Len(lista) = 0 Then lista = "Nenhum feriado neste mês." & vbLf
    prox = ProximoFeriado(DateSerial(ano, mes, ultimo))
    If Len(prox) > 0 Then lista = lista & "próximo: " & prox
    lblLista.Caption = lista
End Sub

Private Sub MarcarPonto(ByVal i As Long, ByVal d As Date, ByVal info As Long)
    Dim it As Variant
    Dim rotulos As String
    If mDatas Is Nothing Then Exit Sub
    For Each it In mDatas
        If CDate(it(0)) = d Then
            If Len(rotulos) > 0 Then rotulos = rotulos & ", "
            rotulos = rotulos & it(1)
        End If
    Next it
    If Len(rotulos) = 0 Then Exit Sub
    mPonto(i).Visible = True
    If info = DIA_FERIADO Then
        mPonto(i).BackColor = FR_BRANCO
    ElseIf Bloqueia(info) Then
        mPonto(i).BackColor = FR_OURO
    Else
        mPonto(i).BackColor = FR_AZUL
    End If
    mCel(i).ControlTipText = mCel(i).ControlTipText & "  |  " & rotulos
    mPonto(i).ControlTipText = mCel(i).ControlTipText
End Sub

' Próximo feriado (bloqueante) depois da data, em até 120 dias.
Private Function ProximoFeriado(ByVal depoisDe As Date) As String
    Dim d As Date
    Dim k As Long
    Dim nome As String
    Dim abr As String
    d = depoisDe
    For k = 1 To 120
        d = d + 1
        If InfoDia(d, mUF, mCidade, nome, abr) = DIA_FERIADO Then
            ProximoFeriado = Format$(d, "dd\/mm") & " " & DiaSemanaCurto(d) & " " & ChrW(183) & " " & nome
            Exit Function
        End If
    Next k
End Function

'-------------------------------------------------------------------------------
' Faixa de avisos
' tipo: 0 dica (cinza), 1 erro (vermelho), 2 atenção (amarelo), 3 ok (azul)
'-------------------------------------------------------------------------------
Public Sub Aviso(ByVal texto As String, ByVal tipo As Long, Optional ByVal botao1 As String = "", Optional ByVal botao2 As String = "")
    If lblFxFundo Is Nothing Then Exit Sub
    Select Case tipo
        Case 1
            lblFxFundo.BackColor = FR_CORAL_FUNDO
            lblFxBarra.BackColor = FR_CORAL
            lblFxTexto.ForeColor = FR_NAVY
        Case 2
            lblFxFundo.BackColor = FR_DESTAQUE
            lblFxBarra.BackColor = FR_AMARELO
            lblFxTexto.ForeColor = FR_NAVY
        Case 3
            lblFxFundo.BackColor = FR_AZUL_FUNDO
            lblFxBarra.BackColor = FR_AZUL
            lblFxTexto.ForeColor = FR_NAVY
        Case Else
            lblFxFundo.BackColor = FR_FUNDO
            lblFxBarra.BackColor = FR_BORDA
            lblFxTexto.ForeColor = FR_CINZA
    End Select
    lblFxTexto.Caption = texto
    lblFxTexto.ControlTipText = texto
    lblFxB1.Caption = botao1
    lblFxB2.Caption = botao2
    lblFxB1.Visible = (Len(botao1) > 0)
    lblFxB2.Visible = (Len(botao2) > 0)
    If Len(botao1) > 0 Or Len(botao2) > 0 Then
        lblFxTexto.Width = lblFxFundo.Width - 300
    Else
        lblFxTexto.Width = lblFxFundo.Width - 20
    End If
End Sub

Public Sub Dica(ByVal texto As String)
    If Len(texto) = 0 Then texto = "Selecione um campo do cronograma e digite a data, ou clique num dia útil do calendário."
    Aviso texto, 0
End Sub
