Option Explicit
'===============================================================================
' mFeriados - feriados nacionais, estaduais e municipais e dias úteis
'
' NACIONAIS: calculados aqui (fixos + Sexta-feira Santa, pela Páscoa).
' PONTOS FACULTATIVOS NACIONAIS: Carnaval (segunda e terça), Quarta-feira de
'   Cinzas e Corpus Christi; aparecem no calendário mas não bloqueiam a data.
' ESTADUAIS E MUNICIPAIS: aba FERIADOS (tabela tb_Feriados), que o time pode
'   completar. Coluna QUANDO aceita:
'     dd/mm             todo ano
'     dd/mm/aaaa        só naquele ano
'     PASCOA+n / PASCOA-n, SEXTA-FEIRA SANTA, CORPUS CHRISTI, CARNAVAL,
'     CARNAVAL TERCA, QUARTA-FEIRA DE CINZAS
'   ABRANGÊNCIA: ESTADUAL, MUNICIPAL (bloqueiam) ou PONTO FACULTATIVO.
'   UF vazia = todas; CIDADE vazia = o estado todo.
'===============================================================================

Public Const FR_NAVY As Long = 3877150
Public Const FR_AZUL As Long = 8950797
Public Const FR_CORAL As Long = 2500316
Public Const FR_AMARELO As Long = 2408443
Public Const FR_CINZA As Long = 9139300
Public Const FR_FUNDO As Long = 16579320
Public Const FR_BORDA As Long = 15788258
Public Const FR_BRANCO As Long = 16777215
Public Const FR_DESTAQUE As Long = 15465471
Public Const FR_CORAL_FUNDO As Long = 14869246
Public Const FR_FDS As Long = 16381425
Public Const FR_FDS_TEXTO As Long = 12100500
Public Const FR_AZUL_FUNDO As Long = 16449008
Public Const FR_OURO As Long = 423897

' resultado de InfoDia
Public Const DIA_UTIL As Long = 0
Public Const DIA_SABADO As Long = 1
Public Const DIA_DOMINGO As Long = 2
Public Const DIA_FERIADO As Long = 3
Public Const DIA_FACULTATIVO As Long = 4

Private Const T_FER As String = "tb_Feriados"

Private mCarregado As Boolean
Private mQtd As Long
Private mQuando() As String
Private mNome() As String
Private mAbr() As String
Private mUF() As String
Private mCid() As String
Private mFonteTexto As String
Private mFonteTitulo As String

'-------------------------------------------------------------------------------
' Tabela FERIADOS
'-------------------------------------------------------------------------------
Public Sub FeriadosRecarregar()
    mCarregado = False
End Sub

Private Function TabelaFeriados() As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        On Error Resume Next
        Set lo = ws.ListObjects(T_FER)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set TabelaFeriados = lo
            Exit Function
        End If
    Next ws
End Function

Private Sub Carregar()
    Dim tb As ListObject
    Dim v As Variant
    Dim i As Long
    Dim cQ As Long
    Dim cN As Long
    Dim cA As Long
    Dim cU As Long
    Dim cC As Long

    mCarregado = True
    mQtd = 0
    Set tb = TabelaFeriados()
    If tb Is Nothing Then Exit Sub
    If tb.DataBodyRange Is Nothing Then Exit Sub
    cQ = ColFer(tb, "QUANDO")
    cN = ColFer(tb, "NOME")
    cA = ColFer(tb, "ABRANGÊNCIA")
    cU = ColFer(tb, "UF")
    cC = ColFer(tb, "CIDADE")
    If cQ = 0 Or cN = 0 Or cA = 0 Then Exit Sub
    v = tb.DataBodyRange.Value
    ReDim mQuando(1 To UBound(v, 1))
    ReDim mNome(1 To UBound(v, 1))
    ReDim mAbr(1 To UBound(v, 1))
    ReDim mUF(1 To UBound(v, 1))
    ReDim mCid(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        If Len(TextoCel(v(i, cQ))) > 0 Then
            mQtd = mQtd + 1
            If VarType(v(i, cQ)) = vbDate Then
                mQuando(mQtd) = Format$(v(i, cQ), "dd\/mm\/yyyy")
            Else
                mQuando(mQtd) = UCase$(SemAcento(TextoCel(v(i, cQ))))
            End If
            mNome(mQtd) = TextoCel(v(i, cN))
            mAbr(mQtd) = UCase$(SemAcento(TextoCel(v(i, cA))))
            If cU > 0 Then mUF(mQtd) = UCase$(TextoCel(v(i, cU)))
            If cC > 0 Then mCid(mQtd) = NormCidade(TextoCel(v(i, cC)))
        End If
    Next i
End Sub

Private Function ColFer(ByVal tb As ListObject, ByVal nome As String) As Long
    Dim i As Long
    For i = 1 To tb.ListColumns.Count
        If UCase$(SemAcento(Trim$(tb.ListColumns(i).Name))) = UCase$(SemAcento(nome)) Then
            ColFer = i
            Exit Function
        End If
    Next i
End Function

Private Function TextoCel(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    TextoCel = Trim$(Replace(CStr(v), Chr$(160), " "))
End Function

Public Function SemAcento(ByVal s As String) As String
    Dim de As String
    Dim para As String
    Dim i As Long
    de = "ÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇáàâãäéèêëíìîïóòôõöúùûüç"
    para = "AAAAAEEEEIIIIOOOOOUUUUCaaaaaeeeeiiiiooooouuuuc"
    For i = 1 To Len(de)
        s = Replace(s, Mid$(de, i, 1), Mid$(para, i, 1), , , vbBinaryCompare)
    Next i
    SemAcento = s
End Function

Public Function NormCidade(ByVal s As String) As String
    s = UCase$(SemAcento(Trim$(Replace(s, Chr$(160), " "))))
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    NormCidade = s
End Function

'-------------------------------------------------------------------------------
' Datas móveis
'-------------------------------------------------------------------------------
Public Function Pascoa(ByVal ano As Long) As Date
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim e As Long
    Dim f As Long
    Dim g As Long
    Dim h As Long
    Dim i As Long
    Dim k As Long
    Dim l As Long
    Dim m As Long
    Dim mes As Long
    Dim dia As Long
    a = ano Mod 19
    b = ano \ 100
    c = ano Mod 100
    d = b \ 4
    e = b Mod 4
    f = (b + 8) \ 25
    g = (b - f + 1) \ 3
    h = (19 * a + b - d - g + 15) Mod 30
    i = c \ 4
    k = c Mod 4
    l = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = (a + 11 * h + 22 * l) \ 451
    mes = (h + l - 7 * m + 114) \ 31
    dia = ((h + l - 7 * m + 114) Mod 31) + 1
    Pascoa = DateSerial(ano, mes, dia)
End Function

' A regra QUANDO vale para a data d?
Private Function Bate(ByVal quando As String, ByVal d As Date) As Boolean
    Dim p As Date
    Dim n As Long
    If quando Like "##/##/####" Then
        Bate = (Format$(d, "dd\/mm\/yyyy") = quando)
    ElseIf quando Like "##/##" Then
        Bate = (Format$(d, "dd\/mm") = quando)
    Else
        p = Pascoa(Year(d))
        Select Case quando
            Case "SEXTA-FEIRA SANTA"
                Bate = (d = p - 2)
            Case "CORPUS CHRISTI"
                Bate = (d = p + 60)
            Case "CARNAVAL"
                Bate = (d = p - 48 Or d = p - 47)
            Case "CARNAVAL TERCA"
                Bate = (d = p - 47)
            Case "QUARTA-FEIRA DE CINZAS"
                Bate = (d = p - 46)
            Case Else
                If quando Like "PASCOA[+-]*" Then
                    If IsNumeric(Mid$(quando, 8)) Then
                        n = CLng(Mid$(quando, 8))
                        If Mid$(quando, 7, 1) = "-" Then n = -n
                        Bate = (d = p + n)
                    End If
                End If
        End Select
    End If
End Function

'-------------------------------------------------------------------------------
' Consulta
'-------------------------------------------------------------------------------
' Nacionais (bloqueiam): devolve o nome ("" se não for).
Private Function Nacional(ByVal d As Date) As String
    Select Case Format$(d, "dd\/mm")
        Case "01/01"
            Nacional = "Confraternização Universal"
        Case "21/04"
            Nacional = "Tiradentes"
        Case "01/05"
            Nacional = "Dia do Trabalho"
        Case "07/09"
            Nacional = "Independência do Brasil"
        Case "12/10"
            Nacional = "Nossa Senhora Aparecida"
        Case "02/11"
            Nacional = "Finados"
        Case "15/11"
            Nacional = "Proclamação da República"
        Case "20/11"
            Nacional = "Dia da Consciência Negra"
        Case "25/12"
            Nacional = "Natal"
    End Select
    If Len(Nacional) = 0 Then
        If d = Pascoa(Year(d)) - 2 Then Nacional = "Sexta-feira Santa"
    End If
End Function

Private Function FacultativoNacional(ByVal d As Date) As String
    Dim p As Date
    p = Pascoa(Year(d))
    If d = p - 48 Or d = p - 47 Then
        FacultativoNacional = "Carnaval"
    ElseIf d = p - 46 Then
        FacultativoNacional = "Quarta-feira de Cinzas"
    ElseIf d = p + 60 Then
        FacultativoNacional = "Corpus Christi"
    End If
End Function

' O que é o dia d na cidade: DIA_UTIL, DIA_SABADO, DIA_DOMINGO, DIA_FERIADO
' ou DIA_FACULTATIVO. nome e abrangencia descrevem o feriado (quando houver).
Public Function InfoDia(ByVal d As Date, ByVal uf As String, ByVal cidade As String, _
                        ByRef nome As String, ByRef abrangencia As String) As Long
    Dim i As Long
    Dim fac As String
    Dim facAbr As String

    nome = ""
    abrangencia = ""
    d = DateSerial(Year(d), Month(d), Day(d))
    If Not mCarregado Then Carregar
    uf = UCase$(Trim$(uf))
    cidade = NormCidade(cidade)

    nome = Nacional(d)
    If Len(nome) > 0 Then
        abrangencia = "nacional"
        InfoDia = DIA_FERIADO
        Exit Function
    End If

    For i = 1 To mQtd
        If Len(mUF(i)) = 0 Or mUF(i) = uf Then
            If Len(mCid(i)) = 0 Or mCid(i) = cidade Then
                If Bate(mQuando(i), d) Then
                    If mAbr(i) = "PONTO FACULTATIVO" Then
                        If Len(fac) = 0 Then
                            fac = mNome(i)
                            facAbr = "ponto facultativo"
                        End If
                    Else
                        nome = mNome(i)
                        If mAbr(i) = "MUNICIPAL" Then
                            abrangencia = "municipal"
                        ElseIf mAbr(i) = "ESTADUAL" Then
                            abrangencia = "estadual"
                        Else
                            abrangencia = LCase$(mAbr(i))
                        End If
                        InfoDia = DIA_FERIADO
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i

    If Weekday(d, vbSunday) = vbSaturday Then
        InfoDia = DIA_SABADO
        Exit Function
    End If
    If Weekday(d, vbSunday) = vbSunday Then
        InfoDia = DIA_DOMINGO
        Exit Function
    End If

    If Len(fac) = 0 Then
        fac = FacultativoNacional(d)
        facAbr = "ponto facultativo"
    End If
    If Len(fac) > 0 Then
        nome = fac
        abrangencia = facAbr
        InfoDia = DIA_FACULTATIVO
    End If
End Function

Public Function Bloqueia(ByVal info As Long) As Boolean
    Bloqueia = (info = DIA_SABADO Or info = DIA_DOMINGO Or info = DIA_FERIADO)
End Function

Public Function EhDiaUtil(ByVal d As Date, ByVal uf As String, ByVal cidade As String) As Boolean
    Dim n As String
    Dim a As String
    EhDiaUtil = Not Bloqueia(InfoDia(d, uf, cidade, n, a))
End Function

Public Function DiaUtilSeguinte(ByVal d As Date, ByVal uf As String, ByVal cidade As String) As Date
    Dim k As Long
    d = d + 1
    Do While Not EhDiaUtil(d, uf, cidade) And k < 30
        d = d + 1
        k = k + 1
    Loop
    DiaUtilSeguinte = d
End Function

Public Function DiaUtilAnterior(ByVal d As Date, ByVal uf As String, ByVal cidade As String) As Date
    Dim k As Long
    d = d - 1
    Do While Not EhDiaUtil(d, uf, cidade) And k < 30
        d = d - 1
        k = k + 1
    Loop
    DiaUtilAnterior = d
End Function

' Texto do motivo: "sábado", "domingo", "feriado nacional (Finados)" ...
Public Function MotivoDia(ByVal info As Long, ByVal nome As String, ByVal abrangencia As String) As String
    Select Case info
        Case DIA_SABADO
            MotivoDia = "sábado"
        Case DIA_DOMINGO
            MotivoDia = "domingo"
        Case DIA_FERIADO
            MotivoDia = "feriado " & abrangencia & " (" & nome & ")"
        Case DIA_FACULTATIVO
            MotivoDia = "ponto facultativo (" & nome & ")"
    End Select
End Function

Public Function EtiquetaDia(ByVal info As Long) As String
    Select Case info
        Case DIA_SABADO
            EtiquetaDia = "SÁBADO"
        Case DIA_DOMINGO
            EtiquetaDia = "DOMINGO"
        Case DIA_FERIADO
            EtiquetaDia = "FERIADO"
        Case DIA_FACULTATIVO
            EtiquetaDia = "P. FACULT."
    End Select
End Function

Public Function DiaSemanaCurto(ByVal d As Date) As String
    DiaSemanaCurto = Choose(Weekday(d, vbSunday), "dom", "seg", "ter", "qua", "qui", "sex", "sáb")
End Function

Public Function NomeMes(ByVal m As Long) As String
    NomeMes = Choose(m, "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", _
                     "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO")
End Function

' "15/11/2026", "15112026" ou data -> data (False se não for data a partir de 2000).
Public Function LerDataFer(ByVal texto As String, ByRef d As Date) As Boolean
    Dim s As String
    Dim dd As Long
    Dim mm As Long
    Dim aa As Long
    s = Replace(Replace(Replace(Trim$(texto), "/", ""), "-", ""), ".", "")
    If Len(s) <> 8 Then Exit Function
    If Not s Like "########" Then Exit Function
    dd = CLng(Left$(s, 2))
    mm = CLng(Mid$(s, 3, 2))
    aa = CLng(Right$(s, 4))
    If mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 Or aa < 2000 Or aa > 2100 Then Exit Function
    d = DateSerial(aa, mm, dd)
    If Day(d) <> dd Then Exit Function
    LerDataFer = True
End Function

'-------------------------------------------------------------------------------
' Fontes da identidade visual (Segoe UI de reserva)
'-------------------------------------------------------------------------------
Public Function FonteFeriados() As String
    If Len(mFonteTexto) = 0 Then DetectarFontes
    FonteFeriados = mFonteTexto
End Function

Public Function FonteTituloFeriados() As String
    If Len(mFonteTitulo) = 0 Then DetectarFontes
    FonteTituloFeriados = mFonteTitulo
End Function

Private Sub DetectarFontes()
    Dim achado As String
    On Error Resume Next
    achado = Dir$(Environ$("WINDIR") & "\Fonts\Segoe UI*")
    If Len(achado) = 0 Then achado = Dir$(Environ$("LOCALAPPDATA") & "\Microsoft\Windows\Fonts\Segoe UI*")
    If Len(achado) > 0 Then
        mFonteTexto = "Segoe UI"
        mFonteTitulo = "Segoe UI Black"
    Else
        mFonteTexto = "Segoe UI"
        mFonteTitulo = "Segoe UI Black"
    End If
End Sub
