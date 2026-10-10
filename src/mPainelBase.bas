Option Explicit

'====================================================================
' PAINEL - MÓDULO PRINCIPAL
'
' Abrir o painel ......: AbrirPainelObras
'
' Regras:
'  - Colunas da tb_Obras ANTES do primeiro cabeçalho amarelo
'    = dados da obra (somente leitura no painel)
'  - Colunas com cabeçalho AMARELO = datas editáveis (painel de datas)
'  - HISTÓRICO = histórico / ÚLTIMO REGISTRO = última observação
'  - Toda gravação feita pelo painel é registrada na
'    tb_Alteracoes (com valor anterior correto)
'  - Células com fórmula nunca são sobrescritas
'====================================================================

'----------------------------- CONFIGURAÇÃO -------------------------
Public Const PO_TABELA As String = "tb_Obras"
Public Const PO_TABELA_HIST As String = "tb_Alteracoes"

Public Const PO_COL_CODIGO As String = "CÓDIGO"
Public Const PO_COL_DESCRICAO As String = "NOME"
Public Const PO_COL_CIDADE As String = "CIDADE"
Public Const PO_COL_UF As String = "UF"
Public Const PO_COL_OBS As String = "HISTÓRICO"
Public Const PO_COL_ULTIMA As String = "ÚLTIMO REGISTRO"
Public Const PO_COL_MODALIDADE As String = "TIPO DE OBRA"
Public Const PO_COL_FASE As String = "ETAPA"
Public Const PO_COL_STATUS As String = "SITUAÇÃO"
Public Const PO_COL_ENTREGA As String = "ENTREGA"

Public Const PO_TAM_SEPARADOR As Long = 60

'Cores do painel
Public Const PO_COR_AZUL As Long = 3877150        'RGB(30, 41, 59)
Public Const PO_COR_AMARELO As Long = 2408443       'RGB(251, 191, 36)
Public Const PO_COR_FUNDO As Long = 16579320      'RGB(248, 250, 252)
Public Const PO_COR_LEITURA As Long = 15921906    'RGB(242, 242, 242)
Public Const PO_COR_EDICAO As Long = 14745599     'RGB(255, 255, 224)
Public Const PO_COR_ERRO As Long = 14935295       'RGB(255, 228, 227)
Public Const PO_COR_BORDA As Long = 14141640      'RGB(200, 200, 215)




'====================================================================
' LOCALIZAÇÃO DE TABELA / COLUNAS
'====================================================================

Public Function PO_TabelaPorNome(ByVal nome As String) As ListObject

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If UCase$(lo.Name) = UCase$(nome) Then
                Set PO_TabelaPorNome = lo
                Exit Function
            End If
        Next lo
    Next ws

End Function

'Índice da coluna pelo nome (0 = não existe)
Public Function PO_ColIdx(ByVal tb As ListObject, ByVal nome As String) As Long

    Dim i As Long
    Dim alvo As String

    alvo = PO_Norm(nome)

    For i = 1 To tb.ListColumns.Count
        If PO_Norm(tb.ListColumns(i).Name) = alvo Then
            PO_ColIdx = i
            Exit Function
        End If
    Next i

End Function

'Cabeçalho pintado de amarelo?
Public Function PO_EhEditavel(ByVal tb As ListObject, ByVal idx As Long) As Boolean

    Dim cor As Long
    Dim r As Long, g As Long, b As Long

    On Error GoTo Fim

    cor = tb.HeaderRowRange.Cells(1, idx).Interior.Color

    r = cor Mod 256
    g = (cor \ 256) Mod 256
    b = (cor \ 65536) Mod 256

    'Aceita tons de amarelo / ouro (FFC000, FFFF00, FFD966...)
    PO_EhEditavel = (r >= 230 And g >= 170 And b <= 210 And (r - b) >= 45)

Fim:
End Function

'Lista (1..n) com os índices das colunas amarelas. Empty se não houver.
Public Function PO_ColunasEditaveis(ByVal tb As ListObject) As Variant

    Dim i As Long, n As Long
    Dim arr() As Long

    ReDim arr(1 To tb.ListColumns.Count)

    For i = 1 To tb.ListColumns.Count
        If PO_EhEditavel(tb, i) Then
            n = n + 1
            arr(n) = i
        End If
    Next i

    If n = 0 Then
        PO_ColunasEditaveis = Empty
    Else
        ReDim Preserve arr(1 To n)
        PO_ColunasEditaveis = arr
    End If

End Function

'Lista (1..n) com as colunas de "dados da obra" (antes do 1º amarelo)
Public Function PO_ColunasInfo(ByVal tb As ListObject) As Variant

    Dim i As Long, n As Long, ultima As Long
    Dim arr() As Long
    Dim ed As Variant

    ed = PO_ColunasEditaveis(tb)

    If IsArray(ed) Then
        ultima = ed(1) - 1
    Else
        ultima = tb.ListColumns.Count
    End If

    If ultima < 1 Then
        PO_ColunasInfo = Empty
        Exit Function
    End If

    ReDim arr(1 To ultima)

    For i = 1 To ultima
        If i <> PO_ColIdx(tb, PO_COL_OBS) And i <> PO_ColIdx(tb, PO_COL_ULTIMA) Then
            n = n + 1
            arr(n) = i
        End If
    Next i

    If n = 0 Then
        PO_ColunasInfo = Empty
    Else
        ReDim Preserve arr(1 To n)
        PO_ColunasInfo = arr
    End If

End Function

'Nome curto da etapa: "DT PREV. ALVARÁ" -> "ALVARÁ"
Public Function PO_NomeEtapa(ByVal cabecalho As String) As String

    Dim s As String
    Dim mudou As Boolean
    Dim prefixos As Variant
    Dim p As Variant

    prefixos = Array("DT. ", "DT ", "PREV. ", "PREV ", "PREVISÃO ", "PREVISTO ")
    s = UCase$(Trim$(cabecalho))

    Do
        mudou = False
        For Each p In prefixos
            If Left$(s, Len(p)) = p Then
                s = Trim$(Mid$(s, Len(p) + 1))
                mudou = True
            End If
        Next p
    Loop While mudou

    If s = "" Then s = cabecalho
    PO_NomeEtapa = s

End Function


'====================================================================
' LEITURA / FORMATAÇÃO DE VALORES
'====================================================================
Public Function PO_Valor(ByVal tb As ListObject, ByVal lr As ListRow, ByVal nomeCol As String) As Variant

    Dim idx As Long

    idx = PO_ColIdx(tb, nomeCol)

    If idx = 0 Then
        PO_Valor = Empty
    Else
        PO_Valor = lr.Range.Cells(1, idx).value
    End If

End Function

Public Function PO_TextoCampo(ByVal tb As ListObject, ByVal lr As ListRow, ByVal nomeCol As String) As String
    PO_TextoCampo = PO_TextoValor(PO_Valor(tb, lr, nomeCol))
End Function

'Converte qualquer valor de célula em texto de exibição
Public Function PO_TextoValor(ByVal v As Variant) As String

    On Error GoTo vazio

    If IsError(v) Then GoTo vazio
    If IsEmpty(v) Then GoTo vazio
    If IsNull(v) Then GoTo vazio

    If VarType(v) = vbDate Then
        If CDbl(v) <= 1 Then GoTo vazio
        If CDbl(v) = Int(CDbl(v)) Then
            PO_TextoValor = Format$(v, "dd\/mm\/yyyy")
        Else
            PO_TextoValor = Format$(v, "dd\/mm\/yyyy hh:mm")
        End If
        Exit Function
    End If

    PO_TextoValor = Trim$(CStr(v))
    Exit Function

vazio:
    PO_TextoValor = ""

End Function

'Maiúsculas, sem acento, sem quebras
Public Function PO_Norm(ByVal s As String) As String

    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    PO_Norm = UCase$(Trim$(PO_SemAcento(s)))

End Function

Public Function PO_SemAcento(ByVal s As String) As String

    Const COM As String = "ÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇáàâãäéèêëíìîïóòôõöúùûüç"
    Const SEM As String = "AAAAAEEEEIIIIOOOOOUUUUCaaaaaeeeeiiiiooooouuuuc"

    Dim i As Long

    For i = 1 To Len(COM)
        s = Replace(s, Mid$(COM, i, 1), Mid$(SEM, i, 1))
    Next i

    PO_SemAcento = s

End Function

'Quebras de linha no padrão do TextBox / da célula (vbCrLf)
Public Function PO_ParaCaixa(ByVal s As String) As String

    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    PO_ParaCaixa = Replace(s, vbLf, vbCrLf)

End Function

'Remove espaços e quebras de linha no fim do texto
Public Function PO_TrimFim(ByVal s As String) As String

    Do While Len(s) > 0
        Select Case Right$(s, 1)
            Case vbCr, vbLf, " ", vbTab
                s = Left$(s, Len(s) - 1)
            Case Else
                Exit Do
        End Select
    Loop

    PO_TrimFim = s

End Function

'Texto "dd/mm/aaaa", "ddmmaaaa", "ddmmaa" -> Date.  "" ou "-" = vazio.
Public Function PO_TextoParaData(ByVal s As String, ByRef d As Date, ByRef vazio As Boolean) As Boolean

    Dim p() As String
    Dim dd As Long, mm As Long, yy As Long

    On Error GoTo Invalido

    s = Trim$(s)
    vazio = False

    If s = "" Or s = "-" Then
        vazio = True
        PO_TextoParaData = True
        Exit Function
    End If

    If InStr(s, "/") > 0 Then
        p = Split(s, "/")
        If UBound(p) <> 2 Then GoTo Invalido
        If Not (p(0) Like "#" Or p(0) Like "##") Then GoTo Invalido
        If Not (p(1) Like "#" Or p(1) Like "##") Then GoTo Invalido
        If Not (p(2) Like "##" Or p(2) Like "####") Then GoTo Invalido
        dd = CLng(p(0)): mm = CLng(p(1)): yy = CLng(p(2))
    ElseIf s Like "########" Then
        dd = CLng(Left$(s, 2)): mm = CLng(Mid$(s, 3, 2)): yy = CLng(Right$(s, 4))
    ElseIf s Like "######" Then
        dd = CLng(Left$(s, 2)): mm = CLng(Mid$(s, 3, 2)): yy = CLng(Right$(s, 2))
    Else
        GoTo Invalido
    End If

    If yy < 100 Then yy = yy + 2000
    If yy < 1990 Or yy > 2100 Then GoTo Invalido
    If mm < 1 Or mm > 12 Then GoTo Invalido
    If dd < 1 Or dd > 31 Then GoTo Invalido

    d = DateSerial(yy, mm, dd)
    If Day(d) <> dd Or Month(d) <> mm Then GoTo Invalido

    PO_TextoParaData = True
    Exit Function

Invalido:
    PO_TextoParaData = False

End Function


'====================================================================
' GRAVAÇÃO NA CÉLULA + LOG
'   Retorno:  1 = alterou | 0 = já era igual | -1 = não gravou
'====================================================================
Public Function PO_Gravar(ByVal tb As ListObject, ByVal lr As ListRow, _
                          ByVal idx As Long, ByVal novo As Variant) As Long

    Dim cel As Range
    Dim antigo As Variant
    Dim evt As Boolean

    evt = Application.EnableEvents

    On Error GoTo Falha

    Set cel = lr.Range.Cells(1, idx)

    If cel.HasFormula Then
        MsgBox "O campo '" & tb.ListColumns(idx).Name & "' contém fórmula e não foi alterado.", _
               vbExclamation, "painel"
        PO_Gravar = -1
        Exit Function
    End If

    antigo = cel.value

    If PO_Iguais(antigo, novo) Then
        PO_Gravar = 0
        Exit Function
    End If

    'Eventos desligados: o log é feito aqui com o valor anterior correto
    Application.EnableEvents = False

    If IsEmpty(novo) Then
        cel.ClearContents
    Else
        If VarType(novo) = vbDate Then
            If cel.NumberFormat = "General" Or cel.NumberFormat = "Geral" Then
                cel.NumberFormat = "dd/mm/yyyy"
            End If
        End If
        cel.value = novo
    End If

    Application.EnableEvents = evt

    PO_RegistrarLog tb, lr, idx, antigo, novo

    PO_Gravar = 1
    Exit Function

Falha:
    Application.EnableEvents = evt
    PO_Gravar = -1
    MsgBox "Não foi possível gravar o campo '" & tb.ListColumns(idx).Name & "'." & _
           vbCrLf & vbCrLf & Err.Description, vbCritical, "painel"

End Function

Private Function PO_Vazio(ByVal v As Variant) As Boolean

    If IsError(v) Then Exit Function
    If IsEmpty(v) Then
        PO_Vazio = True
    ElseIf VarType(v) = vbString Then
        PO_Vazio = (Trim$(v) = "" Or Trim$(v) = "-")
    End If

End Function

Private Function PO_Iguais(ByVal a As Variant, ByVal b As Variant) As Boolean

    On Error GoTo Diferente

    If IsError(a) Or IsError(b) Then GoTo Diferente

    If PO_Vazio(a) And PO_Vazio(b) Then
        PO_Iguais = True
        Exit Function
    End If

    If PO_Vazio(a) Or PO_Vazio(b) Then GoTo Diferente

    If VarType(a) = vbDate Or VarType(b) = vbDate Then
        If IsDate(a) And IsDate(b) Then
            PO_Iguais = (Int(CDbl(CDate(a))) = Int(CDbl(CDate(b))))
            Exit Function
        End If
    End If

    PO_Iguais = (CStr(a) = CStr(b))
    Exit Function

Diferente:
    PO_Iguais = False

End Function

Private Sub PO_RegistrarLog(ByVal tb As ListObject, ByVal lr As ListRow, _
                            ByVal idx As Long, ByVal antigo As Variant, ByVal novo As Variant)

    Dim th As ListObject
    Dim nl As ListRow
    Dim evt As Boolean
    Dim campos As Variant
    Dim c As Variant

    evt = Application.EnableEvents

    On Error GoTo Fim

    Set th = PO_TabelaPorNome(PO_TABELA_HIST)
    If th Is Nothing Then Exit Sub

    Application.EnableEvents = False

    Set nl = th.ListRows.Add

    campos = Array("CÓDIGO", "NOME", "CLIENTE", "UF", "CIDADE")
    For Each c In campos
        PO_SetLog th, nl, CStr(c), PO_ValorLog(PO_Valor(tb, lr, CStr(c)))
    Next c

    PO_SetLog th, nl, "CAMPO", tb.ListColumns(idx).Name
    PO_SetLog th, nl, "ANTES", PO_ValorLog(antigo)
    PO_SetLog th, nl, "DEPOIS", PO_ValorLog(novo)
    PO_SetLog th, nl, "USUÁRIO", Environ$("Username")
    PO_SetLog th, nl, "DATA/HORA", Now

Fim:
    Application.EnableEvents = evt

End Sub

Private Sub PO_SetLog(ByVal th As ListObject, ByVal nl As ListRow, ByVal nomeCol As String, ByVal v As Variant)

    Dim idx As Long

    idx = PO_ColIdx(th, nomeCol)
    If idx > 0 Then nl.Range.Cells(1, idx).value = v

End Sub

Private Function PO_ValorLog(ByVal v As Variant) As Variant

    If IsError(v) Then
        PO_ValorLog = Empty
    Else
        PO_ValorLog = v
    End If

End Function


'====================================================================
' LEITURA DO HISTÓRICO DE OBSERVAÇÕES
'   Cada item da coleção = Array(data As Date|Empty, rotulo, texto)
'   Uma entrada começa com "dd/mm/aaaa ..." ; linhas "-----" separam.
'====================================================================
Public Function PO_ParseObs(ByVal texto As String) As Collection

    Dim col As New Collection
    Dim linhas() As String
    Dim i As Long
    Dim ln As String
    Dim temAtual As Boolean
    Dim dAtual As Variant
    Dim rotAtual As String
    Dim txtAtual As String

    texto = Replace(texto, vbCrLf, vbLf)
    texto = Replace(texto, vbCr, vbLf)
    linhas = Split(texto, vbLf)

    For i = LBound(linhas) To UBound(linhas)

        ln = Trim$(linhas(i))

        If PO_EhSeparador(ln) Then

            If temAtual Then PO_AddEntrada col, dAtual, rotAtual, txtAtual
            temAtual = False

        ElseIf ln Like "##/##/####*" Then

            If temAtual Then PO_AddEntrada col, dAtual, rotAtual, txtAtual
            PO_SepararCabecalho ln, dAtual, rotAtual, txtAtual
            temAtual = True

        ElseIf temAtual Then

            If txtAtual = "" Then
                txtAtual = ln
            Else
                txtAtual = txtAtual & vbLf & ln
            End If

        ElseIf ln <> "" Then

            'Texto sem data no início
            dAtual = Empty
            rotAtual = ""
            txtAtual = ln
            temAtual = True

        End If

    Next i

    If temAtual Then PO_AddEntrada col, dAtual, rotAtual, txtAtual

    Set PO_ParseObs = col

End Function

Private Function PO_EhSeparador(ByVal ln As String) As Boolean

    Dim s As String

    If Len(ln) < 5 Then Exit Function

    s = ln
    s = Replace(s, "-", "")
    s = Replace(s, "_", "")
    s = Replace(s, "=", "")
    s = Replace(s, ChrW(8211), "")
    s = Replace(s, ChrW(8212), "")

    PO_EhSeparador = (Len(Trim$(s)) = 0)

End Function

Private Sub PO_SepararCabecalho(ByVal ln As String, ByRef d As Variant, _
                                ByRef rotulo As String, ByRef corpo As String)

    Dim dt As Date
    Dim vz As Boolean
    Dim resto As String

    If PO_TextoParaData(Left$(ln, 10), dt, vz) And Not vz Then
        d = dt
    Else
        d = Empty
    End If

    rotulo = Left$(ln, 10)
    resto = PO_TirarPrefixo(Mid$(ln, 11))

    If resto Like "##:##*" Then
        rotulo = rotulo & " " & Left$(resto, 5)
        resto = Mid$(resto, 6)
        If resto Like ":##*" Then resto = Mid$(resto, 4)   'segundos
        resto = PO_TirarPrefixo(resto)
    End If

    corpo = resto

End Sub

Private Function PO_TirarPrefixo(ByVal s As String) As String

    Dim chars As String

    chars = " -:." & ChrW(8211) & ChrW(8212) & vbTab

    Do While Len(s) > 0
        If InStr(1, chars, Left$(s, 1)) > 0 Then
            s = Mid$(s, 2)
        Else
            Exit Do
        End If
    Loop

    PO_TirarPrefixo = s

End Function

Private Sub PO_AddEntrada(ByVal col As Collection, ByVal d As Variant, _
                          ByVal rotulo As String, ByVal texto As String)

    texto = PO_TrimFim(texto)

    If texto = "" And rotulo = "" Then Exit Sub

    col.Add Array(d, rotulo, texto)

End Sub


'====================================================================
' ATA EM WORD
'   lista = array (1..qtd) com os índices das linhas (ListRows)
'
'   A ata traz, de cada obra, TUDO o que está na coluna ÚLTIMA
'   ATUALIZAÇÃO (o texto inteiro da célula, com todas as linhas) e nada
'   do histórico. Obra com a ÚLTIMO REGISTRO em branco vai para a
'   lista "Projetos sem atualização".
'====================================================================
Public Sub PO_GerarAta(ByVal tb As ListObject, ByVal lista As Variant, ByVal qtd As Long)

    Dim i As Long, g As Long
    Dim lr As ListRow
    Dim txt() As String          'texto da última atualização (sem a data inicial)
    Dim rot() As String          'data/hora inicial, quando houver ("dd/mm/aaaa hh:mm")
    Dim grupos As New Collection
    Dim nomeGrupo As String
    Dim existe As Boolean
    Dim semAtual As String
    Dim qtdCom As Long

    Dim wd As Object, doc As Object, r As Object
    Dim linhaInfo As String

    If qtd < 1 Then
        MsgBox "Nenhuma obra na lista do painel.", vbExclamation, "Gerar ata"
        Exit Sub
    End If

    If PO_ColIdx(tb, PO_COL_ULTIMA) = 0 Then
        MsgBox "A coluna '" & PO_COL_ULTIMA & "' não foi encontrada na tabela '" & PO_TABELA & "'.", _
               vbExclamation, "Gerar ata"
        Exit Sub
    End If

    '--------------------------------------------- lê a última atualização
    ReDim txt(1 To qtd)
    ReDim rot(1 To qtd)

    For i = 1 To qtd

        Set lr = tb.ListRows(lista(i))

        PO_UltimaParaAta PO_TextoCampo(tb, lr, PO_COL_ULTIMA), rot(i), txt(i)

        If txt(i) <> "" Or rot(i) <> "" Then qtdCom = qtdCom + 1

        'grupos (TIPO DE OBRA) na ordem em que aparecem
        nomeGrupo = PO_GrupoAta(tb, lr)
        existe = False
        For g = 1 To grupos.Count
            If grupos(g) = nomeGrupo Then existe = True: Exit For
        Next g
        If Not existe Then grupos.Add nomeGrupo

    Next i

    If qtdCom = 0 Then
        MsgBox "Nenhuma obra da lista tem texto na coluna '" & PO_COL_ULTIMA & "'.", vbInformation, "Gerar ata"
        Exit Sub
    End If

    '--------------------------------------------- Word
    On Error Resume Next
    Set wd = CreateObject("Word.Application")
    On Error GoTo Falha

    If wd Is Nothing Then
        MsgBox "Não foi possível abrir o Microsoft Word.", vbCritical, "Gerar ata"
        Exit Sub
    End If

    Application.StatusBar = "Gerando ata no Word..."
    Application.Cursor = xlWait

    Set doc = wd.Documents.Add

    With doc.PageSetup
        .TopMargin = 56.7
        .BottomMargin = 56.7
        .LeftMargin = 56.7
        .RightMargin = 56.7
    End With

    doc.Styles(-1).Font.Name = "Calibri"     'Normal
    doc.Styles(-1).Font.Size = 10.5

    PO_WAdd doc, "ATA DE REUNIÃO – FOLLOW UP DE OBRAS", -63               'Título
    PO_WAdd doc, "Data: " & Format$(Date, "dd\/mm\/yyyy") & "     Horário: " & Format$(Time, "hh:mm"), -1
    PO_WAdd doc, "Base: última atualização registrada de cada projeto", -1
    PO_WAdd doc, "Projetos com atualização: " & qtdCom & " de " & qtd, -1
    PO_WAdd doc, "Participantes: ", -1
    PO_WAdd doc, "", -1

    For g = 1 To grupos.Count

        'só cria o título do grupo se houver obra com conteúdo
        existe = False
        For i = 1 To qtd
            If txt(i) <> "" Or rot(i) <> "" Then
                If PO_GrupoAta(tb, tb.ListRows(lista(i))) = grupos(g) Then existe = True: Exit For
            End If
        Next i

        If existe Then

            PO_WAdd doc, grupos(g), -2                                       'Título 1

            For i = 1 To qtd

                Set lr = tb.ListRows(lista(i))

                If (txt(i) <> "" Or rot(i) <> "") And PO_GrupoAta(tb, lr) = grupos(g) Then

                    PO_WAdd doc, PO_TextoCampo(tb, lr, PO_COL_DESCRICAO) & _
                                 "  (CÓDIGO " & PO_TextoCampo(tb, lr, PO_COL_CODIGO) & ")", -3   'Título 2

                    linhaInfo = PO_LinhaInfoAta(tb, lr)
                    If linhaInfo <> "" Then
                        Set r = PO_WAdd(doc, linhaInfo, -1)
                        r.Font.Italic = True
                        r.Font.Size = 9
                        r.Font.Color = 8421504
                    End If

                    'um único item com o texto inteiro (as quebras de linha são mantidas)
                    If rot(i) <> "" And txt(i) <> "" Then
                        Set r = PO_WAdd(doc, rot(i) & " " & ChrW(8211) & " " & PO_ParaWord(txt(i)), -49)
                        doc.Range(r.Start, r.Start + Len(rot(i))).Font.Bold = True
                    ElseIf rot(i) <> "" Then
                        Set r = PO_WAdd(doc, rot(i), -49)
                        r.Font.Bold = True
                    Else
                        PO_WAdd doc, PO_ParaWord(txt(i)), -49
                    End If

                End If

            Next i

        End If

    Next g

    'Obras sem texto na última atualização
    semAtual = ""
    For i = 1 To qtd
        If txt(i) = "" And rot(i) = "" Then
            Set lr = tb.ListRows(lista(i))
            If semAtual <> "" Then semAtual = semAtual & "; "
            semAtual = semAtual & PO_TextoCampo(tb, lr, PO_COL_DESCRICAO)
        End If
    Next i

    If semAtual <> "" Then
        PO_WAdd doc, "Projetos sem atualização", -2
        PO_WAdd doc, semAtual, -1
    End If

    On Error Resume Next
    doc.Paragraphs(doc.Paragraphs.Count).Style = doc.Styles(-1)
    doc.Content.LanguageID = 1046      'Português (Brasil)
    On Error GoTo Falha

    wd.Visible = True
    doc.Activate
    On Error Resume Next
    wd.Activate
    On Error GoTo 0

    Application.StatusBar = False
    Application.Cursor = xlDefault
    Exit Sub

Falha:
    Application.StatusBar = False
    Application.Cursor = xlDefault
    MsgBox "Erro ao gerar a ata:" & vbCrLf & Err.Description, vbCritical, "Gerar ata"
    On Error Resume Next
    If Not wd Is Nothing Then wd.Visible = True

End Sub

'Prepara o texto da ÚLTIMO REGISTRO para a ata: separa a data/hora do
'início (se houver) para sair em negrito e devolve TODO o restante do
'texto, com todas as linhas, exatamente como está na célula.
Private Sub PO_UltimaParaAta(ByVal bruto As String, ByRef rotulo As String, ByRef corpo As String)

    Dim s As String
    Dim primeira As String, resto As String
    Dim p As Long
    Dim d As Variant

    rotulo = ""
    corpo = ""

    s = Replace(bruto, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = PO_TrimFim(s)

    'tira linhas em branco do início
    Do While Left$(s, 1) = vbLf Or Left$(s, 1) = " "
        s = Mid$(s, 2)
    Loop

    If s = "" Then Exit Sub

    If s Like "##/##/####*" Then

        p = InStr(s, vbLf)

        If p > 0 Then
            primeira = Left$(s, p - 1)
            resto = Mid$(s, p + 1)
        Else
            primeira = s
            resto = ""
        End If

        PO_SepararCabecalho Trim$(primeira), d, rotulo, corpo

        If resto <> "" Then
            If corpo <> "" Then
                corpo = corpo & vbLf & resto
            Else
                corpo = resto
            End If
        End If

    Else

        corpo = s

    End If

End Sub

Private Function PO_GrupoAta(ByVal tb As ListObject, ByVal lr As ListRow) As String

    Dim s As String

    s = UCase$(PO_TextoCampo(tb, lr, PO_COL_MODALIDADE))

    Select Case PO_SemAcento(s)
        Case "CONSTRUCAO": s = "CONSTRUÇÃO"
        Case "AMPLIACAO": s = "AMPLIAÇÃO"
        Case "": s = "OUTRAS OBRAS"
    End Select

    PO_GrupoAta = s

End Function

Private Function PO_LinhaInfoAta(ByVal tb As ListObject, ByVal lr As ListRow) As String

    Dim partes As String
    Dim v As String

    v = PO_TextoCampo(tb, lr, PO_COL_CIDADE)
    If v <> "" Then
        If PO_TextoCampo(tb, lr, PO_COL_UF) <> "" Then v = v & "/" & PO_TextoCampo(tb, lr, PO_COL_UF)
        partes = partes & "  |  " & v
    End If

    v = PO_TextoCampo(tb, lr, PO_COL_FASE)
    If v <> "" Then partes = partes & "  |  Etapa: " & v

    v = PO_TextoCampo(tb, lr, PO_COL_STATUS)
    If v <> "" Then partes = partes & "  |  Situação: " & v

    v = PO_TextoCampo(tb, lr, PO_COL_ENTREGA)
    If v <> "" Then partes = partes & "  |  Entrega prevista: " & v

    If Len(partes) > 5 Then partes = Mid$(partes, 6)
    PO_LinhaInfoAta = partes

End Function

'Acrescenta um parágrafo no fim do documento e devolve o Range dele
Private Function PO_WAdd(ByVal doc As Object, ByVal texto As String, ByVal estilo As Long) As Object

    Dim r As Object
    Dim pos As Long

    pos = doc.Content.End - 1
    Set r = doc.Range(pos, pos)

    r.InsertAfter texto
    r.Style = doc.Styles(estilo)
    r.Font.Reset
    r.InsertParagraphAfter

    Set PO_WAdd = doc.Range(r.Start, r.Start + Len(texto))

End Function

Private Function PO_ParaWord(ByVal s As String) As String

    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    PO_ParaWord = Replace(s, vbLf, Chr(11))     'quebra de linha manual

End Function


'====================================================================
' PAINEL DE OBRAS (tabela tb_Obras, aba OBRAS)
'   Abrir o painel ......: AbrirPainelObras
'====================================================================
Public Sub AbrirPainelObras()

    Dim tb As ListObject
    Dim idx As Long
    Dim f As Object
    Dim painel As Object

    Set tb = PO_TabelaPorNome("tb_Obras")

    If tb Is Nothing Then
        MsgBox "A tabela 'tb_Obras' não foi encontrada.", vbExclamation, "Painel de Obras"
        Exit Sub
    End If

    'Se a célula ativa estiver dentro da tabela, abre naquela obra
    idx = 0
    On Error Resume Next
    If ActiveSheet.Name = tb.Parent.Name Then
        If Not Intersect(ActiveCell, tb.DataBodyRange) Is Nothing Then
            idx = ActiveCell.row - tb.DataBodyRange.row + 1
        End If
    End If
    On Error GoTo 0

    'O formulário é localizado pelo nome, para este módulo continuar
    'compilando mesmo antes de o frmPainelObras ser criado.
    For Each f In VBA.UserForms
        If f.Name = "frmPainelObras" Then Set painel = f
    Next f

    If painel Is Nothing Then

        On Error Resume Next
        Set painel = VBA.UserForms.Add("frmPainelObras")
        On Error GoTo 0

        If painel Is Nothing Then
            MsgBox "O formulário 'frmPainelObras' ainda não existe neste arquivo." & vbCrLf & vbCrLf & _
                   "Crie um UserForm com esse nome e cole nele o código do painel.", _
                   vbExclamation, "Painel de Obras"
            Exit Sub
        End If

    End If

    painel.IrParaLinha idx

    If Not painel.Visible Then painel.Show vbModeless

End Sub


