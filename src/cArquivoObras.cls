Option Explicit
'===============================================================================
' cArquivoObras - leva obras entregues da tb_Obras para ENTREGUES
'
'   - a linha inteira vai para a tabela tb_Entregues, casando as colunas pelo
'     NOME (fórmulas entram como valor; CÓDIGO como texto); colunas que não
'     existem lá ficam de fora e são informadas
'   - registra no histórico de alterações quem moveu e quando
'   - conclui o marcador "para analisar" da obra, se houver
'   - filtros e segmentações são guardados, limpos e reaplicados; a saída
'     da obra do follow up é conferida
'===============================================================================

Private Const T_Obras As String = "tb_Obras"
Private Const T_ENTREGA As String = "tb_Entregues"
Private Const T_HIST As String = "tb_Alteracoes"
Private Const T_MARC As String = "tb_Marcadores"

Private mErro As String
Private mAvisoFiltro As String
Private mFiltros As Collection
Private mSlicers As Collection
Private mEventos As Boolean
Private mTela As Boolean

Public Property Get UltimoErro() As String
    UltimoErro = mErro
End Property

' Observação sobre filtros que não puderam ser refeitos (vazio = tudo certo).
Public Property Get AvisoFiltro() As String
    AvisoFiltro = mAvisoFiltro
End Property

'-------------------------------------------------------------------------------
' Tabelas e colunas
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
' Tabelas e colunas
'-------------------------------------------------------------------------------
Public Function Tabela(ByVal nome As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        On Error Resume Next
        Set lo = ws.ListObjects(nome)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set Tabela = lo
            Exit Function
        End If
    Next ws
End Function

Private Function Norm(ByVal s As String) As String
    s = Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), Chr$(160), " ")
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    Norm = UCase$(Trim$(s))
End Function

Public Function Col(ByVal tb As ListObject, ByVal cabecalho As String) As Long
    Dim i As Long
    Dim alvo As String
    If tb Is Nothing Then Exit Function
    alvo = Norm(cabecalho)
    For i = 1 To tb.ListColumns.Count
        If Norm(tb.ListColumns(i).Name) = alvo Then
            Col = i
            Exit Function
        End If
    Next i
End Function

Private Function Chave(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    Chave = Trim$(Replace(CStr(v), Chr$(160), " "))
End Function

' Linha da obra na tabela, procurando o CÓDIGO como texto (0 = não achou).
Public Function LinhaDe(ByVal tb As ListObject, ByVal codObra As String) As Long
    Dim c As Long
    Dim i As Long
    Dim alvo As String
    If tb Is Nothing Then Exit Function
    If tb.ListRows.Count = 0 Then Exit Function
    c = Col(tb, "CÓDIGO")
    If c = 0 Then Exit Function
    alvo = Trim$(codObra)
    If Len(alvo) = 0 Then Exit Function
    For i = 1 To tb.ListRows.Count
        If Chave(tb.DataBodyRange.Cells(i, c).Value) = alvo Then
            LinhaDe = i
            Exit Function
        End If
    Next i
End Function

'-------------------------------------------------------------------------------
' Consultas
'-------------------------------------------------------------------------------
' Valor da coluna na obra do follow up, como texto de exibição.
Public Function Texto(ByVal codObra As String, ByVal coluna As String) As String
    Dim tb As ListObject
    Dim i As Long
    Dim c As Long
    Dim v As Variant
    On Error GoTo Fim
    Set tb = Tabela(T_Obras)
    i = LinhaDe(tb, codObra)
    c = Col(tb, coluna)
    If i = 0 Or c = 0 Then Exit Function
    v = tb.DataBodyRange.Cells(i, c).Value
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        Texto = Format$(v, "dd\/mm\/yyyy")
    Else
        Texto = Trim$(CStr(v))
    End If
Fim:
End Function

' Data de entrega da obra (Empty se não houver).
Public Function DataEntrega(ByVal codObra As String) As Variant
    Dim tb As ListObject
    Dim i As Long
    Dim c As Long
    Dim v As Variant
    DataEntrega = Empty
    On Error GoTo Fim
    Set tb = Tabela(T_Obras)
    i = LinhaDe(tb, codObra)
    c = Col(tb, "ENTREGA")
    If i = 0 Or c = 0 Then Exit Function
    v = tb.DataBodyRange.Cells(i, c).Value
    If IsDate(v) Then DataEntrega = CDate(v)
Fim:
End Function

' Quantas colunas do follow up existem (pelo nome) no arquivo de entregues.
Public Function ColunasQueVao() As Long
    Dim tbF As ListObject
    Dim tbI As ListObject
    Dim c As Long
    Set tbF = Tabela(T_Obras)
    Set tbI = Tabela(T_ENTREGA)
    If tbF Is Nothing Or tbI Is Nothing Then Exit Function
    For c = 1 To tbF.ListColumns.Count
        If Col(tbI, tbF.ListColumns(c).Name) > 0 Then ColunasQueVao = ColunasQueVao + 1
    Next c
End Function

' Colunas do follow up que não existem no arquivo ("" = nenhuma).
Public Function ColunasFora() As String
    Dim tbF As ListObject
    Dim tbI As ListObject
    Dim c As Long
    Set tbF = Tabela(T_Obras)
    Set tbI = Tabela(T_ENTREGA)
    If tbF Is Nothing Or tbI Is Nothing Then Exit Function
    For c = 1 To tbF.ListColumns.Count
        If Col(tbI, tbF.ListColumns(c).Name) = 0 Then
            If Len(ColunasFora) > 0 Then ColunasFora = ColunasFora & ", "
            ColunasFora = ColunasFora & tbF.ListColumns(c).Name
        End If
    Next c
End Function

Public Function TemArquivo() As Boolean
    TemArquivo = Not (Tabela(T_ENTREGA) Is Nothing)
End Function

' Quantas entradas com data ("dd/mm/aaaa ...") a HISTÓRICO da obra tem.
Public Function QtdObservacoes(ByVal codObra As String) As Long
    Dim s As String
    Dim partes() As String
    Dim i As Long
    s = Replace(Texto(codObra, "HISTÓRICO"), vbCr, "")
    If Len(s) = 0 Then Exit Function
    partes = Split(s, vbLf)
    For i = LBound(partes) To UBound(partes)
        If Trim$(partes(i)) Like "##/##/####*" Then QtdObservacoes = QtdObservacoes + 1
    Next i
End Function

' A obra tem marcador "para analisar" aberto?
Public Function TemMarcador(ByVal codObra As String) As Boolean
    TemMarcador = (LinhasMarcador(codObra, False) > 0)
End Function

' Conta (e, se concluir = True, conclui) os marcadores abertos da obra.
Private Function LinhasMarcador(ByVal codObra As String, ByVal fechar As Boolean) As Long
    Dim tb As ListObject
    Dim i As Long
    Dim cA As Long
    Dim cS As Long
    Dim cP As Long
    Dim cE As Long
    On Error GoTo Fim
    Set tb = Tabela(T_MARC)
    If tb Is Nothing Then Exit Function
    If tb.DataBodyRange Is Nothing Then Exit Function
    cA = Col(tb, "CÓDIGO")
    cS = Col(tb, "SITUAÇÃO")
    cP = Col(tb, "CONCLUÍDO POR")
    cE = Col(tb, "CONCLUÍDO EM")
    If cA = 0 Or cS = 0 Then Exit Function
    For i = 1 To tb.ListRows.Count
        If Chave(tb.DataBodyRange.Cells(i, cA).Value) = Trim$(codObra) Then
            If UCase$(Chave(tb.DataBodyRange.Cells(i, cS).Value)) = "ABERTA" Then
                LinhasMarcador = LinhasMarcador + 1
                If fechar Then
                    tb.DataBodyRange.Cells(i, cS).Value = "CONCLUÍDA"
                    If cP > 0 Then tb.DataBodyRange.Cells(i, cP).Value = Environ$("Username")
                    If cE > 0 Then
                        tb.DataBodyRange.Cells(i, cE).NumberFormat = "dd/mm/yyyy hh:mm"
                        tb.DataBodyRange.Cells(i, cE).Value = Now
                    End If
                End If
            End If
        End If
    Next i
Fim:
End Function

' CÓDIGO das obras com SITUAÇÃO = ENTREGUE (botão em lote da aba OBRAS).
Public Function CodObraEntregues() As Collection
    Dim tb As ListObject
    Dim res As Collection
    Dim i As Long
    Dim cS As Long
    Dim cA As Long
    Set res = New Collection
    Set CodObraEntregues = res
    Set tb = Tabela(T_Obras)
    If tb Is Nothing Then Exit Function
    If tb.ListRows.Count = 0 Then Exit Function
    cS = Col(tb, "SITUAÇÃO")
    cA = Col(tb, "CÓDIGO")
    If cS = 0 Or cA = 0 Then Exit Function
    For i = 1 To tb.ListRows.Count
        If UCase$(Chave(tb.DataBodyRange.Cells(i, cS).Value)) = "ENTREGUE" Then
            If Len(Chave(tb.DataBodyRange.Cells(i, cA).Value)) > 0 Then res.Add Chave(tb.DataBodyRange.Cells(i, cA).Value)
        End If
    Next i
End Function

'-------------------------------------------------------------------------------
' Mover para Entregues
'-------------------------------------------------------------------------------
Public Function Entregar(ByVal codObra As String) As Boolean
    Dim tbF As ListObject
    Dim tbI As ListObject
    Dim i As Long
    Dim etapa As String
    Dim ok As Boolean

    On Error GoTo Falha
    mErro = ""
    mAvisoFiltro = ""
    codObra = Trim$(codObra)
    Set tbF = Tabela(T_Obras)
    Set tbI = Tabela(T_ENTREGA)
    If tbI Is Nothing Then
        mErro = "A tabela de entregues (" & T_ENTREGA & ") não foi encontrada."
        Exit Function
    End If
    If LinhaDe(tbF, codObra) = 0 Then
        mErro = "Obra não encontrada no follow up."
        Exit Function
    End If

    Preparar
    etapa = "limpeza dos filtros"
    Set mFiltros = New Collection
    Set mSlicers = New Collection
    GuardarFiltros tbF
    GuardarFiltros tbI
    i = LinhaDe(tbF, codObra)

    etapa = "cópia para ENTREGUES"
    CopiarPorNome tbF, i, tbI
    If LinhaDe(tbI, codObra) = 0 Then Err.Raise vbObjectError + 610, , "a obra não apareceu no arquivo de entregues"

    etapa = "registro no histórico de alterações"
    Registrar tbF, i, "MOVIMENTAÇÃO", "OBRAS", "ENTREGUES"

    etapa = "marcador da obra"
    LinhasMarcador codObra, True

    etapa = "saída da tb_Obras"
    tbF.ListRows(i).Delete
    If LinhaDe(tbF, codObra) > 0 Then
        mErro = "A obra foi gravada em Entregues, mas não saiu do follow up. Apague a linha do follow up manualmente."
        GoTo Fim
    End If
    ok = True

Fim:
    On Error Resume Next
    RestaurarFiltros
    Concluir
    Entregar = ok
    Exit Function
Falha:
    mErro = "A movimentação parou na etapa " & etapa & ": " & Err.Description
    Resume Fim
End Function

' Linha i do follow up para uma linha nova do arquivo, coluna a coluna pelo nome.
Private Sub CopiarPorNome(ByVal origem As ListObject, ByVal i As Long, ByVal destino As ListObject)
    Dim nova As ListRow
    Dim c As Long
    Dim d As Long
    Dim cel As Range
    Dim nome As String

    Set nova = LinhaLivre(destino)
    For c = 1 To origem.ListColumns.Count
        nome = origem.ListColumns(c).Name
        d = Col(destino, nome)
        If d > 0 Then
            Set cel = nova.Range.Cells(1, d)
            If Not cel.HasFormula Then
                If Norm(nome) = "CÓDIGO" Then
                    cel.NumberFormat = "@"
                    cel.Value = Chave(origem.DataBodyRange.Cells(i, c).Value)
                Else
                    cel.NumberFormat = origem.DataBodyRange.Cells(i, c).NumberFormat
                    cel.Value = Seguro(origem.DataBodyRange.Cells(i, c).Value)
                End If
            End If
        End If
    Next c
End Sub

' Texto iniciado por = + - @ não pode virar fórmula.
Private Function Seguro(ByVal v As Variant) As Variant
    If VarType(v) = vbString Then
        If Len(v) > 0 Then
            If InStr(1, "=+-@", Left$(v, 1), vbBinaryCompare) > 0 Then v = "'" & v
        End If
    End If
    Seguro = v
End Function

Private Function LinhaLivre(ByVal tb As ListObject) As ListRow
    If tb.ListRows.Count = 1 Then
        If Application.WorksheetFunction.CountA(tb.ListRows(1).Range) = 0 Then
            Set LinhaLivre = tb.ListRows(1)
            Exit Function
        End If
    End If
    Set LinhaLivre = tb.ListRows.Add
End Function

'-------------------------------------------------------------------------------
' Histórico de alterações
'-------------------------------------------------------------------------------
Private Sub Registrar(ByVal tbOrigem As ListObject, ByVal i As Long, ByVal campo As String, ByVal anterior As String, ByVal novo As String)
    Dim tbH As ListObject
    Dim lr As ListRow
    Dim nomes As Variant
    Dim k As Long
    Dim c As Long
    Dim o As Long

    On Error GoTo Fim
    Set tbH = Tabela(T_HIST)
    If tbH Is Nothing Then Exit Sub
    Set lr = LinhaLivre(tbH)
    nomes = Array("CÓDIGO", "NOME", "CLIENTE", "UF", "CIDADE")
    For k = 0 To 4
        c = Col(tbH, CStr(nomes(k)))
        o = Col(tbOrigem, CStr(nomes(k)))
        If c > 0 And o > 0 Then
            If k = 0 Then lr.Range.Cells(1, c).NumberFormat = "@"
            lr.Range.Cells(1, c).Value = Chave(tbOrigem.DataBodyRange.Cells(i, o).Value)
        End If
    Next k
    EscreverHist tbH, lr, "CAMPO", campo
    EscreverHist tbH, lr, "ANTES", anterior
    EscreverHist tbH, lr, "DEPOIS", novo
    EscreverHist tbH, lr, "USUÁRIO", Environ$("Username")
    c = Col(tbH, "DATA/HORA")
    If c > 0 Then lr.Range.Cells(1, c).Value = Now
Fim:
End Sub

Private Sub EscreverHist(ByVal tbH As ListObject, ByVal lr As ListRow, ByVal coluna As String, ByVal valor As String)
    Dim c As Long
    c = Col(tbH, coluna)
    If c > 0 Then
        lr.Range.Cells(1, c).NumberFormat = "@"
        lr.Range.Cells(1, c).Value = valor
    End If
End Sub

'-------------------------------------------------------------------------------
' Filtros e segmentações
'-------------------------------------------------------------------------------
Private Sub Preparar()
    mEventos = Application.EnableEvents
    mTela = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Sub

Private Sub Concluir()
    On Error Resume Next
    Application.EnableEvents = mEventos
    Application.ScreenUpdating = mTela
    Application.EnableEvents = True
End Sub

' Guarda filtros e segmentações da tabela e deixa todas as linhas visíveis.
Private Sub GuardarFiltros(ByVal tb As ListObject)
    Dim sc As Object
    Dim it As Object
    Dim sel As String
    Dim i As Long
    Dim op As Long
    Dim c1 As Variant
    Dim c2 As Variant
    Dim nomeTabela As String
    Dim filtrado As Boolean
    Dim ligado As Boolean

    On Error Resume Next
    ' segmentações (slicers) ligadas a esta tabela
    ' (o nome da tabela vai para uma variável antes do If: com Resume Next,
    ' um erro dentro da condição faria o VBA entrar no bloco)
    For Each sc In ThisWorkbook.SlicerCaches
        nomeTabela = ""
        filtrado = False
        nomeTabela = sc.ListObject.Name
        filtrado = Not sc.FilterCleared
        Err.Clear
        If nomeTabela = tb.Name And filtrado Then
            sel = "|"
            For Each it In sc.SlicerItems
                If it.Selected Then sel = sel & it.Name & "|"
            Next it
            mSlicers.Add Array(sc, sel)
            sc.ClearManualFilter
        End If
        Err.Clear
    Next sc

    ' filtros comuns
    ligado = False
    ligado = tb.ShowAutoFilter
    If ligado Then ligado = tb.AutoFilter.FilterMode
    Err.Clear
    If ligado Then
        For i = 1 To tb.AutoFilter.Filters.Count
            filtrado = False
            filtrado = tb.AutoFilter.Filters(i).On
            Err.Clear
            If filtrado Then
                op = 0
                c1 = Empty
                c2 = Empty
                op = tb.AutoFilter.Filters(i).Operator
                c1 = tb.AutoFilter.Filters(i).Criteria1
                If op = xlAnd Or op = xlOr Then c2 = tb.AutoFilter.Filters(i).Criteria2
                If Not IsEmpty(c1) Then mFiltros.Add Array(tb, i, op, c1, c2)
            End If
            Err.Clear
        Next i
        tb.AutoFilter.ShowAllData
    End If
    Err.Clear
End Sub

' Reaplica o que foi guardado; o que não der para refazer vira aviso.
Private Sub RestaurarFiltros()
    Dim item As Variant
    Dim tb As ListObject
    Dim sc As Object
    Dim it As Object
    Dim sel As String
    Dim falhou As Boolean

    On Error Resume Next
    If Not mFiltros Is Nothing Then
        For Each item In mFiltros
            Set tb = item(0)
            Err.Clear
            Select Case item(2)
                Case 0
                    tb.Range.AutoFilter Field:=item(1), Criteria1:=item(3)
                Case xlAnd, xlOr
                    tb.Range.AutoFilter Field:=item(1), Criteria1:=item(3), Operator:=item(2), Criteria2:=item(4)
                Case Else
                    tb.Range.AutoFilter Field:=item(1), Criteria1:=item(3), Operator:=item(2)
            End Select
            If Err.Number <> 0 Then falhou = True
        Next item
    End If

    If Not mSlicers Is Nothing Then
        For Each item In mSlicers
            Set sc = item(0)
            sel = item(1)
            Err.Clear
            For Each it In sc.SlicerItems
                If InStr(1, sel, "|" & it.Name & "|", vbBinaryCompare) > 0 Then it.Selected = True
            Next it
            For Each it In sc.SlicerItems
                If InStr(1, sel, "|" & it.Name & "|", vbBinaryCompare) = 0 Then it.Selected = False
            Next it
            If Err.Number <> 0 Then falhou = True
        Next item
    End If

    If falhou Then mAvisoFiltro = "Algum filtro ou segmentação não pôde ser refeito; confira os filtros da aba."
    Set mFiltros = Nothing
    Set mSlicers = Nothing
End Sub

