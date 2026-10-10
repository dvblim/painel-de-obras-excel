Option Explicit
'===============================================================================
' cMarcadores - leitura e gravação dos marcadores (aba oculta MARCADORES)
'
' Um marcador ABERTO por obra para cada pessoa: o do time (todos veem) ou
' o "só eu" de quem marcou. Concluir não apaga: a linha fica com SITUAÇÃO
' CONCLUÍDA, quem concluiu e quando.
'===============================================================================

Private Const T_MARC As String = "tb_Marcadores"
Private Const ABA_MARC As String = "MARCADORES"
Private Const T_Obras As String = "tb_Obras"
Private Const SIT_ABERTA As String = "ABERTA"
Private Const SIT_CONCLUIDA As String = "CONCLUÍDA"
Private Const VIS_TIME As String = "TIME"
Private Const VIS_SO_EU As String = "SÓ QUEM MARCOU"

Private mErro As String
Private mCache As String
Private mTemCache As Boolean

Public Property Get UltimoErro() As String
    UltimoErro = mErro
End Property

Public Function Usuario() As String
    Usuario = Environ$("Username")
    If Len(Usuario) = 0 Then Usuario = Application.UserName
End Function

'-------------------------------------------------------------------------------
' Tabela
'-------------------------------------------------------------------------------
Private Function Cabecalhos() As Variant
    Cabecalhos = Array("CÓDIGO", "OBRA", "O QUE ANALISAR", "REVISAR EM", "QUEM VÊ", _
                       "MARCADO POR", "MARCADO EM", "SITUAÇÃO", "CONCLUÍDO POR", "CONCLUÍDO EM")
End Function

Private Function TabelaPorNome(ByVal nome As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        On Error Resume Next
        Set lo = ws.ListObjects(nome)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set TabelaPorNome = lo
            Exit Function
        End If
    Next ws
End Function

' Tabela dos marcadores; cria a aba oculta na primeira vez quando criar = True.
Public Function Tabela(ByVal criar As Boolean) As ListObject
    Dim tb As ListObject
    Dim ws As Worksheet
    Dim cab As Variant
    Dim i As Long
    Dim ativa As Object

    Set tb = TabelaPorNome(T_MARC)
    If tb Is Nothing And criar Then
        Set ativa = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = ABA_MARC
        cab = Cabecalhos()
        For i = 0 To UBound(cab)
            ws.Cells(1, i + 1).Value = cab(i)
        Next i
        Set tb = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(cab) + 1)), , xlYes)
        tb.Name = T_MARC
        On Error Resume Next
        tb.TableStyle = "TableStyleMedium2"
        tb.ListColumns("CÓDIGO").DataBodyRange.NumberFormat = "@"
        tb.ListColumns("REVISAR EM").DataBodyRange.NumberFormat = "dd/mm/yyyy"
        tb.ListColumns("MARCADO EM").DataBodyRange.NumberFormat = "dd/mm/yyyy hh:mm"
        tb.ListColumns("CONCLUÍDO EM").DataBodyRange.NumberFormat = "dd/mm/yyyy hh:mm"
        ws.Columns("A:J").ColumnWidth = 16
        ws.Columns("C").ColumnWidth = 50
        ws.Visible = xlSheetHidden
        If Not ativa Is Nothing Then ativa.Activate
        On Error GoTo 0
    End If
    Set Tabela = tb
End Function

Private Function Norm(ByVal s As String) As String
    Norm = UCase$(Trim$(Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), Chr$(160), " ")))
End Function

Private Function Col(ByVal tb As ListObject, ByVal nome As String) As Long
    Dim i As Long
    If tb Is Nothing Then Exit Function
    For i = 1 To tb.ListColumns.Count
        If Norm(tb.ListColumns(i).Name) = Norm(nome) Then
            Col = i
            Exit Function
        End If
    Next i
End Function

Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    Txt = Trim$(Replace(CStr(v), Chr$(160), " "))
End Function

Private Function Cel(ByVal tb As ListObject, ByVal i As Long, ByVal nome As String) As Variant
    Dim c As Long
    c = Col(tb, nome)
    If c = 0 Then
        Cel = Empty
    Else
        Cel = tb.DataBodyRange.Cells(i, c).Value
    End If
End Function

Private Sub Gravar(ByVal tb As ListObject, ByVal i As Long, ByVal nome As String, ByVal valor As Variant, Optional ByVal formato As String = "")
    Dim c As Long
    c = Col(tb, nome)
    If c = 0 Then Exit Sub
    If Len(formato) > 0 Then tb.DataBodyRange.Cells(i, c).NumberFormat = formato
    If VarType(valor) = vbString Then
        If Len(valor) > 0 Then
            If InStr(1, "=+-@", Left$(valor, 1), vbBinaryCompare) > 0 Then valor = "'" & valor
        End If
    End If
    tb.DataBodyRange.Cells(i, c).Value = valor
End Sub

' A linha é um marcador aberto que o usuário atual pode ver?
Private Function Visivel(ByVal tb As ListObject, ByVal i As Long) As Boolean
    If UCase$(Txt(Cel(tb, i, "SITUAÇÃO"))) <> SIT_ABERTA Then Exit Function
    If UCase$(Txt(Cel(tb, i, "QUEM VÊ"))) = VIS_SO_EU Then
        Visivel = (UCase$(Txt(Cel(tb, i, "MARCADO POR"))) = UCase$(Usuario()))
    Else
        Visivel = True
    End If
End Function

' Linha do marcador aberto da obra visível para o usuário (0 = não tem).
Public Function LinhaAberta(ByVal codObra As String) As Long
    Dim tb As ListObject
    Dim i As Long
    Dim cA As Long
    Set tb = Tabela(False)
    If tb Is Nothing Then Exit Function
    If tb.DataBodyRange Is Nothing Then Exit Function
    cA = Col(tb, "CÓDIGO")
    If cA = 0 Then Exit Function
    codObra = Trim$(codObra)
    If Len(codObra) = 0 Then Exit Function
    For i = 1 To tb.ListRows.Count
        If Txt(tb.DataBodyRange.Cells(i, cA).Value) = codObra Then
            If Visivel(tb, i) Then
                LinhaAberta = i
                Exit Function
            End If
        End If
    Next i
End Function

Public Function Marcada(ByVal codObra As String) As Boolean
    If Not mTemCache Then MontarCache
    Marcada = (InStr(1, mCache, "|" & Trim$(codObra) & "|", vbBinaryCompare) > 0)
End Function

Private Sub MontarCache()
    Dim itens As Variant
    Dim i As Long
    mCache = "|"
    itens = Lista(False)
    If IsArray(itens) Then
        For i = LBound(itens, 1) To UBound(itens, 1)
            mCache = mCache & itens(i, 0) & "|"
        Next i
    End If
    mTemCache = True
End Sub

Public Sub Recarregar()
    mTemCache = False
End Sub

' Dados do marcador aberto da obra (False = obra sem marcador).
Public Function Ler(ByVal codObra As String, ByRef texto As String, ByRef revisar As Variant, _
                    ByRef soQuemMarcou As Boolean, ByRef por As String, ByRef em As Variant) As Boolean
    Dim tb As ListObject
    Dim i As Long
    texto = ""
    revisar = Empty
    soQuemMarcou = False
    por = ""
    em = Empty
    i = LinhaAberta(codObra)
    If i = 0 Then Exit Function
    Set tb = Tabela(False)
    texto = Txt(Cel(tb, i, "O QUE ANALISAR"))
    revisar = Cel(tb, i, "REVISAR EM")
    If Not IsDate(revisar) Then revisar = Empty
    soQuemMarcou = (UCase$(Txt(Cel(tb, i, "QUEM VÊ"))) = VIS_SO_EU)
    por = Txt(Cel(tb, i, "MARCADO POR"))
    em = Cel(tb, i, "MARCADO EM")
    Ler = True
End Function

'-------------------------------------------------------------------------------
' Gravação
'-------------------------------------------------------------------------------
' Cria ou atualiza o marcador aberto da obra. revisar = Empty para sem data.
Public Function Salvar(ByVal codObra As String, ByVal obra As String, ByVal texto As String, _
                       ByVal revisar As Variant, ByVal soQuemMarcou As Boolean) As Boolean
    Dim tb As ListObject
    Dim i As Long
    Dim nova As Boolean
    Dim ev As Boolean

    On Error GoTo Falha
    mErro = ""
    codObra = Trim$(codObra)
    If Len(codObra) = 0 Then
        mErro = "Obra sem CÓDIGO."
        Exit Function
    End If
    ev = Application.EnableEvents
    Application.EnableEvents = False
    Set tb = Tabela(True)
    i = LinhaAberta(codObra)
    If i = 0 Then
        i = LinhaLivre(tb)
        nova = True
    End If

    Gravar tb, i, "CÓDIGO", codObra, "@"
    Gravar tb, i, "OBRA", obra
    Gravar tb, i, "O QUE ANALISAR", texto
    If IsDate(revisar) Then
        Gravar tb, i, "REVISAR EM", CDate(Int(CDbl(CDate(revisar)))), "dd/mm/yyyy"
    Else
        Gravar tb, i, "REVISAR EM", Empty
    End If
    If soQuemMarcou Then
        Gravar tb, i, "QUEM VÊ", VIS_SO_EU
    Else
        Gravar tb, i, "QUEM VÊ", VIS_TIME
    End If
    If nova Then
        Gravar tb, i, "MARCADO POR", Usuario()
        Gravar tb, i, "MARCADO EM", Now, "dd/mm/yyyy hh:mm"
        Gravar tb, i, "SITUAÇÃO", SIT_ABERTA
    End If
    Application.EnableEvents = ev
    mTemCache = False
    Salvar = True
    Exit Function
Falha:
    mErro = "Não foi possível gravar o marcador: " & Err.Description
    Application.EnableEvents = True
End Function

' Tira a marcação: a linha fica como CONCLUÍDA, com quem e quando.
Public Function Concluir(ByVal codObra As String) As Boolean
    Dim tb As ListObject
    Dim i As Long
    Dim ev As Boolean

    On Error GoTo Falha
    mErro = ""
    i = LinhaAberta(codObra)
    If i = 0 Then
        mErro = "Esta obra não tem marcador aberto."
        Exit Function
    End If
    ev = Application.EnableEvents
    Application.EnableEvents = False
    Set tb = Tabela(False)
    Gravar tb, i, "SITUAÇÃO", SIT_CONCLUIDA
    Gravar tb, i, "CONCLUÍDO POR", Usuario()
    Gravar tb, i, "CONCLUÍDO EM", Now, "dd/mm/yyyy hh:mm"
    Application.EnableEvents = ev
    mTemCache = False
    Concluir = True
    Exit Function
Falha:
    mErro = "Não foi possível concluir: " & Err.Description
    Application.EnableEvents = True
End Function

Private Function LinhaLivre(ByVal tb As ListObject) As Long
    If tb.ListRows.Count = 1 Then
        If Application.WorksheetFunction.CountA(tb.ListRows(1).Range) = 0 Then
            LinhaLivre = 1
            Exit Function
        End If
    End If
    tb.ListRows.Add
    LinhaLivre = tb.ListRows.Count
End Function

'-------------------------------------------------------------------------------
' Lista e resumo
'-------------------------------------------------------------------------------
' Marcadores abertos visíveis, por prazo (atrasadas primeiro, sem data no fim):
' (i, 0) CÓDIGO  (i, 1) OBRA  (i, 2) O QUE ANALISAR  (i, 3) REVISAR EM
' (i, 4) MARCADO POR  (i, 5) MARCADO EM  (i, 6) True se só quem marcou vê
Public Function Lista(ByVal soMinhas As Boolean) As Variant
    Dim tb As ListObject
    Dim i As Long
    Dim n As Long
    Dim k As Long
    Dim j As Long
    Dim r() As Variant
    Dim ordem() As Double
    Dim saida() As Variant
    Dim idx() As Long
    Dim t As Long
    Dim v As Variant
    Dim eu As String

    Set tb = Tabela(False)
    If tb Is Nothing Then Exit Function
    If tb.DataBodyRange Is Nothing Then Exit Function
    eu = UCase$(Usuario())
    ReDim r(1 To tb.ListRows.Count, 0 To 6)
    ReDim ordem(1 To tb.ListRows.Count)
    For i = 1 To tb.ListRows.Count
        If Len(Txt(Cel(tb, i, "CÓDIGO"))) > 0 Then
            If Visivel(tb, i) Then
                If (Not soMinhas) Or UCase$(Txt(Cel(tb, i, "MARCADO POR"))) = eu Then
                    n = n + 1
                    r(n, 0) = Txt(Cel(tb, i, "CÓDIGO"))
                    r(n, 1) = Txt(Cel(tb, i, "OBRA"))
                    r(n, 2) = Txt(Cel(tb, i, "O QUE ANALISAR"))
                    v = Cel(tb, i, "REVISAR EM")
                    If IsDate(v) Then
                        r(n, 3) = CDate(v)
                        ordem(n) = CDbl(CDate(v))
                    Else
                        r(n, 3) = Empty
                        ordem(n) = 2958465
                    End If
                    r(n, 4) = Txt(Cel(tb, i, "MARCADO POR"))
                    r(n, 5) = Cel(tb, i, "MARCADO EM")
                    r(n, 6) = (UCase$(Txt(Cel(tb, i, "QUEM VÊ"))) = VIS_SO_EU)
                End If
            End If
        End If
    Next i
    If n = 0 Then Exit Function

    ' ordem estável por data de revisão
    ReDim idx(1 To n)
    For i = 1 To n
        idx(i) = i
    Next i
    For i = 2 To n
        t = idx(i)
        j = i - 1
        Do While j >= 1
            If ordem(idx(j)) <= ordem(t) Then Exit Do
            idx(j + 1) = idx(j)
            j = j - 1
        Loop
        idx(j + 1) = t
    Next i

    ReDim saida(0 To n - 1, 0 To 6)
    For i = 1 To n
        For k = 0 To 6
            saida(i - 1, k) = r(idx(i), k)
        Next k
    Next i
    Lista = saida
End Function

Public Sub Resumo(ByRef total As Long, ByRef hoje As Long, ByRef atrasadas As Long)
    Dim itens As Variant
    Dim i As Long
    total = 0
    hoje = 0
    atrasadas = 0
    itens = Lista(False)
    If Not IsArray(itens) Then Exit Sub
    For i = LBound(itens, 1) To UBound(itens, 1)
        total = total + 1
        Select Case TipoPrazo(itens(i, 3))
            Case PZ_HOJE
                hoje = hoje + 1
            Case PZ_ATRASADA
                atrasadas = atrasadas + 1
        End Select
    Next i
End Sub

'-------------------------------------------------------------------------------
' follow up
'-------------------------------------------------------------------------------
Public Function LinhaNoFollowUp(ByVal codObra As String) As Long
    LinhaNoFollowUp = LinhaEm(T_Obras, codObra)
End Function

Private Function LinhaEm(ByVal nomeTabela As String, ByVal codObra As String) As Long
    Dim tb As ListObject
    Dim i As Long
    Dim c As Long
    Set tb = TabelaPorNome(nomeTabela)
    If tb Is Nothing Then Exit Function
    If tb.DataBodyRange Is Nothing Then Exit Function
    c = Col(tb, "CÓDIGO")
    If c = 0 Then Exit Function
    For i = 1 To tb.ListRows.Count
        If Txt(tb.DataBodyRange.Cells(i, c).Value) = Trim$(codObra) Then
            LinhaEm = i
            Exit Function
        End If
    Next i
End Function
