Attribute VB_Name = "mInstaladorPainelObras"
Option Explicit
'===============================================================================
' INSTALADOR - Follow up Projetos 2028 (Painel de Obras completo)
'
' Instala, numa planilha que tenha a tabela tb_Obras, o Painel de Obras inteiro:
'   - painel da obra (datas, observacoes, ata no Word, nova obra)
'   - calendario com feriados e fins de semana (aba FERIADOS)
'   - marcador "para analisar" (aba oculta MARCADORES)
'   - NOVA OBRA conferida na base, dentro do painel
'   - OBRA ENTREGUE: move para ENTREGUES dentro do painel
'
' Tabelas usadas (formato de tabela do Excel, Ctrl+T):
'   tb_Obras (obrigatoria), tb_Cronograma, tb_Carteira,
'   tb_Entregues, tb_Alteracoes
'
' Como usar:
'   1. Salve a planilha como .xlsm (Pasta de Trabalho Habilitada para Macro).
'   2. Em Arquivo > Opcoes > Central de Confiabilidade > Configuracoes de
'      Macro, marque "Confiar no acesso ao modelo de objeto do projeto do VBA".
'   3. Importe este arquivo no editor VBA (Arquivo > Importar arquivo).
'   4. Execute InstalarFollowUp e salve a planilha.
'
' Pode ser executado de novo: substitui o codigo do painel e mantem os dados
' (as abas FERIADOS e MARCADORES nao sao apagadas).
'===============================================================================

Private Const TITULO As String = "Instalador - Follow up Projetos"

Public Sub InstalarFollowUp()
    Dim proj As Object
    Dim etapa As String
    Dim pasta As String
    Dim faltam As String
    Dim nome As Variant
    Dim rAba As String
    Dim rBotoes As String
    Dim rMarc As String

    On Error Resume Next
    Set proj = ThisWorkbook.VBProject
    On Error GoTo 0

    If proj Is Nothing Then
        MsgBox "O Excel bloqueou o acesso ao projeto VBA." & vbCrLf & vbCrLf & "Ative em Arquivo > Op" & ChrW(231) & "" & ChrW(245) & "es > Central de Confiabilidade > Configura" & ChrW(231) & "" & ChrW(245) & "es de Macro a op" & ChrW(231) & "" & ChrW(227) & "o ""Confiar no acesso ao modelo de objeto do projeto do VBA"" e execute de novo.", vbExclamation, TITULO
        Exit Sub
    End If
    If proj.Protection = 1 Then
        MsgBox "O projeto VBA est" & ChrW(225) & " bloqueado. Abra o projeto no editor (Alt+F11), informe a senha e execute de novo.", vbExclamation, TITULO
        Exit Sub
    End If
    If LCase$(Right$(ThisWorkbook.Name, 5)) = ".xlsx" Then
        MsgBox "Salve a planilha como .xlsm (Pasta de Trabalho Habilitada para Macro) antes de instalar: um arquivo .xlsx perde as macros.", vbExclamation, TITULO
        Exit Sub
    End If
    For Each nome In Array("tb_Obras", "tb_Alteracoes")
        If TabelaPorNome(CStr(nome)) Is Nothing Then faltam = faltam & vbCrLf & "- " & nome
    Next nome
    If Len(faltam) > 0 Then
        MsgBox "N" & ChrW(227) & "o encontrei estas tabelas (formato de tabela do Excel):" & faltam & vbCrLf & vbCrLf & "Converta os dados em tabela (Ctrl+T) com esses nomes e execute de novo.", vbExclamation, TITULO
        Exit Sub
    End If

    On Error GoTo Falha
    etapa = "fechamento dos pain" & ChrW(233) & "is abertos"
    FecharFormularios
    etapa = "c" & ChrW(243) & "pia de seguran" & ChrW(231) & "a do c" & ChrW(243) & "digo atual"
    pasta = FazerBackup(proj)

    etapa = "remo" & ChrW(231) & "" & ChrW(227) & "o da vers" & ChrW(227) & "o anterior"
    For Each nome In Array("frmPainelObras", "frmMarcador", "frmMarcadas", "mPainelBase", "clsObra", "clsCampoData", _
                           "mFeriados", "cFerClique", "cFerCampo", "cCalendario", _
                           "mMarcador", "cMarcaClique", "cMarcaIcone", "cMarcadores", _
                           "mNovaObra", "cNovaClique", "cNovaObraPainel", _
                           "mEntregues", "cEntregaClique", "cArquivoObras", "cEntregaPainel")
        Descartar proj, CStr(nome)
    Next nome

    etapa = "m" & ChrW(243) & "dulos do painel"
    CriarComponente proj, 1, "mPainelBase", CodigoBase()
    CriarComponente proj, 2, "clsObra", CodigoObra()
    CriarComponente proj, 2, "clsCampoData", CodigoCampoData()
    etapa = "calend" & ChrW(225) & "rio e feriados"
    CriarComponente proj, 1, "mFeriados", CodigoFerModulo()
    CriarComponente proj, 2, "cFerClique", CodigoFerClique()
    CriarComponente proj, 2, "cFerCampo", CodigoFerCampo()
    CriarComponente proj, 2, "cCalendario", CodigoFerCalendario()
    etapa = "marcador"
    CriarComponente proj, 1, "mMarcador", CodigoMarcModulo()
    CriarComponente proj, 2, "cMarcaClique", CodigoMarcClique()
    CriarComponente proj, 2, "cMarcaIcone", CodigoMarcIcone()
    CriarComponente proj, 2, "cMarcadores", CodigoMarcDados()
    CriarFormulario proj, "frmMarcador", "Marcar para analisar", 452, 322, CodigoMarcForm()
    CriarFormulario proj, "frmMarcadas", "Obras marcadas para analisar", 732, 454, CodigoMarcLista()
    etapa = "nova obra"
    CriarComponente proj, 1, "mNovaObra", CodigoNovaModulo()
    CriarComponente proj, 2, "cNovaClique", CodigoNovaClique()
    CriarComponente proj, 2, "cNovaObraPainel", CodigoNovaClasse()
    etapa = "obra entregue"
    CriarComponente proj, 1, "mEntregues", CodigoIaModulo()
    CriarComponente proj, 2, "cEntregaClique", CodigoIaClique()
    CriarComponente proj, 2, "cArquivoObras", CodigoIaArquivo()
    CriarComponente proj, 2, "cEntregaPainel", CodigoIaClasse()
    etapa = "Painel de Obras"
    CriarFormulario proj, "frmPainelObras", "PAINEL DE OBRAS - DADOS DA OBRA", 1098, 550, CodigoPainel()

    etapa = "aba FERIADOS"
    rAba = CriarAbaFeriados()
    etapa = "aba MARCADORES"
    rMarc = CriarMarcadores()
    etapa = "bot" & ChrW(245) & "es da aba OBRAS"
    rBotoes = CriarBotoes()

    MsgBox "Follow up instalado." & vbCrLf & vbCrLf & _
        "- " & rAba & vbCrLf & "- " & rMarc & vbCrLf & "- " & rBotoes & vbCrLf & vbCrLf & _
        "Abra pelo bot" & ChrW(227) & "o ABRIR PAINEL DE OBRAS (ou pela macro AbrirPainelObras)." & vbCrLf & _
        "Salve a planilha." & vbCrLf & vbCrLf & "C" & ChrW(243) & "pia do c" & ChrW(243) & "digo anterior: " & pasta, vbInformation, TITULO
    Exit Sub
Falha:
    MsgBox "A instala" & ChrW(231) & "" & ChrW(227) & "o parou na etapa: " & etapa & "." & vbCrLf & Err.Description & vbCrLf & vbCrLf & "C" & ChrW(243) & "pia do c" & ChrW(243) & "digo anterior: " & pasta, vbExclamation, TITULO
End Sub

'-------------------------------------------------------------------------------
' Apoio
'-------------------------------------------------------------------------------
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

Private Function NormNome(ByVal s As String) As String
    s = Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), Chr$(160), " ")
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    NormNome = UCase$(Trim$(s))
End Function

Private Function ColPorNome(ByVal tb As ListObject, ByVal nome As String) As Long
    Dim i As Long
    For i = 1 To tb.ListColumns.Count
        If NormNome(tb.ListColumns(i).Name) = NormNome(nome) Then
            ColPorNome = i
            Exit Function
        End If
    Next i
End Function

Private Sub FecharFormularios()
    Dim i As Long
    On Error Resume Next
    For i = VBA.UserForms.Count - 1 To 0 Step -1
        Unload VBA.UserForms(i)
    Next i
End Sub

Private Function FazerBackup(ByVal proj As Object) As String
    Dim pasta As String
    Dim comp As Object
    Dim ext As String

    pasta = Environ$("TEMP") & "\Backup_VBA_Obras_" & Format$(Now, "yyyymmdd_hhnnss")
    MkDir pasta
    On Error Resume Next
    For Each comp In proj.VBComponents
        Select Case comp.Type
            Case 1
                ext = ".bas"
            Case 3
                ext = ".frm"
            Case Else
                ext = ".cls"
        End Select
        comp.Export pasta & "\" & comp.Name & ext
    Next comp
    On Error GoTo 0
    FazerBackup = pasta
End Function

' Renomeia antes de remover: o VBE so libera o nome ao fim da execucao.
Private Sub Descartar(ByVal proj As Object, ByVal nome As String)
    Static seq As Long
    Dim comp As Object
    On Error Resume Next
    Set comp = proj.VBComponents(nome)
    On Error GoTo 0
    If comp Is Nothing Then Exit Sub
    seq = seq + 1
    On Error Resume Next
    comp.Name = "zzAntigo" & Format$(Now, "hhnnss") & seq
    On Error GoTo 0
    proj.VBComponents.Remove comp
End Sub

Private Sub CriarComponente(ByVal proj As Object, ByVal tipo As Long, ByVal nome As String, ByVal codigo As String)
    Dim comp As Object
    Set comp = proj.VBComponents.Add(tipo)
    comp.Name = nome
    GravarCodigo comp, codigo
End Sub

Private Sub GravarCodigo(ByVal comp As Object, ByVal codigo As String)
    If comp.CodeModule.CountOfLines > 0 Then comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
    comp.CodeModule.AddFromString codigo
End Sub


' UserForm vazio: a propria janela cria os controles ao abrir.
Private Sub CriarFormulario(ByVal proj As Object, ByVal nome As String, ByVal legenda As String, _
                            ByVal largura As Single, ByVal altura As Single, ByVal codigo As String)
    Dim comp As Object
    Set comp = proj.VBComponents.Add(3)
    comp.Name = nome
    comp.Properties("Caption") = legenda
    comp.Properties("Width") = largura
    comp.Properties("Height") = altura
    comp.Properties("StartUpPosition") = 1
    GravarCodigo comp, codigo
End Sub

Private Function CriarMarcadores() As String
    Dim ok As Variant
    On Error GoTo Falhou
    If Not TabelaPorNome("tb_Marcadores") Is Nothing Then
        CriarMarcadores = "aba MARCADORES j" & ChrW(225) & " existia (marcadores mantidos)"
        Exit Function
    End If
    ok = Application.Run("'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PrepararMarcadores")
    CriarMarcadores = "aba oculta MARCADORES criada"
    Exit Function
Falhou:
    CriarMarcadores = "aba MARCADORES ser" & ChrW(225) & " criada no primeiro marcador salvo"
End Function

' Botoes na linha 1 da aba do follow up: abrir o painel e arquivar entregues em lote.
Private Function CriarBotoes() As String
    Dim tb As ListObject
    Dim ws As Worksheet
    Dim x As Single
    Set tb = TabelaPorNome("tb_Obras")
    Set ws = tb.Parent
    On Error Resume Next
    ws.Shapes("btnPainelObras").Delete
    ws.Shapes("btnArquivarEntrega").Delete
    On Error GoTo Falhou
    If ws.Rows(1).RowHeight < 30 Then ws.Rows(1).RowHeight = 30
    x = tb.HeaderRowRange.Cells(1, 1).Left
    If tb.ListColumns.Count >= 8 Then x = tb.HeaderRowRange.Cells(1, 8).Left
    NovoBotaoAba ws, "btnPainelObras", "ABRIR PAINEL DE OBRAS", x, 160, RGB(30, 41, 59), RGB(251, 191, 36), "AbrirPainelObras"
    NovoBotaoAba ws, "btnArquivarEntrega", "ARQUIVAR ENTREGUES", x + 172, 180, RGB(22, 163, 74), RGB(22, 163, 74), "ArquivarEntregues"
    CriarBotoes = "bot" & ChrW(245) & "es ABRIR PAINEL DE OBRAS e ARQUIVAR ENTREGUES criados na aba " & ws.Name
    Exit Function
Falhou:
    CriarBotoes = "bot" & ChrW(245) & "es n" & ChrW(227) & "o criados (" & Err.Description & "): use as macros AbrirPainelObras e ArquivarEntregues"
End Function

Private Sub NovoBotaoAba(ByVal ws As Worksheet, ByVal nome As String, ByVal texto As String, ByVal x As Single, _
                         ByVal largura As Single, ByVal fundo As Long, ByVal borda As Long, ByVal macro As String)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, x, ws.Range("A1").Top + 2, largura, 26)
    shp.Name = nome
    shp.Placement = 3
    shp.Fill.ForeColor.RGB = fundo
    shp.Line.ForeColor.RGB = borda
    shp.Line.Weight = 2
    shp.TextFrame2.VerticalAnchor = 3
    shp.TextFrame2.TextRange.Text = texto
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = -1
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.OnAction = macro
End Sub

'-------------------------------------------------------------------------------
' Aba FERIADOS (criada uma vez; depois e do time)
'-------------------------------------------------------------------------------
Private Function CriarAbaFeriados() As String
    Dim tb As ListObject
    Dim ws As Worksheet
    Dim linhas() As String
    Dim partes() As String
    Dim dados() As Variant
    Dim cab As Variant
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim ativa As Object

    Set tb = TabelaPorNome("tb_Feriados")
    If Not tb Is Nothing Then
        CriarAbaFeriados = "aba FERIADOS j" & ChrW(225) & " existia: mantida como est" & ChrW(225) & " (" & tb.ListRows.Count & " linha(s))"
        Exit Function
    End If

    linhas = Split(Replace(CodigoSemente(), vbCr, ""), vbLf)
    ReDim dados(1 To UBound(linhas) + 1, 1 To 6)
    For i = LBound(linhas) To UBound(linhas)
        If Len(Trim$(linhas(i))) > 0 Then
            partes = Split(linhas(i), ";")
            n = n + 1
            For j = 0 To 5
                If j <= UBound(partes) Then dados(n, j + 1) = partes(j)
            Next j
        End If
    Next i

    Set ativa = ActiveSheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "FERIADOS"
    cab = Array("QUANDO", "NOME", "ABRANG" & ChrW(202) & "NCIA", "UF", "CIDADE", "HIST" & ChrW(211) & "RICO")
    For j = 0 To 5
        ws.Cells(1, j + 1).Value = cab(j)
    Next j
    ws.Columns("A:F").NumberFormat = "@"
    If n > 0 Then ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, 6)).Value = dados
    Set tb = ws.ListObjects.Add(1, ws.Range(ws.Cells(1, 1), ws.Cells(Application.WorksheetFunction.Max(n + 1, 2), 6)), , 1)
    tb.Name = "tb_Feriados"
    On Error Resume Next
    tb.TableStyle = "TableStyleMedium2"
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 46
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 6
    ws.Columns("E").ColumnWidth = 26
    ws.Columns("F").ColumnWidth = 60
    ws.Range("H1").Value = "Como preencher QUANDO:"
    ws.Range("H2").Value = "dd/mm = todo ano  |  dd/mm/aaaa = s" & ChrW(243) & " naquele ano"
    ws.Range("H3").Value = "PASCOA+n ou PASCOA-n, SEXTA-FEIRA SANTA, CORPUS CHRISTI, CARNAVAL, CARNAVAL TERCA, QUARTA-FEIRA DE CINZAS"
    ws.Range("H4").Value = "ABRANG" & ChrW(202) & "NCIA: ESTADUAL ou MUNICIPAL (bloqueiam a data) ou PONTO FACULTATIVO (s" & ChrW(243) & " aparece no calend" & ChrW(225) & "rio)"
    ws.Range("H5").Value = "UF vazia = todas  |  CIDADE vazia = o estado todo  |  feriados nacionais s" & ChrW(227) & "o calculados pelo painel"
    ws.Range("H1").Font.Bold = True
    If Not ativa Is Nothing Then ativa.Activate
    On Error GoTo 0
    CriarAbaFeriados = "aba FERIADOS criada com " & n & " feriado(s) estaduais e municipais"
End Function

Private Function CodigoFerModulo() As String
    Dim s As String
    s = s & CodFerModulo1()
    s = s & CodFerModulo2()
    s = s & CodFerModulo3()
    s = s & CodFerModulo4()
    s = s & CodFerModulo5()
    s = s & CodFerModulo6()
    s = s & CodFerModulo7()
    s = s & CodFerModulo8()
    s = s & CodFerModulo9()
    s = s & CodFerModulo10()
    CodigoFerModulo = s
End Function

Private Function CodFerModulo1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' mFeriados - feriados nacionais, estaduais e municipais e dias " & ChrW(250) & "teis" & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "' NACIONAIS: calculados aqui (fixos + Sexta-feira Santa, pela P" & ChrW(225) & "scoa)." & vbCrLf
    s = s & "' PONTOS FACULTATIVOS NACIONAIS: Carnaval (segunda e ter" & ChrW(231) & "a), Quarta-feira de" & vbCrLf
    s = s & "'   Cinzas e Corpus Christi; aparecem no calend" & ChrW(225) & "rio mas n" & ChrW(227) & "o bloqueiam a data." & vbCrLf
    s = s & "' ESTADUAIS E MUNICIPAIS: aba FERIADOS (tabela tb_Feriados), que o time pode" & vbCrLf
    s = s & "'   completar. Coluna QUANDO aceita:" & vbCrLf
    s = s & "'     dd/mm             todo ano" & vbCrLf
    s = s & "'     dd/mm/aaaa        s" & ChrW(243) & " naquele ano" & vbCrLf
    s = s & "'     PASCOA+n / PASCOA-n, SEXTA-FEIRA SANTA, CORPUS CHRISTI, CARNAVAL," & vbCrLf
    s = s & "'     CARNAVAL TERCA, QUARTA-FEIRA DE CINZAS" & vbCrLf
    s = s & "'   ABRANG" & ChrW(202) & "NCIA: ESTADUAL, MUNICIPAL (bloqueiam) ou PONTO FACULTATIVO." & vbCrLf
    s = s & "'   UF vazia = todas; CIDADE vazia = o estado todo." & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Const FR_NAVY As Long = 3877150" & vbCrLf
    s = s & "Public Const FR_AZUL As Long = 8950797" & vbCrLf
    s = s & "Public Const FR_CORAL As Long = 2500316" & vbCrLf
    s = s & "Public Const FR_AMARELO As Long = 2408443" & vbCrLf
    s = s & "Public Const FR_CINZA As Long = 9139300" & vbCrLf
    s = s & "Public Const FR_FUNDO As Long = 16579320" & vbCrLf
    s = s & "Public Const FR_BORDA As Long = 15788258" & vbCrLf
    s = s & "Public Const FR_BRANCO As Long = 16777215" & vbCrLf
    s = s & "Public Const FR_DESTAQUE As Long = 15465471" & vbCrLf
    s = s & "Public Const FR_CORAL_FUNDO As Long = 14869246" & vbCrLf
    s = s & "Public Const FR_FDS As Long = 16381425" & vbCrLf
    s = s & "Public Const FR_FDS_TEXTO As Long = 12100500" & vbCrLf
    s = s & "Public Const FR_AZUL_FUNDO As Long = 16449008" & vbCrLf
    s = s & "Public Const FR_OURO As Long = 423897" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' resultado de InfoDia" & vbCrLf
    s = s & "Public Const DIA_UTIL As Long = 0" & vbCrLf
    s = s & "Public Const DIA_SABADO As Long = 1" & vbCrLf
    s = s & "Public Const DIA_DOMINGO As Long = 2" & vbCrLf
    s = s & "Public Const DIA_FERIADO As Long = 3" & vbCrLf
    s = s & "Public Const DIA_FACULTATIVO As Long = 4" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const T_FER As String = ""tb_Feriados""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mCarregado As Boolean" & vbCrLf
    s = s & "Private mQtd As Long" & vbCrLf
    s = s & "Private mQuando() As String" & vbCrLf
    s = s & "Private mNome() As String" & vbCrLf
    CodFerModulo1 = s
End Function

Private Function CodFerModulo2() As String
    Dim s As String
    s = s & "Private mAbr() As String" & vbCrLf
    s = s & "Private mUF() As String" & vbCrLf
    s = s & "Private mCid() As String" & vbCrLf
    s = s & "Private mFonteTexto As String" & vbCrLf
    s = s & "Private mFonteTitulo As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Tabela FERIADOS" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "Public Sub FeriadosRecarregar()" & vbCrLf
    s = s & "    mCarregado = False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function TabelaFeriados() As ListObject" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Dim lo As ListObject" & vbCrLf
    s = s & "    For Each ws In ThisWorkbook.Worksheets" & vbCrLf
    s = s & "        Set lo = Nothing" & vbCrLf
    s = s & "        On Error Resume Next" & vbCrLf
    s = s & "        Set lo = ws.ListObjects(T_FER)" & vbCrLf
    s = s & "        On Error GoTo 0" & vbCrLf
    s = s & "        If Not lo Is Nothing Then" & vbCrLf
    s = s & "            Set TabelaFeriados = lo" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next ws" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Carregar()" & vbCrLf
    s = s & "    Dim tb As ListObject" & vbCrLf
    s = s & "    Dim v As Variant" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim cQ As Long" & vbCrLf
    s = s & "    Dim cN As Long" & vbCrLf
    s = s & "    Dim cA As Long" & vbCrLf
    s = s & "    Dim cU As Long" & vbCrLf
    s = s & "    Dim cC As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mCarregado = True" & vbCrLf
    s = s & "    mQtd = 0" & vbCrLf
    s = s & "    Set tb = TabelaFeriados()" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    If tb.DataBodyRange Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    cQ = ColFer(tb, ""QUANDO"")" & vbCrLf
    s = s & "    cN = ColFer(tb, ""NOME"")" & vbCrLf
    CodFerModulo2 = s
End Function

Private Function CodFerModulo3() As String
    Dim s As String
    s = s & "    cA = ColFer(tb, ""ABRANG" & ChrW(202) & "NCIA"")" & vbCrLf
    s = s & "    cU = ColFer(tb, ""UF"")" & vbCrLf
    s = s & "    cC = ColFer(tb, ""CIDADE"")" & vbCrLf
    s = s & "    If cQ = 0 Or cN = 0 Or cA = 0 Then Exit Sub" & vbCrLf
    s = s & "    v = tb.DataBodyRange.Value" & vbCrLf
    s = s & "    ReDim mQuando(1 To UBound(v, 1))" & vbCrLf
    s = s & "    ReDim mNome(1 To UBound(v, 1))" & vbCrLf
    s = s & "    ReDim mAbr(1 To UBound(v, 1))" & vbCrLf
    s = s & "    ReDim mUF(1 To UBound(v, 1))" & vbCrLf
    s = s & "    ReDim mCid(1 To UBound(v, 1))" & vbCrLf
    s = s & "    For i = 1 To UBound(v, 1)" & vbCrLf
    s = s & "        If Len(TextoCel(v(i, cQ))) > 0 Then" & vbCrLf
    s = s & "            mQtd = mQtd + 1" & vbCrLf
    s = s & "            If VarType(v(i, cQ)) = vbDate Then" & vbCrLf
    s = s & "                mQuando(mQtd) = Format$(v(i, cQ), ""dd\/mm\/yyyy"")" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                mQuando(mQtd) = UCase$(SemAcento(TextoCel(v(i, cQ))))" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "            mNome(mQtd) = TextoCel(v(i, cN))" & vbCrLf
    s = s & "            mAbr(mQtd) = UCase$(SemAcento(TextoCel(v(i, cA))))" & vbCrLf
    s = s & "            If cU > 0 Then mUF(mQtd) = UCase$(TextoCel(v(i, cU)))" & vbCrLf
    s = s & "            If cC > 0 Then mCid(mQtd) = NormCidade(TextoCel(v(i, cC)))" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ColFer(ByVal tb As ListObject, ByVal nome As String) As Long" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    For i = 1 To tb.ListColumns.Count" & vbCrLf
    s = s & "        If UCase$(SemAcento(Trim$(tb.ListColumns(i).Name))) = UCase$(SemAcento(nome)) Then" & vbCrLf
    s = s & "            ColFer = i" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function TextoCel(ByVal v As Variant) As String" & vbCrLf
    s = s & "    If IsError(v) Then Exit Function" & vbCrLf
    s = s & "    If IsEmpty(v) Then Exit Function" & vbCrLf
    s = s & "    TextoCel = Trim$(Replace(CStr(v), Chr$(160), "" ""))" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function SemAcento(ByVal s As String) As String" & vbCrLf
    s = s & "    Dim de As String" & vbCrLf
    s = s & "    Dim para As String" & vbCrLf
    CodFerModulo3 = s
End Function

Private Function CodFerModulo4() As String
    Dim s As String
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    de = """ & ChrW(193) & "" & ChrW(192) & "" & ChrW(194) & "" & ChrW(195) & "" & ChrW(196) & "" & ChrW(201) & "" & ChrW(200) & "" & ChrW(202) & "" & ChrW(203) & "" & ChrW(205) & "" & ChrW(204) & "" & ChrW(206) & "" & ChrW(207) & "" & ChrW(211) & "" & ChrW(210) & "" & ChrW(212) & "" & ChrW(213) & "" & ChrW(214) & "" & ChrW(218) & "" & ChrW(217) & "" & ChrW(219) & "" & ChrW(220) & "" & ChrW(199) & "" & ChrW(225) & "" & ChrW(224) & "" & ChrW(226) & "" & ChrW(227) & "" & ChrW(228) & "" & ChrW(233) & "" & ChrW(232) & "" & ChrW(234) & "" & ChrW(235) & "" & ChrW(237) & "" & ChrW(236) & "" & ChrW(238) & "" & ChrW(239) & "" & ChrW(243) & "" & ChrW(242) & "" & ChrW(244) & "" & ChrW(245) & "" & ChrW(246) & "" & ChrW(250) & "" & ChrW(249) & "" & ChrW(251) & "" & ChrW(252) & "" & ChrW(231) & """" & vbCrLf
    s = s & "    para = ""AAAAAEEEEIIIIOOOOOUUUUCaaaaaeeeeiiiiooooouuuuc""" & vbCrLf
    s = s & "    For i = 1 To Len(de)" & vbCrLf
    s = s & "        s = Replace(s, Mid$(de, i, 1), Mid$(para, i, 1), , , vbBinaryCompare)" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    SemAcento = s" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function NormCidade(ByVal s As String) As String" & vbCrLf
    s = s & "    s = UCase$(SemAcento(Trim$(Replace(s, Chr$(160), "" ""))))" & vbCrLf
    s = s & "    Do While InStr(1, s, ""  "", vbBinaryCompare) > 0" & vbCrLf
    s = s & "        s = Replace(s, ""  "", "" "")" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "    NormCidade = s" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Datas m" & ChrW(243) & "veis" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "Public Function Pascoa(ByVal ano As Long) As Date" & vbCrLf
    s = s & "    Dim a As Long" & vbCrLf
    s = s & "    Dim b As Long" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Dim d As Long" & vbCrLf
    s = s & "    Dim e As Long" & vbCrLf
    s = s & "    Dim f As Long" & vbCrLf
    s = s & "    Dim g As Long" & vbCrLf
    s = s & "    Dim h As Long" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim k As Long" & vbCrLf
    s = s & "    Dim l As Long" & vbCrLf
    s = s & "    Dim m As Long" & vbCrLf
    s = s & "    Dim mes As Long" & vbCrLf
    s = s & "    Dim dia As Long" & vbCrLf
    s = s & "    a = ano Mod 19" & vbCrLf
    s = s & "    b = ano \ 100" & vbCrLf
    s = s & "    c = ano Mod 100" & vbCrLf
    s = s & "    d = b \ 4" & vbCrLf
    s = s & "    e = b Mod 4" & vbCrLf
    s = s & "    f = (b + 8) \ 25" & vbCrLf
    s = s & "    g = (b - f + 1) \ 3" & vbCrLf
    s = s & "    h = (19 * a + b - d - g + 15) Mod 30" & vbCrLf
    s = s & "    i = c \ 4" & vbCrLf
    s = s & "    k = c Mod 4" & vbCrLf
    CodFerModulo4 = s
End Function

Private Function CodFerModulo5() As String
    Dim s As String
    s = s & "    l = (32 + 2 * e + 2 * i - h - k) Mod 7" & vbCrLf
    s = s & "    m = (a + 11 * h + 22 * l) \ 451" & vbCrLf
    s = s & "    mes = (h + l - 7 * m + 114) \ 31" & vbCrLf
    s = s & "    dia = ((h + l - 7 * m + 114) Mod 31) + 1" & vbCrLf
    s = s & "    Pascoa = DateSerial(ano, mes, dia)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' A regra QUANDO vale para a data d?" & vbCrLf
    s = s & "Private Function Bate(ByVal quando As String, ByVal d As Date) As Boolean" & vbCrLf
    s = s & "    Dim p As Date" & vbCrLf
    s = s & "    Dim n As Long" & vbCrLf
    s = s & "    If quando Like ""##/##/####"" Then" & vbCrLf
    s = s & "        Bate = (Format$(d, ""dd\/mm\/yyyy"") = quando)" & vbCrLf
    s = s & "    ElseIf quando Like ""##/##"" Then" & vbCrLf
    s = s & "        Bate = (Format$(d, ""dd\/mm"") = quando)" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        p = Pascoa(Year(d))" & vbCrLf
    s = s & "        Select Case quando" & vbCrLf
    s = s & "            Case ""SEXTA-FEIRA SANTA""" & vbCrLf
    s = s & "                Bate = (d = p - 2)" & vbCrLf
    s = s & "            Case ""CORPUS CHRISTI""" & vbCrLf
    s = s & "                Bate = (d = p + 60)" & vbCrLf
    s = s & "            Case ""CARNAVAL""" & vbCrLf
    s = s & "                Bate = (d = p - 48 Or d = p - 47)" & vbCrLf
    s = s & "            Case ""CARNAVAL TERCA""" & vbCrLf
    s = s & "                Bate = (d = p - 47)" & vbCrLf
    s = s & "            Case ""QUARTA-FEIRA DE CINZAS""" & vbCrLf
    s = s & "                Bate = (d = p - 46)" & vbCrLf
    s = s & "            Case Else" & vbCrLf
    s = s & "                If quando Like ""PASCOA[+-]*"" Then" & vbCrLf
    s = s & "                    If IsNumeric(Mid$(quando, 8)) Then" & vbCrLf
    s = s & "                        n = CLng(Mid$(quando, 8))" & vbCrLf
    s = s & "                        If Mid$(quando, 7, 1) = ""-"" Then n = -n" & vbCrLf
    s = s & "                        Bate = (d = p + n)" & vbCrLf
    s = s & "                    End If" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "        End Select" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Consulta" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Nacionais (bloqueiam): devolve o nome ("""" se n" & ChrW(227) & "o for)." & vbCrLf
    s = s & "Private Function Nacional(ByVal d As Date) As String" & vbCrLf
    CodFerModulo5 = s
End Function

Private Function CodFerModulo6() As String
    Dim s As String
    s = s & "    Select Case Format$(d, ""dd\/mm"")" & vbCrLf
    s = s & "        Case ""01/01""" & vbCrLf
    s = s & "            Nacional = ""Confraterniza" & ChrW(231) & "" & ChrW(227) & "o Universal""" & vbCrLf
    s = s & "        Case ""21/04""" & vbCrLf
    s = s & "            Nacional = ""Tiradentes""" & vbCrLf
    s = s & "        Case ""01/05""" & vbCrLf
    s = s & "            Nacional = ""Dia do Trabalho""" & vbCrLf
    s = s & "        Case ""07/09""" & vbCrLf
    s = s & "            Nacional = ""Independ" & ChrW(234) & "ncia do Brasil""" & vbCrLf
    s = s & "        Case ""12/10""" & vbCrLf
    s = s & "            Nacional = ""Nossa Senhora Aparecida""" & vbCrLf
    s = s & "        Case ""02/11""" & vbCrLf
    s = s & "            Nacional = ""Finados""" & vbCrLf
    s = s & "        Case ""15/11""" & vbCrLf
    s = s & "            Nacional = ""Proclama" & ChrW(231) & "" & ChrW(227) & "o da Rep" & ChrW(250) & "blica""" & vbCrLf
    s = s & "        Case ""20/11""" & vbCrLf
    s = s & "            Nacional = ""Dia da Consci" & ChrW(234) & "ncia Negra""" & vbCrLf
    s = s & "        Case ""25/12""" & vbCrLf
    s = s & "            Nacional = ""Natal""" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "    If Len(Nacional) = 0 Then" & vbCrLf
    s = s & "        If d = Pascoa(Year(d)) - 2 Then Nacional = ""Sexta-feira Santa""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FacultativoNacional(ByVal d As Date) As String" & vbCrLf
    s = s & "    Dim p As Date" & vbCrLf
    s = s & "    p = Pascoa(Year(d))" & vbCrLf
    s = s & "    If d = p - 48 Or d = p - 47 Then" & vbCrLf
    s = s & "        FacultativoNacional = ""Carnaval""" & vbCrLf
    s = s & "    ElseIf d = p - 46 Then" & vbCrLf
    s = s & "        FacultativoNacional = ""Quarta-feira de Cinzas""" & vbCrLf
    s = s & "    ElseIf d = p + 60 Then" & vbCrLf
    s = s & "        FacultativoNacional = ""Corpus Christi""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' O que " & ChrW(233) & " o dia d na cidade: DIA_UTIL, DIA_SABADO, DIA_DOMINGO, DIA_FERIADO" & vbCrLf
    s = s & "' ou DIA_FACULTATIVO. nome e abrangencia descrevem o feriado (quando houver)." & vbCrLf
    s = s & "Public Function InfoDia(ByVal d As Date, ByVal uf As String, ByVal cidade As String, _" & vbCrLf
    s = s & "                        ByRef nome As String, ByRef abrangencia As String) As Long" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim fac As String" & vbCrLf
    s = s & "    Dim facAbr As String" & vbCrLf
    s = s & "" & vbCrLf
    CodFerModulo6 = s
End Function

Private Function CodFerModulo7() As String
    Dim s As String
    s = s & "    nome = """"" & vbCrLf
    s = s & "    abrangencia = """"" & vbCrLf
    s = s & "    d = DateSerial(Year(d), Month(d), Day(d))" & vbCrLf
    s = s & "    If Not mCarregado Then Carregar" & vbCrLf
    s = s & "    uf = UCase$(Trim$(uf))" & vbCrLf
    s = s & "    cidade = NormCidade(cidade)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    nome = Nacional(d)" & vbCrLf
    s = s & "    If Len(nome) > 0 Then" & vbCrLf
    s = s & "        abrangencia = ""nacional""" & vbCrLf
    s = s & "        InfoDia = DIA_FERIADO" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 1 To mQtd" & vbCrLf
    s = s & "        If Len(mUF(i)) = 0 Or mUF(i) = uf Then" & vbCrLf
    s = s & "            If Len(mCid(i)) = 0 Or mCid(i) = cidade Then" & vbCrLf
    s = s & "                If Bate(mQuando(i), d) Then" & vbCrLf
    s = s & "                    If mAbr(i) = ""PONTO FACULTATIVO"" Then" & vbCrLf
    s = s & "                        If Len(fac) = 0 Then" & vbCrLf
    s = s & "                            fac = mNome(i)" & vbCrLf
    s = s & "                            facAbr = ""ponto facultativo""" & vbCrLf
    s = s & "                        End If" & vbCrLf
    s = s & "                    Else" & vbCrLf
    s = s & "                        nome = mNome(i)" & vbCrLf
    s = s & "                        If mAbr(i) = ""MUNICIPAL"" Then" & vbCrLf
    s = s & "                            abrangencia = ""municipal""" & vbCrLf
    s = s & "                        ElseIf mAbr(i) = ""ESTADUAL"" Then" & vbCrLf
    s = s & "                            abrangencia = ""estadual""" & vbCrLf
    s = s & "                        Else" & vbCrLf
    s = s & "                            abrangencia = LCase$(mAbr(i))" & vbCrLf
    s = s & "                        End If" & vbCrLf
    s = s & "                        InfoDia = DIA_FERIADO" & vbCrLf
    s = s & "                        Exit Function" & vbCrLf
    s = s & "                    End If" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Weekday(d, vbSunday) = vbSaturday Then" & vbCrLf
    s = s & "        InfoDia = DIA_SABADO" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Weekday(d, vbSunday) = vbSunday Then" & vbCrLf
    CodFerModulo7 = s
End Function

Private Function CodFerModulo8() As String
    Dim s As String
    s = s & "        InfoDia = DIA_DOMINGO" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Len(fac) = 0 Then" & vbCrLf
    s = s & "        fac = FacultativoNacional(d)" & vbCrLf
    s = s & "        facAbr = ""ponto facultativo""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Len(fac) > 0 Then" & vbCrLf
    s = s & "        nome = fac" & vbCrLf
    s = s & "        abrangencia = facAbr" & vbCrLf
    s = s & "        InfoDia = DIA_FACULTATIVO" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function Bloqueia(ByVal info As Long) As Boolean" & vbCrLf
    s = s & "    Bloqueia = (info = DIA_SABADO Or info = DIA_DOMINGO Or info = DIA_FERIADO)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function EhDiaUtil(ByVal d As Date, ByVal uf As String, ByVal cidade As String) As Boolean" & vbCrLf
    s = s & "    Dim n As String" & vbCrLf
    s = s & "    Dim a As String" & vbCrLf
    s = s & "    EhDiaUtil = Not Bloqueia(InfoDia(d, uf, cidade, n, a))" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function DiaUtilSeguinte(ByVal d As Date, ByVal uf As String, ByVal cidade As String) As Date" & vbCrLf
    s = s & "    Dim k As Long" & vbCrLf
    s = s & "    d = d + 1" & vbCrLf
    s = s & "    Do While Not EhDiaUtil(d, uf, cidade) And k < 30" & vbCrLf
    s = s & "        d = d + 1" & vbCrLf
    s = s & "        k = k + 1" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "    DiaUtilSeguinte = d" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function DiaUtilAnterior(ByVal d As Date, ByVal uf As String, ByVal cidade As String) As Date" & vbCrLf
    s = s & "    Dim k As Long" & vbCrLf
    s = s & "    d = d - 1" & vbCrLf
    s = s & "    Do While Not EhDiaUtil(d, uf, cidade) And k < 30" & vbCrLf
    s = s & "        d = d - 1" & vbCrLf
    s = s & "        k = k + 1" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "    DiaUtilAnterior = d" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    CodFerModulo8 = s
End Function

Private Function CodFerModulo9() As String
    Dim s As String
    s = s & "' Texto do motivo: ""s" & ChrW(225) & "bado"", ""domingo"", ""feriado nacional (Finados)"" ..." & vbCrLf
    s = s & "Public Function MotivoDia(ByVal info As Long, ByVal nome As String, ByVal abrangencia As String) As String" & vbCrLf
    s = s & "    Select Case info" & vbCrLf
    s = s & "        Case DIA_SABADO" & vbCrLf
    s = s & "            MotivoDia = ""s" & ChrW(225) & "bado""" & vbCrLf
    s = s & "        Case DIA_DOMINGO" & vbCrLf
    s = s & "            MotivoDia = ""domingo""" & vbCrLf
    s = s & "        Case DIA_FERIADO" & vbCrLf
    s = s & "            MotivoDia = ""feriado "" & abrangencia & "" ("" & nome & "")""" & vbCrLf
    s = s & "        Case DIA_FACULTATIVO" & vbCrLf
    s = s & "            MotivoDia = ""ponto facultativo ("" & nome & "")""" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function EtiquetaDia(ByVal info As Long) As String" & vbCrLf
    s = s & "    Select Case info" & vbCrLf
    s = s & "        Case DIA_SABADO" & vbCrLf
    s = s & "            EtiquetaDia = ""S" & ChrW(193) & "BADO""" & vbCrLf
    s = s & "        Case DIA_DOMINGO" & vbCrLf
    s = s & "            EtiquetaDia = ""DOMINGO""" & vbCrLf
    s = s & "        Case DIA_FERIADO" & vbCrLf
    s = s & "            EtiquetaDia = ""FERIADO""" & vbCrLf
    s = s & "        Case DIA_FACULTATIVO" & vbCrLf
    s = s & "            EtiquetaDia = ""P. FACULT.""" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function DiaSemanaCurto(ByVal d As Date) As String" & vbCrLf
    s = s & "    DiaSemanaCurto = Choose(Weekday(d, vbSunday), ""dom"", ""seg"", ""ter"", ""qua"", ""qui"", ""sex"", ""s" & ChrW(225) & "b"")" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function NomeMes(ByVal m As Long) As String" & vbCrLf
    s = s & "    NomeMes = Choose(m, ""JANEIRO"", ""FEVEREIRO"", ""MAR" & ChrW(199) & "O"", ""ABRIL"", ""MAIO"", ""JUNHO"", ""JULHO"", _" & vbCrLf
    s = s & "                     ""AGOSTO"", ""SETEMBRO"", ""OUTUBRO"", ""NOVEMBRO"", ""DEZEMBRO"")" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ""15/11/2026"", ""15112026"" ou data -> data (False se n" & ChrW(227) & "o for data a partir de 2000)." & vbCrLf
    s = s & "Public Function LerDataFer(ByVal texto As String, ByRef d As Date) As Boolean" & vbCrLf
    s = s & "    Dim s As String" & vbCrLf
    s = s & "    Dim dd As Long" & vbCrLf
    s = s & "    Dim mm As Long" & vbCrLf
    s = s & "    Dim aa As Long" & vbCrLf
    s = s & "    s = Replace(Replace(Replace(Trim$(texto), ""/"", """"), ""-"", """"), ""."", """")" & vbCrLf
    s = s & "    If Len(s) <> 8 Then Exit Function" & vbCrLf
    s = s & "    If Not s Like ""########"" Then Exit Function" & vbCrLf
    CodFerModulo9 = s
End Function

Private Function CodFerModulo10() As String
    Dim s As String
    s = s & "    dd = CLng(Left$(s, 2))" & vbCrLf
    s = s & "    mm = CLng(Mid$(s, 3, 2))" & vbCrLf
    s = s & "    aa = CLng(Right$(s, 4))" & vbCrLf
    s = s & "    If mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 Or aa < 2000 Or aa > 2100 Then Exit Function" & vbCrLf
    s = s & "    d = DateSerial(aa, mm, dd)" & vbCrLf
    s = s & "    If Day(d) <> dd Then Exit Function" & vbCrLf
    s = s & "    LerDataFer = True" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Fontes da identidade visual (Segoe UI de reserva)" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "Public Function FonteFeriados() As String" & vbCrLf
    s = s & "    If Len(mFonteTexto) = 0 Then DetectarFontes" & vbCrLf
    s = s & "    FonteFeriados = mFonteTexto" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function FonteTituloFeriados() As String" & vbCrLf
    s = s & "    If Len(mFonteTitulo) = 0 Then DetectarFontes" & vbCrLf
    s = s & "    FonteTituloFeriados = mFonteTitulo" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub DetectarFontes()" & vbCrLf
    s = s & "    Dim achado As String" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    achado = Dir$(Environ$(""WINDIR"") & ""\Fonts\Segoe UI*"")" & vbCrLf
    s = s & "    If Len(achado) = 0 Then achado = Dir$(Environ$(""LOCALAPPDATA"") & ""\Microsoft\Windows\Fonts\Segoe UI*"")" & vbCrLf
    s = s & "    If Len(achado) > 0 Then" & vbCrLf
    s = s & "        mFonteTexto = ""Segoe UI""" & vbCrLf
    s = s & "        mFonteTitulo = ""Segoe UI Black""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        mFonteTexto = ""Segoe UI""" & vbCrLf
    s = s & "        mFonteTitulo = ""Segoe UI Black""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodFerModulo10 = s
End Function

Private Function CodigoFerClique() As String
    Dim s As String
    s = s & CodFerClique1()
    CodigoFerClique = s
End Function

Private Function CodFerClique1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' cFerClique - repassa o clique de um r" & ChrW(243) & "tulo do calend" & ChrW(225) & "rio: dono.FerClique chave" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private WithEvents mRotulo As MSForms.Label" & vbCrLf
    s = s & "Private mDono As Object" & vbCrLf
    s = s & "Private mChave As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Conectar(ByVal rotulo As MSForms.Label, ByVal chave As String, ByVal dono As Object)" & vbCrLf
    s = s & "    Set mRotulo = rotulo" & vbCrLf
    s = s & "    mChave = chave" & vbCrLf
    s = s & "    Set mDono = dono" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Desconectar()" & vbCrLf
    s = s & "    Set mRotulo = Nothing" & vbCrLf
    s = s & "    Set mDono = Nothing" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub mRotulo_Click()" & vbCrLf
    s = s & "    If Not mDono Is Nothing Then mDono.FerClique mChave" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodFerClique1 = s
End Function

Private Function CodigoFerCampo() As String
    Dim s As String
    s = s & CodFerCampo1()
    CodigoFerCampo = s
End Function

Private Function CodFerCampo1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' cFerCampo - acompanha uma caixa de data do cronograma do painel:" & vbCrLf
    s = s & "' dono.FerCampoMudou indice (texto mudou) e dono.FerCampoFoco indice (clique)" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private WithEvents mCaixa As MSForms.TextBox" & vbCrLf
    s = s & "Private mDono As Object" & vbCrLf
    s = s & "Private mIndice As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Conectar(ByVal caixa As MSForms.TextBox, ByVal indice As Long, ByVal dono As Object)" & vbCrLf
    s = s & "    Set mCaixa = caixa" & vbCrLf
    s = s & "    mIndice = indice" & vbCrLf
    s = s & "    Set mDono = dono" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Desconectar()" & vbCrLf
    s = s & "    Set mCaixa = Nothing" & vbCrLf
    s = s & "    Set mDono = Nothing" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub mCaixa_Change()" & vbCrLf
    s = s & "    If Not mDono Is Nothing Then mDono.FerCampoMudou mIndice" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub mCaixa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)" & vbCrLf
    s = s & "    If Not mDono Is Nothing Then mDono.FerCampoFoco mIndice" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub mCaixa_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" & vbCrLf
    s = s & "    If Not mDono Is Nothing Then mDono.FerCampoFoco mIndice" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodFerCampo1 = s
End Function

Private Function CodigoFerCalendario() As String
    Dim s As String
    s = s & CodFerCalendario1()
    s = s & CodFerCalendario2()
    s = s & CodFerCalendario3()
    s = s & CodFerCalendario4()
    s = s & CodFerCalendario5()
    s = s & CodFerCalendario6()
    s = s & CodFerCalendario7()
    s = s & CodFerCalendario8()
    s = s & CodFerCalendario9()
    CodigoFerCalendario = s
End Function

Private Function CodFerCalendario1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' cCalendario - calend" & ChrW(225) & "rio do m" & ChrW(234) & "s e faixa de avisos do painel" & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "' Feriado = vermelho; ponto facultativo = amarelo claro; fim de semana =" & vbCrLf
    s = s & "' cinza; hoje = contorno escuro; data selecionada = contorno azul (vermelho" & vbCrLf
    s = s & "' se n" & ChrW(227) & "o for dia " & ChrW(250) & "til); datas da obra = pontinho embaixo do dia." & vbCrLf
    s = s & "' Cliques: dono.FerClique ""cal:prev"" | ""cal:next"" | ""cal:c:<n>"" | ""aviso:1|2""" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const CEL_W As Single = 26" & vbCrLf
    s = s & "Private Const CEL_H As Single = 21" & vbCrLf
    s = s & "Private Const CEL_G As Single = 2" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mDono As Object" & vbCrLf
    s = s & "Private mForm As Object" & vbCrLf
    s = s & "Private mCliques As Collection" & vbCrLf
    s = s & "Private mFonte As String" & vbCrLf
    s = s & "Private mFonteTit As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mL As Single" & vbCrLf
    s = s & "Private mT As Single" & vbCrLf
    s = s & "Private mW As Single" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mAno As Long" & vbCrLf
    s = s & "Private mMes As Long" & vbCrLf
    s = s & "Private mUF As String" & vbCrLf
    s = s & "Private mCidade As String" & vbCrLf
    s = s & "Private mSel As Date" & vbCrLf
    s = s & "Private mDatas As Collection" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private lblMes As MSForms.Label" & vbCrLf
    s = s & "Private lblLocal As MSForms.Label" & vbCrLf
    s = s & "Private lblLista As MSForms.Label" & vbCrLf
    s = s & "Private mCel(1 To 42) As MSForms.Label" & vbCrLf
    s = s & "Private mPonto(1 To 42) As MSForms.Label" & vbCrLf
    s = s & "Private mDataCel(1 To 42) As Date" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' faixa de avisos" & vbCrLf
    s = s & "Private lblFxFundo As MSForms.Label" & vbCrLf
    s = s & "Private lblFxBarra As MSForms.Label" & vbCrLf
    s = s & "Private lblFxTexto As MSForms.Label" & vbCrLf
    s = s & "Private lblFxB1 As MSForms.Label" & vbCrLf
    s = s & "Private lblFxB2 As MSForms.Label" & vbCrLf
    s = s & "" & vbCrLf
    CodFerCalendario1 = s
End Function

Private Function CodFerCalendario2() As String
    Dim s As String
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Montagem" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "Private Function Rot(ByVal nome As String, ByVal texto As String, _" & vbCrLf
    s = s & "                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _" & vbCrLf
    s = s & "                     ByVal cor As Long, ByVal tam As Single, ByVal negrito As Boolean, _" & vbCrLf
    s = s & "                     Optional ByVal fundo As Long = -1, Optional ByVal alinh As Long = 1) As MSForms.Label" & vbCrLf
    s = s & "    Dim c As MSForms.Label" & vbCrLf
    s = s & "    Set c = mForm.Controls.Add(""Forms.Label.1"", nome, True)" & vbCrLf
    s = s & "    c.Left = l" & vbCrLf
    s = s & "    c.Top = t" & vbCrLf
    s = s & "    c.Width = w" & vbCrLf
    s = s & "    c.Height = h" & vbCrLf
    s = s & "    c.Font.Name = mFonte" & vbCrLf
    s = s & "    c.Font.Size = tam" & vbCrLf
    s = s & "    c.Font.Bold = negrito" & vbCrLf
    s = s & "    c.ForeColor = cor" & vbCrLf
    s = s & "    If fundo >= 0 Then" & vbCrLf
    s = s & "        c.BackStyle = 1" & vbCrLf
    s = s & "        c.BackColor = fundo" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        c.BackStyle = 0" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    c.TextAlign = alinh" & vbCrLf
    s = s & "    c.WordWrap = False" & vbCrLf
    s = s & "    c.Caption = texto" & vbCrLf
    s = s & "    Set Rot = c" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Clique(ByVal rotulo As MSForms.Label, ByVal chave As String)" & vbCrLf
    s = s & "    Dim ev As cFerClique" & vbCrLf
    s = s & "    Set ev = New cFerClique" & vbCrLf
    s = s & "    ev.Conectar rotulo, chave, mDono" & vbCrLf
    s = s & "    mCliques.Add ev" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' l, t, w: cart" & ChrW(227) & "o do calend" & ChrW(225) & "rio (altura 324); fl, ft, fw: faixa de avisos (altura 24)" & vbCrLf
    s = s & "Public Sub Criar(ByVal dono As Object, ByVal l As Single, ByVal t As Single, ByVal w As Single, _" & vbCrLf
    s = s & "                 ByVal fl As Single, ByVal ft As Single, ByVal fw As Single)" & vbCrLf
    s = s & "    Dim c As MSForms.Label" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim x0 As Single" & vbCrLf
    s = s & "    Dim x As Single" & vbCrLf
    s = s & "    Dim y As Single" & vbCrLf
    s = s & "    Dim dias As Variant" & vbCrLf
    CodFerCalendario2 = s
End Function

Private Function CodFerCalendario3() As String
    Dim s As String
    s = s & "    Dim leg As Variant" & vbCrLf
    s = s & "    Dim corL As Variant" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set mDono = dono" & vbCrLf
    s = s & "    Set mForm = dono" & vbCrLf
    s = s & "    Set mCliques = New Collection" & vbCrLf
    s = s & "    Set mDatas = New Collection" & vbCrLf
    s = s & "    mFonte = FonteFeriados()" & vbCrLf
    s = s & "    mFonteTit = FonteTituloFeriados()" & vbCrLf
    s = s & "    mL = l" & vbCrLf
    s = s & "    mT = t" & vbCrLf
    s = s & "    mW = w" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set c = Rot(""calFundo"", """", l, t, w, 324, FR_NAVY, 8, False, FR_BRANCO)" & vbCrLf
    s = s & "    c.BorderStyle = 1" & vbCrLf
    s = s & "    c.BorderColor = FR_BORDA" & vbCrLf
    s = s & "    Rot ""calTopo"", """", l, t, w, 22, FR_BRANCO, 8, False, FR_NAVY" & vbCrLf
    s = s & "    Rot ""calTopoAm"", """", l, t + 22, w, 2, FR_BRANCO, 8, False, FR_AMARELO" & vbCrLf
    s = s & "    Set c = Rot(""calAnt"", ""<"", l + 2, t + 3, 22, 16, FR_BRANCO, 10, True, -1, 2)" & vbCrLf
    s = s & "    c.ControlTipText = ""M" & ChrW(234) & "s anterior""" & vbCrLf
    s = s & "    Clique c, ""cal:prev""" & vbCrLf
    s = s & "    Set c = Rot(""calProx"", "">"", l + w - 24, t + 3, 22, 16, FR_BRANCO, 10, True, -1, 2)" & vbCrLf
    s = s & "    c.ControlTipText = ""Pr" & ChrW(243) & "ximo m" & ChrW(234) & "s""" & vbCrLf
    s = s & "    Clique c, ""cal:next""" & vbCrLf
    s = s & "    Set lblMes = Rot(""calMes"", """", l + 26, t + 4, w - 52, 15, FR_BRANCO, 9.5, False, -1, 2)" & vbCrLf
    s = s & "    lblMes.Font.Name = mFonteTit" & vbCrLf
    s = s & "    Set lblLocal = Rot(""calLocal"", """", l + 4, t + 27, w - 8, 11, FR_CINZA, 7, False, -1, 2)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    x0 = l + (w - (7 * CEL_W + 6 * CEL_G)) / 2" & vbCrLf
    s = s & "    dias = Array(""D"", ""S"", ""T"", ""Q"", ""Q"", ""S"", ""S"")" & vbCrLf
    s = s & "    For i = 0 To 6" & vbCrLf
    s = s & "        Rot ""calSem"" & i, CStr(dias(i)), x0 + i * (CEL_W + CEL_G), t + 40, CEL_W, 11, FR_CINZA, 7, True, -1, 2" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    For i = 1 To 42" & vbCrLf
    s = s & "        x = x0 + ((i - 1) Mod 7) * (CEL_W + CEL_G)" & vbCrLf
    s = s & "        y = t + 53 + ((i - 1) \ 7) * (CEL_H + CEL_G)" & vbCrLf
    s = s & "        Set mCel(i) = Rot(""calCel"" & i, """", x, y, CEL_W, CEL_H, FR_NAVY, 8.5, False, FR_BRANCO, 2)" & vbCrLf
    s = s & "        mCel(i).BorderStyle = 0" & vbCrLf
    s = s & "        Clique mCel(i), ""cal:c:"" & i" & vbCrLf
    s = s & "        Set mPonto(i) = Rot(""calPto"" & i, """", x + CEL_W / 2 - 2, y + CEL_H - 5, 4, 3, FR_AZUL, 6, False, FR_AZUL)" & vbCrLf
    s = s & "        mPonto(i).Visible = False" & vbCrLf
    s = s & "        Clique mPonto(i), ""cal:c:"" & i" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' legenda" & vbCrLf
    CodFerCalendario3 = s
End Function

Private Function CodFerCalendario4() As String
    Dim s As String
    s = s & "    y = t + 53 + 6 * (CEL_H + CEL_G) + 2" & vbCrLf
    s = s & "    leg = Array(""feriado"", ""ponto facult."", ""fim de semana"", ""data da obra"")" & vbCrLf
    s = s & "    corL = Array(FR_CORAL, FR_DESTAQUE, FR_FDS, FR_AZUL)" & vbCrLf
    s = s & "    For i = 0 To 3" & vbCrLf
    s = s & "        x = l + 8 + (i Mod 2) * 96" & vbCrLf
    s = s & "        Set c = Rot(""calLegC"" & i, """", x, y + 3 + (i \ 2) * 12, 7, 7, FR_NAVY, 6, False, CLng(corL(i)))" & vbCrLf
    s = s & "        If i = 1 Then" & vbCrLf
    s = s & "            c.BorderStyle = 1" & vbCrLf
    s = s & "            c.BorderColor = FR_AMARELO" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        Rot ""calLegT"" & i, CStr(leg(i)), x + 10, y + (i \ 2) * 12, 84, 11, FR_CINZA, 7, False" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set c = Rot(""calSep"", """", l + 6, y + 27, w - 12, 1, FR_NAVY, 6, False, FR_BORDA)" & vbCrLf
    s = s & "    Set lblLista = Rot(""calLista"", """", l + 8, y + 31, w - 16, t + 320 - (y + 31), FR_NAVY, 7.5, False)" & vbCrLf
    s = s & "    lblLista.WordWrap = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' faixa de avisos" & vbCrLf
    s = s & "    Set lblFxFundo = Rot(""fxFundo"", """", fl, ft, fw, 24, FR_NAVY, 8, False, FR_FUNDO)" & vbCrLf
    s = s & "    lblFxFundo.BorderStyle = 1" & vbCrLf
    s = s & "    lblFxFundo.BorderColor = FR_BORDA" & vbCrLf
    s = s & "    Set lblFxBarra = Rot(""fxBarra"", """", fl, ft, 4, 24, FR_NAVY, 8, False, FR_BORDA)" & vbCrLf
    s = s & "    Set lblFxTexto = Rot(""fxTexto"", """", fl + 12, ft + 5, fw - 300, 14, FR_CINZA, 8.5, False)" & vbCrLf
    s = s & "    Set lblFxB1 = Rot(""fxB1"", """", fl + fw - 276, ft + 3, 132, 18, FR_NAVY, 8.5, True, FR_BRANCO, 2)" & vbCrLf
    s = s & "    lblFxB1.BorderStyle = 1" & vbCrLf
    s = s & "    lblFxB1.BorderColor = FR_BORDA" & vbCrLf
    s = s & "    Clique lblFxB1, ""aviso:1""" & vbCrLf
    s = s & "    Set lblFxB2 = Rot(""fxB2"", """", fl + fw - 138, ft + 3, 132, 18, FR_BRANCO, 8.5, True, FR_NAVY, 2)" & vbCrLf
    s = s & "    Clique lblFxB2, ""aviso:2""" & vbCrLf
    s = s & "    Dica """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Desligar()" & vbCrLf
    s = s & "    Dim ev As cFerClique" & vbCrLf
    s = s & "    If mCliques Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    For Each ev In mCliques" & vbCrLf
    s = s & "        ev.Desconectar" & vbCrLf
    s = s & "    Next ev" & vbCrLf
    s = s & "    Set mCliques = Nothing" & vbCrLf
    s = s & "    Set mDono = Nothing" & vbCrLf
    s = s & "    Set mForm = Nothing" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Dados" & vbCrLf
    CodFerCalendario4 = s
End Function

Private Function CodFerCalendario5() As String
    Dim s As String
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "Public Sub Obra(ByVal uf As String, ByVal cidade As String)" & vbCrLf
    s = s & "    mUF = uf" & vbCrLf
    s = s & "    mCidade = cidade" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' datas da obra: cada item = Array(data, r" & ChrW(243) & "tulo)" & vbCrLf
    s = s & "Public Sub Datas(ByVal itens As Collection)" & vbCrLf
    s = s & "    If itens Is Nothing Then" & vbCrLf
    s = s & "        Set mDatas = New Collection" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Set mDatas = itens" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get DataDaCelula(ByVal n As Long) As Date" & vbCrLf
    s = s & "    If n >= 1 And n <= 42 Then DataDaCelula = mDataCel(n)" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Mover(ByVal passo As Long)" & vbCrLf
    s = s & "    Dim d As Date" & vbCrLf
    s = s & "    d = DateSerial(mAno, mMes + passo, 1)" & vbCrLf
    s = s & "    Mostrar Year(d), Month(d), mSel" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Desenho" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "Public Sub Mostrar(ByVal ano As Long, ByVal mes As Long, Optional ByVal selecionada As Date = 0)" & vbCrLf
    s = s & "    Dim primeiro As Date" & vbCrLf
    s = s & "    Dim d As Date" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim info As Long" & vbCrLf
    s = s & "    Dim nome As String" & vbCrLf
    s = s & "    Dim abr As String" & vbCrLf
    s = s & "    Dim lista As String" & vbCrLf
    s = s & "    Dim prox As String" & vbCrLf
    s = s & "    Dim k As Long" & vbCrLf
    s = s & "    Dim ultimo As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If lblMes Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    mAno = ano" & vbCrLf
    s = s & "    mMes = mes" & vbCrLf
    s = s & "    mSel = selecionada" & vbCrLf
    s = s & "    lblMes.Caption = NomeMes(mes) & "" "" & ano" & vbCrLf
    CodFerCalendario5 = s
End Function

Private Function CodFerCalendario6() As String
    Dim s As String
    s = s & "    If Len(mCidade) > 0 Then" & vbCrLf
    s = s & "        lblLocal.Caption = ""feriados de "" & UCase$(mCidade) & IIf(Len(mUF) > 0, ""-"" & UCase$(mUF), """")" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblLocal.Caption = ""feriados nacionais"" & IIf(Len(mUF) > 0, "" e de "" & UCase$(mUF), """")" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    primeiro = DateSerial(ano, mes, 1)" & vbCrLf
    s = s & "    ultimo = Day(DateSerial(ano, mes + 1, 0))" & vbCrLf
    s = s & "    For i = 1 To 42" & vbCrLf
    s = s & "        d = primeiro + (i - Weekday(primeiro, vbSunday))" & vbCrLf
    s = s & "        mPonto(i).Visible = False" & vbCrLf
    s = s & "        mCel(i).BorderStyle = 0" & vbCrLf
    s = s & "        If Month(d) <> mes Then" & vbCrLf
    s = s & "            mDataCel(i) = 0" & vbCrLf
    s = s & "            mCel(i).Caption = """"" & vbCrLf
    s = s & "            mCel(i).BackColor = FR_BRANCO" & vbCrLf
    s = s & "            mCel(i).ControlTipText = """"" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            mDataCel(i) = d" & vbCrLf
    s = s & "            mCel(i).Caption = CStr(Day(d))" & vbCrLf
    s = s & "            info = InfoDia(d, mUF, mCidade, nome, abr)" & vbCrLf
    s = s & "            mCel(i).Font.Bold = False" & vbCrLf
    s = s & "            Select Case info" & vbCrLf
    s = s & "                Case DIA_FERIADO" & vbCrLf
    s = s & "                    mCel(i).BackColor = FR_CORAL" & vbCrLf
    s = s & "                    mCel(i).ForeColor = FR_BRANCO" & vbCrLf
    s = s & "                    mCel(i).Font.Bold = True" & vbCrLf
    s = s & "                    mCel(i).ControlTipText = Format$(d, ""dd\/mm"") & "" - "" & nome & "" ("" & abr & "")""" & vbCrLf
    s = s & "                    lista = lista & Format$(d, ""dd\/mm"") & "" "" & DiaSemanaCurto(d) & "" "" & ChrW(183) & "" "" & nome & "" ("" & abr & "")"" & vbLf" & vbCrLf
    s = s & "                Case DIA_FACULTATIVO" & vbCrLf
    s = s & "                    mCel(i).BackColor = FR_DESTAQUE" & vbCrLf
    s = s & "                    mCel(i).ForeColor = FR_NAVY" & vbCrLf
    s = s & "                    mCel(i).BorderStyle = 1" & vbCrLf
    s = s & "                    mCel(i).BorderColor = FR_AMARELO" & vbCrLf
    s = s & "                    mCel(i).ControlTipText = Format$(d, ""dd\/mm"") & "" - "" & nome & "" (ponto facultativo)""" & vbCrLf
    s = s & "                    lista = lista & Format$(d, ""dd\/mm"") & "" "" & DiaSemanaCurto(d) & "" "" & ChrW(183) & "" "" & nome & "" (ponto facultativo)"" & vbLf" & vbCrLf
    s = s & "                Case DIA_SABADO, DIA_DOMINGO" & vbCrLf
    s = s & "                    mCel(i).BackColor = FR_FDS" & vbCrLf
    s = s & "                    mCel(i).ForeColor = FR_FDS_TEXTO" & vbCrLf
    s = s & "                    mCel(i).ControlTipText = Format$(d, ""dd\/mm"") & "" - fim de semana""" & vbCrLf
    s = s & "                Case Else" & vbCrLf
    s = s & "                    mCel(i).BackColor = FR_BRANCO" & vbCrLf
    s = s & "                    mCel(i).ForeColor = FR_NAVY" & vbCrLf
    s = s & "                    mCel(i).ControlTipText = Format$(d, ""dd\/mm"") & "" - dia " & ChrW(250) & "til""" & vbCrLf
    s = s & "            End Select" & vbCrLf
    CodFerCalendario6 = s
End Function

Private Function CodFerCalendario7() As String
    Dim s As String
    s = s & "            If d = Date Then" & vbCrLf
    s = s & "                mCel(i).BorderStyle = 1" & vbCrLf
    s = s & "                mCel(i).BorderColor = FR_NAVY" & vbCrLf
    s = s & "                mCel(i).Font.Bold = True" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "            If d = selecionada Then" & vbCrLf
    s = s & "                mCel(i).BorderStyle = 1" & vbCrLf
    s = s & "                mCel(i).Font.Bold = True" & vbCrLf
    s = s & "                If Bloqueia(info) Then" & vbCrLf
    s = s & "                    mCel(i).BorderColor = FR_CORAL" & vbCrLf
    s = s & "                    If info <> DIA_FERIADO Then" & vbCrLf
    s = s & "                        mCel(i).BackColor = FR_CORAL_FUNDO" & vbCrLf
    s = s & "                        mCel(i).ForeColor = FR_CORAL" & vbCrLf
    s = s & "                    End If" & vbCrLf
    s = s & "                Else" & vbCrLf
    s = s & "                    mCel(i).BorderColor = FR_AZUL" & vbCrLf
    s = s & "                    mCel(i).BackColor = FR_AZUL_FUNDO" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "            MarcarPonto i, d, info" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Len(lista) = 0 Then lista = ""Nenhum feriado neste m" & ChrW(234) & "s."" & vbLf" & vbCrLf
    s = s & "    prox = ProximoFeriado(DateSerial(ano, mes, ultimo))" & vbCrLf
    s = s & "    If Len(prox) > 0 Then lista = lista & ""pr" & ChrW(243) & "ximo: "" & prox" & vbCrLf
    s = s & "    lblLista.Caption = lista" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub MarcarPonto(ByVal i As Long, ByVal d As Date, ByVal info As Long)" & vbCrLf
    s = s & "    Dim it As Variant" & vbCrLf
    s = s & "    Dim rotulos As String" & vbCrLf
    s = s & "    If mDatas Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    For Each it In mDatas" & vbCrLf
    s = s & "        If CDate(it(0)) = d Then" & vbCrLf
    s = s & "            If Len(rotulos) > 0 Then rotulos = rotulos & "", """ & vbCrLf
    s = s & "            rotulos = rotulos & it(1)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next it" & vbCrLf
    s = s & "    If Len(rotulos) = 0 Then Exit Sub" & vbCrLf
    s = s & "    mPonto(i).Visible = True" & vbCrLf
    s = s & "    If info = DIA_FERIADO Then" & vbCrLf
    s = s & "        mPonto(i).BackColor = FR_BRANCO" & vbCrLf
    s = s & "    ElseIf Bloqueia(info) Then" & vbCrLf
    s = s & "        mPonto(i).BackColor = FR_OURO" & vbCrLf
    CodFerCalendario7 = s
End Function

Private Function CodFerCalendario8() As String
    Dim s As String
    s = s & "    Else" & vbCrLf
    s = s & "        mPonto(i).BackColor = FR_AZUL" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    mCel(i).ControlTipText = mCel(i).ControlTipText & ""  |  "" & rotulos" & vbCrLf
    s = s & "    mPonto(i).ControlTipText = mCel(i).ControlTipText" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Pr" & ChrW(243) & "ximo feriado (bloqueante) depois da data, em at" & ChrW(233) & " 120 dias." & vbCrLf
    s = s & "Private Function ProximoFeriado(ByVal depoisDe As Date) As String" & vbCrLf
    s = s & "    Dim d As Date" & vbCrLf
    s = s & "    Dim k As Long" & vbCrLf
    s = s & "    Dim nome As String" & vbCrLf
    s = s & "    Dim abr As String" & vbCrLf
    s = s & "    d = depoisDe" & vbCrLf
    s = s & "    For k = 1 To 120" & vbCrLf
    s = s & "        d = d + 1" & vbCrLf
    s = s & "        If InfoDia(d, mUF, mCidade, nome, abr) = DIA_FERIADO Then" & vbCrLf
    s = s & "            ProximoFeriado = Format$(d, ""dd\/mm"") & "" "" & DiaSemanaCurto(d) & "" "" & ChrW(183) & "" "" & nome" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next k" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Faixa de avisos" & vbCrLf
    s = s & "' tipo: 0 dica (cinza), 1 erro (vermelho), 2 aten" & ChrW(231) & "" & ChrW(227) & "o (amarelo), 3 ok (azul)" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "Public Sub Aviso(ByVal texto As String, ByVal tipo As Long, Optional ByVal botao1 As String = """", Optional ByVal botao2 As String = """")" & vbCrLf
    s = s & "    If lblFxFundo Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    Select Case tipo" & vbCrLf
    s = s & "        Case 1" & vbCrLf
    s = s & "            lblFxFundo.BackColor = FR_CORAL_FUNDO" & vbCrLf
    s = s & "            lblFxBarra.BackColor = FR_CORAL" & vbCrLf
    s = s & "            lblFxTexto.ForeColor = FR_NAVY" & vbCrLf
    s = s & "        Case 2" & vbCrLf
    s = s & "            lblFxFundo.BackColor = FR_DESTAQUE" & vbCrLf
    s = s & "            lblFxBarra.BackColor = FR_AMARELO" & vbCrLf
    s = s & "            lblFxTexto.ForeColor = FR_NAVY" & vbCrLf
    s = s & "        Case 3" & vbCrLf
    s = s & "            lblFxFundo.BackColor = FR_AZUL_FUNDO" & vbCrLf
    s = s & "            lblFxBarra.BackColor = FR_AZUL" & vbCrLf
    s = s & "            lblFxTexto.ForeColor = FR_NAVY" & vbCrLf
    s = s & "        Case Else" & vbCrLf
    s = s & "            lblFxFundo.BackColor = FR_FUNDO" & vbCrLf
    s = s & "            lblFxBarra.BackColor = FR_BORDA" & vbCrLf
    CodFerCalendario8 = s
End Function

Private Function CodFerCalendario9() As String
    Dim s As String
    s = s & "            lblFxTexto.ForeColor = FR_CINZA" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "    lblFxTexto.Caption = texto" & vbCrLf
    s = s & "    lblFxTexto.ControlTipText = texto" & vbCrLf
    s = s & "    lblFxB1.Caption = botao1" & vbCrLf
    s = s & "    lblFxB2.Caption = botao2" & vbCrLf
    s = s & "    lblFxB1.Visible = (Len(botao1) > 0)" & vbCrLf
    s = s & "    lblFxB2.Visible = (Len(botao2) > 0)" & vbCrLf
    s = s & "    If Len(botao1) > 0 Or Len(botao2) > 0 Then" & vbCrLf
    s = s & "        lblFxTexto.Width = lblFxFundo.Width - 300" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblFxTexto.Width = lblFxFundo.Width - 20" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Dica(ByVal texto As String)" & vbCrLf
    s = s & "    If Len(texto) = 0 Then texto = ""Selecione um campo do cronograma e digite a data, ou clique num dia " & ChrW(250) & "til do calend" & ChrW(225) & "rio.""" & vbCrLf
    s = s & "    Aviso texto, 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodFerCalendario9 = s
End Function

Private Function CodigoMarcModulo() As String
    Dim s As String
    s = s & CodMarcModulo1()
    s = s & CodMarcModulo2()
    s = s & CodMarcModulo3()
    s = s & CodMarcModulo4()
    s = s & CodMarcModulo5()
    CodigoMarcModulo = s
End Function

Private Function CodMarcModulo1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' mMarcador - marcador ""para analisar"" das obras (marca-p" & ChrW(225) & "gina)" & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "' Entradas:" & vbCrLf
    s = s & "'   AbrirMarcadas          - lista das obras marcadas (abre a obra escolhida" & vbCrLf
    s = s & "'                            no Painel de Obras)" & vbCrLf
    s = s & "'   AbrirMarcador cod, obra - janela para marcar / editar / concluir" & vbCrLf
    s = s & "'   AbrirListaMarcadas     - lista; devolve o C" & ChrW(211) & "DIGO escolhido ("""" = nenhum)" & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "' Os marcadores ficam na aba oculta MARCADORES (tabela tb_Marcadores)." & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Const MC_NAVY As Long = 3877150" & vbCrLf
    s = s & "Public Const MC_NAVY2 As Long = 5587251" & vbCrLf
    s = s & "Public Const MC_AZUL As Long = 8950797" & vbCrLf
    s = s & "Public Const MC_AZULCLARO As Long = 15858636" & vbCrLf
    s = s & "Public Const MC_CORAL As Long = 2500316" & vbCrLf
    s = s & "Public Const MC_AMARELO As Long = 2408443" & vbCrLf
    s = s & "Public Const MC_CINZA As Long = 9139300" & vbCrLf
    s = s & "Public Const MC_CINZA_CLARO As Long = 12100500" & vbCrLf
    s = s & "Public Const MC_TEXTO_CLARO As Long = 14800331" & vbCrLf
    s = s & "Public Const MC_FUNDO As Long = 16579320" & vbCrLf
    s = s & "Public Const MC_BORDA As Long = 15788258" & vbCrLf
    s = s & "Public Const MC_BRANCO As Long = 16777215" & vbCrLf
    s = s & "Public Const MC_DESTAQUE As Long = 15465471" & vbCrLf
    s = s & "Public Const MC_CORAL_FUNDO As Long = 14869246" & vbCrLf
    s = s & "Public Const MC_EDICAO As Long = 15465471" & vbCrLf
    s = s & "Public Const MC_SELECAO As Long = 15858636" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' situa" & ChrW(231) & "" & ChrW(227) & "o do prazo" & vbCrLf
    s = s & "Public Const PZ_SEM_DATA As Long = 0" & vbCrLf
    s = s & "Public Const PZ_FUTURA As Long = 1" & vbCrLf
    s = s & "Public Const PZ_HOJE As Long = 2" & vbCrLf
    s = s & "Public Const PZ_ATRASADA As Long = 3" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mFonteTexto As String" & vbCrLf
    s = s & "Private mFonteTitulo As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Entradas" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Devolve 0 (nada mudou), 1 (marcador salvo) ou 2 (an" & ChrW(225) & "lise conclu" & ChrW(237) & "da)." & vbCrLf
    s = s & "Public Function AbrirMarcador(ByVal codObra As String, ByVal obra As String) As Long" & vbCrLf
    s = s & "    frmMarcador.Preparar codObra, obra" & vbCrLf
    CodMarcModulo1 = s
End Function

Private Function CodMarcModulo2() As String
    Dim s As String
    s = s & "    frmMarcador.Show vbModal" & vbCrLf
    s = s & "    AbrirMarcador = frmMarcador.Resultado" & vbCrLf
    s = s & "    Unload frmMarcador" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Devolve o C" & ChrW(211) & "DIGO da obra escolhida para abrir ("""" = nenhuma)." & vbCrLf
    s = s & "Public Function AbrirListaMarcadas() As String" & vbCrLf
    s = s & "    frmMarcadas.Preparar" & vbCrLf
    s = s & "    frmMarcadas.Show vbModal" & vbCrLf
    s = s & "    AbrirListaMarcadas = frmMarcadas.Escolhida" & vbCrLf
    s = s & "    Unload frmMarcadas" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Macro: lista das obras marcadas; a escolhida abre no painel dela." & vbCrLf
    s = s & "Public Sub AbrirMarcadas()" & vbCrLf
    s = s & "    Dim cod As String" & vbCrLf
    s = s & "    cod = AbrirListaMarcadas()" & vbCrLf
    s = s & "    If Len(cod) > 0 Then AbrirObraMarcada cod" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Abre a obra no Painel de Obras." & vbCrLf
    s = s & "Public Sub AbrirObraMarcada(ByVal codObra As String)" & vbCrLf
    s = s & "    Dim marcas As cMarcadores" & vbCrLf
    s = s & "    Dim idx As Long" & vbCrLf
    s = s & "    Dim f As Object" & vbCrLf
    s = s & "    Dim painel As Object" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set marcas = New cMarcadores" & vbCrLf
    s = s & "    idx = marcas.LinhaNoFollowUp(codObra)" & vbCrLf
    s = s & "    If idx = 0 Then Exit Sub" & vbCrLf
    s = s & "    For Each f In VBA.UserForms" & vbCrLf
    s = s & "        If f.Name = ""frmPainelObras"" Then Set painel = f" & vbCrLf
    s = s & "    Next f" & vbCrLf
    s = s & "    If painel Is Nothing Then" & vbCrLf
    s = s & "        On Error Resume Next" & vbCrLf
    s = s & "        Set painel = VBA.UserForms.Add(""frmPainelObras"")" & vbCrLf
    s = s & "        On Error GoTo 0" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If painel Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    painel.IrParaLinha idx" & vbCrLf
    s = s & "    If Not painel.Visible Then painel.Show vbModeless" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Usada pelo instalador: cria a aba oculta MARCADORES (se ainda n" & ChrW(227) & "o existir)." & vbCrLf
    s = s & "Public Function PrepararMarcadores() As Boolean" & vbCrLf
    CodMarcModulo2 = s
End Function

Private Function CodMarcModulo3() As String
    Dim s As String
    s = s & "    Dim marcas As cMarcadores" & vbCrLf
    s = s & "    Set marcas = New cMarcadores" & vbCrLf
    s = s & "    PrepararMarcadores = Not (marcas.Tabela(True) Is Nothing)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Prazos" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "Public Function TipoPrazo(ByVal revisar As Variant) As Long" & vbCrLf
    s = s & "    If Not IsDate(revisar) Then" & vbCrLf
    s = s & "        TipoPrazo = PZ_SEM_DATA" & vbCrLf
    s = s & "    ElseIf Int(CDbl(CDate(revisar))) < Int(CDbl(Date)) Then" & vbCrLf
    s = s & "        TipoPrazo = PZ_ATRASADA" & vbCrLf
    s = s & "    ElseIf Int(CDbl(CDate(revisar))) = Int(CDbl(Date)) Then" & vbCrLf
    s = s & "        TipoPrazo = PZ_HOJE" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        TipoPrazo = PZ_FUTURA" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Texto curto do prazo: ""revisar hoje " & ChrW(183) & " 08/10"", ""atrasada " & ChrW(183) & " 06/10"" ..." & vbCrLf
    s = s & "Public Function TextoPrazo(ByVal revisar As Variant) As String" & vbCrLf
    s = s & "    Select Case TipoPrazo(revisar)" & vbCrLf
    s = s & "        Case PZ_ATRASADA" & vbCrLf
    s = s & "            TextoPrazo = ""atrasada "" & ChrW(183) & "" "" & Format$(CDate(revisar), ""dd/mm"")" & vbCrLf
    s = s & "        Case PZ_HOJE" & vbCrLf
    s = s & "            TextoPrazo = ""revisar hoje "" & ChrW(183) & "" "" & Format$(CDate(revisar), ""dd/mm"")" & vbCrLf
    s = s & "        Case PZ_FUTURA" & vbCrLf
    s = s & "            TextoPrazo = ""revisar em "" & Format$(CDate(revisar), ""dd/mm"")" & vbCrLf
    s = s & "        Case Else" & vbCrLf
    s = s & "            TextoPrazo = ""sem data""" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function CorPrazo(ByVal revisar As Variant) As Long" & vbCrLf
    s = s & "    Select Case TipoPrazo(revisar)" & vbCrLf
    s = s & "        Case PZ_ATRASADA" & vbCrLf
    s = s & "            CorPrazo = MC_CORAL" & vbCrLf
    s = s & "        Case PZ_HOJE" & vbCrLf
    s = s & "            CorPrazo = MC_AMARELO" & vbCrLf
    s = s & "        Case Else" & vbCrLf
    s = s & "            CorPrazo = MC_CINZA_CLARO" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    CodMarcModulo3 = s
End Function

Private Function CodMarcModulo4() As String
    Dim s As String
    s = s & "Public Function DiaDaSemana(ByVal d As Date) As String" & vbCrLf
    s = s & "    DiaDaSemana = Choose(Weekday(d, vbSunday), ""domingo"", ""segunda"", ""ter" & ChrW(231) & "a"", ""quarta"", ""quinta"", ""sexta"", ""s" & ChrW(225) & "bado"")" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ""15/10/2026"" ou ""15102026"" -> data; devolve False se n" & ChrW(227) & "o for data v" & ChrW(225) & "lida." & vbCrLf
    s = s & "Public Function LerData(ByVal texto As String, ByRef d As Date) As Boolean" & vbCrLf
    s = s & "    Dim s As String" & vbCrLf
    s = s & "    Dim dd As Long" & vbCrLf
    s = s & "    Dim mm As Long" & vbCrLf
    s = s & "    Dim aa As Long" & vbCrLf
    s = s & "    s = Replace(Replace(Replace(Trim$(texto), ""/"", """"), ""-"", """"), ""."", """")" & vbCrLf
    s = s & "    If Len(s) = 6 Then s = Left$(s, 4) & ""20"" & Right$(s, 2)" & vbCrLf
    s = s & "    If Len(s) <> 8 Then Exit Function" & vbCrLf
    s = s & "    If Not s Like ""########"" Then Exit Function" & vbCrLf
    s = s & "    dd = CLng(Left$(s, 2))" & vbCrLf
    s = s & "    mm = CLng(Mid$(s, 3, 2))" & vbCrLf
    s = s & "    aa = CLng(Right$(s, 4))" & vbCrLf
    s = s & "    If mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 Or aa < 2000 Or aa > 2100 Then Exit Function" & vbCrLf
    s = s & "    d = DateSerial(aa, mm, dd)" & vbCrLf
    s = s & "    If Day(d) <> dd Then Exit Function" & vbCrLf
    s = s & "    LerData = True" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Fontes da identidade visual (Segoe UI de reserva)" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "Public Function FonteMarcador() As String" & vbCrLf
    s = s & "    If Len(mFonteTexto) = 0 Then DetectarFontes" & vbCrLf
    s = s & "    FonteMarcador = mFonteTexto" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function FonteTituloMarcador() As String" & vbCrLf
    s = s & "    If Len(mFonteTitulo) = 0 Then DetectarFontes" & vbCrLf
    s = s & "    FonteTituloMarcador = mFonteTitulo" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub DetectarFontes()" & vbCrLf
    s = s & "    Dim achado As String" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    achado = Dir$(Environ$(""WINDIR"") & ""\Fonts\Segoe UI*"")" & vbCrLf
    s = s & "    If Len(achado) = 0 Then achado = Dir$(Environ$(""LOCALAPPDATA"") & ""\Microsoft\Windows\Fonts\Segoe UI*"")" & vbCrLf
    s = s & "    If Len(achado) > 0 Then" & vbCrLf
    s = s & "        mFonteTexto = ""Segoe UI""" & vbCrLf
    s = s & "        mFonteTitulo = ""Segoe UI Black""" & vbCrLf
    s = s & "    Else" & vbCrLf
    CodMarcModulo4 = s
End Function

Private Function CodMarcModulo5() As String
    Dim s As String
    s = s & "        mFonteTexto = ""Segoe UI""" & vbCrLf
    s = s & "        mFonteTitulo = ""Segoe UI Black""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodMarcModulo5 = s
End Function

Private Function CodigoMarcClique() As String
    Dim s As String
    s = s & CodMarcClique1()
    CodigoMarcClique = s
End Function

Private Function CodMarcClique1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' cMarcaClique - repassa clique e duplo clique de um r" & ChrW(243) & "tulo criado por c" & ChrW(243) & "digo:" & vbCrLf
    s = s & "' dono.TratarClique chave   /   dono.TratarDuploClique chave" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private WithEvents mRotulo As MSForms.Label" & vbCrLf
    s = s & "Private mDono As Object" & vbCrLf
    s = s & "Private mChave As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Conectar(ByVal rotulo As MSForms.Label, ByVal chave As String, ByVal dono As Object)" & vbCrLf
    s = s & "    Set mRotulo = rotulo" & vbCrLf
    s = s & "    mChave = chave" & vbCrLf
    s = s & "    Set mDono = dono" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Desconectar()" & vbCrLf
    s = s & "    Set mRotulo = Nothing" & vbCrLf
    s = s & "    Set mDono = Nothing" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub mRotulo_Click()" & vbCrLf
    s = s & "    If Not mDono Is Nothing Then mDono.TratarClique mChave" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub mRotulo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    If Not mDono Is Nothing Then mDono.TratarDuploClique mChave" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodMarcClique1 = s
End Function

Private Function CodigoMarcIcone() As String
    Dim s As String
    s = s & CodMarcIcone1()
    s = s & CodMarcIcone2()
    CodigoMarcIcone = s
End Function

Private Function CodMarcIcone1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' cMarcaIcone - desenha o marcador (fita de marca-p" & ChrW(225) & "gina) com tr" & ChrW(234) & "s r" & ChrW(243) & "tulos:" & vbCrLf
    s = s & "' o corpo e duas ""pernas"" que deixam o recorte em V embaixo." & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mCorpo As MSForms.Label" & vbCrLf
    s = s & "Private mPernaE As MSForms.Label" & vbCrLf
    s = s & "Private mPernaD As MSForms.Label" & vbCrLf
    s = s & "Private mK As Single" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Criar(ByVal pai As Object, ByVal nome As String, ByVal l As Single, ByVal t As Single, _" & vbCrLf
    s = s & "                 ByVal escala As Single, ByVal corInicial As Long)" & vbCrLf
    s = s & "    mK = escala" & vbCrLf
    s = s & "    Set mCorpo = NovaParte(pai, nome & ""C"")" & vbCrLf
    s = s & "    Set mPernaE = NovaParte(pai, nome & ""E"")" & vbCrLf
    s = s & "    Set mPernaD = NovaParte(pai, nome & ""D"")" & vbCrLf
    s = s & "    Mover l, t" & vbCrLf
    s = s & "    Cor corInicial" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function NovaParte(ByVal pai As Object, ByVal nome As String) As MSForms.Label" & vbCrLf
    s = s & "    Dim c As MSForms.Label" & vbCrLf
    s = s & "    Set c = pai.Controls.Add(""Forms.Label.1"", nome, True)" & vbCrLf
    s = s & "    c.Caption = """"" & vbCrLf
    s = s & "    c.BackStyle = 1" & vbCrLf
    s = s & "    c.BorderStyle = 0" & vbCrLf
    s = s & "    Set NovaParte = c" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' 11 x 15 pontos na escala 1" & vbCrLf
    s = s & "Public Sub Mover(ByVal l As Single, ByVal t As Single)" & vbCrLf
    s = s & "    If mCorpo Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    mCorpo.Left = l" & vbCrLf
    s = s & "    mCorpo.Top = t" & vbCrLf
    s = s & "    mCorpo.Width = 11 * mK" & vbCrLf
    s = s & "    mCorpo.Height = 11 * mK" & vbCrLf
    s = s & "    mPernaE.Left = l" & vbCrLf
    s = s & "    mPernaE.Top = t + 11 * mK" & vbCrLf
    s = s & "    mPernaE.Width = 4 * mK" & vbCrLf
    s = s & "    mPernaE.Height = 4 * mK" & vbCrLf
    s = s & "    mPernaD.Left = l + 7 * mK" & vbCrLf
    s = s & "    mPernaD.Top = t + 11 * mK" & vbCrLf
    s = s & "    mPernaD.Width = 4 * mK" & vbCrLf
    s = s & "    mPernaD.Height = 4 * mK" & vbCrLf
    CodMarcIcone1 = s
End Function

Private Function CodMarcIcone2() As String
    Dim s As String
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Cor(ByVal valor As Long)" & vbCrLf
    s = s & "    If mCorpo Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    mCorpo.BackColor = valor" & vbCrLf
    s = s & "    mPernaE.BackColor = valor" & vbCrLf
    s = s & "    mPernaD.BackColor = valor" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Mostrar(ByVal ligado As Boolean)" & vbCrLf
    s = s & "    If mCorpo Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    mCorpo.Visible = ligado" & vbCrLf
    s = s & "    mPernaE.Visible = ligado" & vbCrLf
    s = s & "    mPernaD.Visible = ligado" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Dica(ByVal texto As String)" & vbCrLf
    s = s & "    If mCorpo Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    mCorpo.ControlTipText = texto" & vbCrLf
    s = s & "    mPernaE.ControlTipText = texto" & vbCrLf
    s = s & "    mPernaD.ControlTipText = texto" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Frente()" & vbCrLf
    s = s & "    If mCorpo Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    mCorpo.ZOrder 0" & vbCrLf
    s = s & "    mPernaE.ZOrder 0" & vbCrLf
    s = s & "    mPernaD.ZOrder 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Corpo() As MSForms.Label" & vbCrLf
    s = s & "    Set Corpo = mCorpo" & vbCrLf
    s = s & "End Property" & vbCrLf
    CodMarcIcone2 = s
End Function

Private Function CodigoMarcDados() As String
    Dim s As String
    s = s & CodMarcDados1()
    s = s & CodMarcDados2()
    s = s & CodMarcDados3()
    s = s & CodMarcDados4()
    s = s & CodMarcDados5()
    s = s & CodMarcDados6()
    s = s & CodMarcDados7()
    s = s & CodMarcDados8()
    s = s & CodMarcDados9()
    s = s & CodMarcDados10()
    CodigoMarcDados = s
End Function

Private Function CodMarcDados1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' cMarcadores - leitura e grava" & ChrW(231) & "" & ChrW(227) & "o dos marcadores (aba oculta MARCADORES)" & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "' Um marcador ABERTO por obra para cada pessoa: o do time (todos veem) ou" & vbCrLf
    s = s & "' o ""s" & ChrW(243) & " eu"" de quem marcou. Concluir n" & ChrW(227) & "o apaga: a linha fica com SITUA" & ChrW(199) & "" & ChrW(195) & "O" & vbCrLf
    s = s & "' CONCLU" & ChrW(205) & "DA, quem concluiu e quando." & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const T_MARC As String = ""tb_Marcadores""" & vbCrLf
    s = s & "Private Const ABA_MARC As String = ""MARCADORES""" & vbCrLf
    s = s & "Private Const T_Obras As String = ""tb_Obras""" & vbCrLf
    s = s & "Private Const SIT_ABERTA As String = ""ABERTA""" & vbCrLf
    s = s & "Private Const SIT_CONCLUIDA As String = ""CONCLU" & ChrW(205) & "DA""" & vbCrLf
    s = s & "Private Const VIS_TIME As String = ""TIME""" & vbCrLf
    s = s & "Private Const VIS_SO_EU As String = ""S" & ChrW(211) & " QUEM MARCOU""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mErro As String" & vbCrLf
    s = s & "Private mCache As String" & vbCrLf
    s = s & "Private mTemCache As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get UltimoErro() As String" & vbCrLf
    s = s & "    UltimoErro = mErro" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function Usuario() As String" & vbCrLf
    s = s & "    Usuario = Environ$(""Username"")" & vbCrLf
    s = s & "    If Len(Usuario) = 0 Then Usuario = Application.UserName" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Tabela" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "Private Function Cabecalhos() As Variant" & vbCrLf
    s = s & "    Cabecalhos = Array(""C" & ChrW(211) & "DIGO"", ""OBRA"", ""O QUE ANALISAR"", ""REVISAR EM"", ""QUEM V" & ChrW(202) & """, _" & vbCrLf
    s = s & "                       ""MARCADO POR"", ""MARCADO EM"", ""SITUA" & ChrW(199) & "" & ChrW(195) & "O"", ""CONCLU" & ChrW(205) & "DO POR"", ""CONCLU" & ChrW(205) & "DO EM"")" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function TabelaPorNome(ByVal nome As String) As ListObject" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Dim lo As ListObject" & vbCrLf
    s = s & "    For Each ws In ThisWorkbook.Worksheets" & vbCrLf
    s = s & "        Set lo = Nothing" & vbCrLf
    s = s & "        On Error Resume Next" & vbCrLf
    s = s & "        Set lo = ws.ListObjects(nome)" & vbCrLf
    CodMarcDados1 = s
End Function

Private Function CodMarcDados2() As String
    Dim s As String
    s = s & "        On Error GoTo 0" & vbCrLf
    s = s & "        If Not lo Is Nothing Then" & vbCrLf
    s = s & "            Set TabelaPorNome = lo" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next ws" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Tabela dos marcadores; cria a aba oculta na primeira vez quando criar = True." & vbCrLf
    s = s & "Public Function Tabela(ByVal criar As Boolean) As ListObject" & vbCrLf
    s = s & "    Dim tb As ListObject" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Dim cab As Variant" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim ativa As Object" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set tb = TabelaPorNome(T_MARC)" & vbCrLf
    s = s & "    If tb Is Nothing And criar Then" & vbCrLf
    s = s & "        Set ativa = ActiveSheet" & vbCrLf
    s = s & "        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))" & vbCrLf
    s = s & "        ws.Name = ABA_MARC" & vbCrLf
    s = s & "        cab = Cabecalhos()" & vbCrLf
    s = s & "        For i = 0 To UBound(cab)" & vbCrLf
    s = s & "            ws.Cells(1, i + 1).Value = cab(i)" & vbCrLf
    s = s & "        Next i" & vbCrLf
    s = s & "        Set tb = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(cab) + 1)), , xlYes)" & vbCrLf
    s = s & "        tb.Name = T_MARC" & vbCrLf
    s = s & "        On Error Resume Next" & vbCrLf
    s = s & "        tb.TableStyle = ""TableStyleMedium2""" & vbCrLf
    s = s & "        tb.ListColumns(""C" & ChrW(211) & "DIGO"").DataBodyRange.NumberFormat = ""@""" & vbCrLf
    s = s & "        tb.ListColumns(""REVISAR EM"").DataBodyRange.NumberFormat = ""dd/mm/yyyy""" & vbCrLf
    s = s & "        tb.ListColumns(""MARCADO EM"").DataBodyRange.NumberFormat = ""dd/mm/yyyy hh:mm""" & vbCrLf
    s = s & "        tb.ListColumns(""CONCLU" & ChrW(205) & "DO EM"").DataBodyRange.NumberFormat = ""dd/mm/yyyy hh:mm""" & vbCrLf
    s = s & "        ws.Columns(""A:J"").ColumnWidth = 16" & vbCrLf
    s = s & "        ws.Columns(""C"").ColumnWidth = 50" & vbCrLf
    s = s & "        ws.Visible = xlSheetHidden" & vbCrLf
    s = s & "        If Not ativa Is Nothing Then ativa.Activate" & vbCrLf
    s = s & "        On Error GoTo 0" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Set Tabela = tb" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function Norm(ByVal s As String) As String" & vbCrLf
    s = s & "    Norm = UCase$(Trim$(Replace(Replace(Replace(s, vbCr, "" ""), vbLf, "" ""), Chr$(160), "" "")))" & vbCrLf
    s = s & "End Function" & vbCrLf
    CodMarcDados2 = s
End Function

Private Function CodMarcDados3() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "Private Function Col(ByVal tb As ListObject, ByVal nome As String) As Long" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Function" & vbCrLf
    s = s & "    For i = 1 To tb.ListColumns.Count" & vbCrLf
    s = s & "        If Norm(tb.ListColumns(i).Name) = Norm(nome) Then" & vbCrLf
    s = s & "            Col = i" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function Txt(ByVal v As Variant) As String" & vbCrLf
    s = s & "    If IsError(v) Then Exit Function" & vbCrLf
    s = s & "    If IsEmpty(v) Then Exit Function" & vbCrLf
    s = s & "    Txt = Trim$(Replace(CStr(v), Chr$(160), "" ""))" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function Cel(ByVal tb As ListObject, ByVal i As Long, ByVal nome As String) As Variant" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    c = Col(tb, nome)" & vbCrLf
    s = s & "    If c = 0 Then" & vbCrLf
    s = s & "        Cel = Empty" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Cel = tb.DataBodyRange.Cells(i, c).Value" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Gravar(ByVal tb As ListObject, ByVal i As Long, ByVal nome As String, ByVal valor As Variant, Optional ByVal formato As String = """")" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    c = Col(tb, nome)" & vbCrLf
    s = s & "    If c = 0 Then Exit Sub" & vbCrLf
    s = s & "    If Len(formato) > 0 Then tb.DataBodyRange.Cells(i, c).NumberFormat = formato" & vbCrLf
    s = s & "    If VarType(valor) = vbString Then" & vbCrLf
    s = s & "        If Len(valor) > 0 Then" & vbCrLf
    s = s & "            If InStr(1, ""=+-@"", Left$(valor, 1), vbBinaryCompare) > 0 Then valor = ""'"" & valor" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    tb.DataBodyRange.Cells(i, c).Value = valor" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' A linha " & ChrW(233) & " um marcador aberto que o usu" & ChrW(225) & "rio atual pode ver?" & vbCrLf
    s = s & "Private Function Visivel(ByVal tb As ListObject, ByVal i As Long) As Boolean" & vbCrLf
    s = s & "    If UCase$(Txt(Cel(tb, i, ""SITUA" & ChrW(199) & "" & ChrW(195) & "O""))) <> SIT_ABERTA Then Exit Function" & vbCrLf
    s = s & "    If UCase$(Txt(Cel(tb, i, ""QUEM V" & ChrW(202) & """))) = VIS_SO_EU Then" & vbCrLf
    CodMarcDados3 = s
End Function

Private Function CodMarcDados4() As String
    Dim s As String
    s = s & "        Visivel = (UCase$(Txt(Cel(tb, i, ""MARCADO POR""))) = UCase$(Usuario()))" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Visivel = True" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Linha do marcador aberto da obra vis" & ChrW(237) & "vel para o usu" & ChrW(225) & "rio (0 = n" & ChrW(227) & "o tem)." & vbCrLf
    s = s & "Public Function LinhaAberta(ByVal codObra As String) As Long" & vbCrLf
    s = s & "    Dim tb As ListObject" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim cA As Long" & vbCrLf
    s = s & "    Set tb = Tabela(False)" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Function" & vbCrLf
    s = s & "    If tb.DataBodyRange Is Nothing Then Exit Function" & vbCrLf
    s = s & "    cA = Col(tb, ""C" & ChrW(211) & "DIGO"")" & vbCrLf
    s = s & "    If cA = 0 Then Exit Function" & vbCrLf
    s = s & "    codObra = Trim$(codObra)" & vbCrLf
    s = s & "    If Len(codObra) = 0 Then Exit Function" & vbCrLf
    s = s & "    For i = 1 To tb.ListRows.Count" & vbCrLf
    s = s & "        If Txt(tb.DataBodyRange.Cells(i, cA).Value) = codObra Then" & vbCrLf
    s = s & "            If Visivel(tb, i) Then" & vbCrLf
    s = s & "                LinhaAberta = i" & vbCrLf
    s = s & "                Exit Function" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function Marcada(ByVal codObra As String) As Boolean" & vbCrLf
    s = s & "    If Not mTemCache Then MontarCache" & vbCrLf
    s = s & "    Marcada = (InStr(1, mCache, ""|"" & Trim$(codObra) & ""|"", vbBinaryCompare) > 0)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub MontarCache()" & vbCrLf
    s = s & "    Dim itens As Variant" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    mCache = ""|""" & vbCrLf
    s = s & "    itens = Lista(False)" & vbCrLf
    s = s & "    If IsArray(itens) Then" & vbCrLf
    s = s & "        For i = LBound(itens, 1) To UBound(itens, 1)" & vbCrLf
    s = s & "            mCache = mCache & itens(i, 0) & ""|""" & vbCrLf
    s = s & "        Next i" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    mTemCache = True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodMarcDados4 = s
End Function

Private Function CodMarcDados5() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "Public Sub Recarregar()" & vbCrLf
    s = s & "    mTemCache = False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Dados do marcador aberto da obra (False = obra sem marcador)." & vbCrLf
    s = s & "Public Function Ler(ByVal codObra As String, ByRef texto As String, ByRef revisar As Variant, _" & vbCrLf
    s = s & "                    ByRef soQuemMarcou As Boolean, ByRef por As String, ByRef em As Variant) As Boolean" & vbCrLf
    s = s & "    Dim tb As ListObject" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    texto = """"" & vbCrLf
    s = s & "    revisar = Empty" & vbCrLf
    s = s & "    soQuemMarcou = False" & vbCrLf
    s = s & "    por = """"" & vbCrLf
    s = s & "    em = Empty" & vbCrLf
    s = s & "    i = LinhaAberta(codObra)" & vbCrLf
    s = s & "    If i = 0 Then Exit Function" & vbCrLf
    s = s & "    Set tb = Tabela(False)" & vbCrLf
    s = s & "    texto = Txt(Cel(tb, i, ""O QUE ANALISAR""))" & vbCrLf
    s = s & "    revisar = Cel(tb, i, ""REVISAR EM"")" & vbCrLf
    s = s & "    If Not IsDate(revisar) Then revisar = Empty" & vbCrLf
    s = s & "    soQuemMarcou = (UCase$(Txt(Cel(tb, i, ""QUEM V" & ChrW(202) & """))) = VIS_SO_EU)" & vbCrLf
    s = s & "    por = Txt(Cel(tb, i, ""MARCADO POR""))" & vbCrLf
    s = s & "    em = Cel(tb, i, ""MARCADO EM"")" & vbCrLf
    s = s & "    Ler = True" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Grava" & ChrW(231) & "" & ChrW(227) & "o" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Cria ou atualiza o marcador aberto da obra. revisar = Empty para sem data." & vbCrLf
    s = s & "Public Function Salvar(ByVal codObra As String, ByVal obra As String, ByVal texto As String, _" & vbCrLf
    s = s & "                       ByVal revisar As Variant, ByVal soQuemMarcou As Boolean) As Boolean" & vbCrLf
    s = s & "    Dim tb As ListObject" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim nova As Boolean" & vbCrLf
    s = s & "    Dim ev As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Falha" & vbCrLf
    s = s & "    mErro = """"" & vbCrLf
    s = s & "    codObra = Trim$(codObra)" & vbCrLf
    s = s & "    If Len(codObra) = 0 Then" & vbCrLf
    s = s & "        mErro = ""Obra sem C" & ChrW(211) & "DIGO.""" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    CodMarcDados5 = s
End Function

Private Function CodMarcDados6() As String
    Dim s As String
    s = s & "    ev = Application.EnableEvents" & vbCrLf
    s = s & "    Application.EnableEvents = False" & vbCrLf
    s = s & "    Set tb = Tabela(True)" & vbCrLf
    s = s & "    i = LinhaAberta(codObra)" & vbCrLf
    s = s & "    If i = 0 Then" & vbCrLf
    s = s & "        i = LinhaLivre(tb)" & vbCrLf
    s = s & "        nova = True" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Gravar tb, i, ""C" & ChrW(211) & "DIGO"", codObra, ""@""" & vbCrLf
    s = s & "    Gravar tb, i, ""OBRA"", obra" & vbCrLf
    s = s & "    Gravar tb, i, ""O QUE ANALISAR"", texto" & vbCrLf
    s = s & "    If IsDate(revisar) Then" & vbCrLf
    s = s & "        Gravar tb, i, ""REVISAR EM"", CDate(Int(CDbl(CDate(revisar)))), ""dd/mm/yyyy""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Gravar tb, i, ""REVISAR EM"", Empty" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If soQuemMarcou Then" & vbCrLf
    s = s & "        Gravar tb, i, ""QUEM V" & ChrW(202) & """, VIS_SO_EU" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Gravar tb, i, ""QUEM V" & ChrW(202) & """, VIS_TIME" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If nova Then" & vbCrLf
    s = s & "        Gravar tb, i, ""MARCADO POR"", Usuario()" & vbCrLf
    s = s & "        Gravar tb, i, ""MARCADO EM"", Now, ""dd/mm/yyyy hh:mm""" & vbCrLf
    s = s & "        Gravar tb, i, ""SITUA" & ChrW(199) & "" & ChrW(195) & "O"", SIT_ABERTA" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Application.EnableEvents = ev" & vbCrLf
    s = s & "    mTemCache = False" & vbCrLf
    s = s & "    Salvar = True" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "Falha:" & vbCrLf
    s = s & "    mErro = ""N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel gravar o marcador: "" & Err.Description" & vbCrLf
    s = s & "    Application.EnableEvents = True" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Tira a marca" & ChrW(231) & "" & ChrW(227) & "o: a linha fica como CONCLU" & ChrW(205) & "DA, com quem e quando." & vbCrLf
    s = s & "Public Function Concluir(ByVal codObra As String) As Boolean" & vbCrLf
    s = s & "    Dim tb As ListObject" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim ev As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Falha" & vbCrLf
    s = s & "    mErro = """"" & vbCrLf
    s = s & "    i = LinhaAberta(codObra)" & vbCrLf
    CodMarcDados6 = s
End Function

Private Function CodMarcDados7() As String
    Dim s As String
    s = s & "    If i = 0 Then" & vbCrLf
    s = s & "        mErro = ""Esta obra n" & ChrW(227) & "o tem marcador aberto.""" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    ev = Application.EnableEvents" & vbCrLf
    s = s & "    Application.EnableEvents = False" & vbCrLf
    s = s & "    Set tb = Tabela(False)" & vbCrLf
    s = s & "    Gravar tb, i, ""SITUA" & ChrW(199) & "" & ChrW(195) & "O"", SIT_CONCLUIDA" & vbCrLf
    s = s & "    Gravar tb, i, ""CONCLU" & ChrW(205) & "DO POR"", Usuario()" & vbCrLf
    s = s & "    Gravar tb, i, ""CONCLU" & ChrW(205) & "DO EM"", Now, ""dd/mm/yyyy hh:mm""" & vbCrLf
    s = s & "    Application.EnableEvents = ev" & vbCrLf
    s = s & "    mTemCache = False" & vbCrLf
    s = s & "    Concluir = True" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "Falha:" & vbCrLf
    s = s & "    mErro = ""N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel concluir: "" & Err.Description" & vbCrLf
    s = s & "    Application.EnableEvents = True" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function LinhaLivre(ByVal tb As ListObject) As Long" & vbCrLf
    s = s & "    If tb.ListRows.Count = 1 Then" & vbCrLf
    s = s & "        If Application.WorksheetFunction.CountA(tb.ListRows(1).Range) = 0 Then" & vbCrLf
    s = s & "            LinhaLivre = 1" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    tb.ListRows.Add" & vbCrLf
    s = s & "    LinhaLivre = tb.ListRows.Count" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Lista e resumo" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Marcadores abertos vis" & ChrW(237) & "veis, por prazo (atrasadas primeiro, sem data no fim):" & vbCrLf
    s = s & "' (i, 0) C" & ChrW(211) & "DIGO  (i, 1) OBRA  (i, 2) O QUE ANALISAR  (i, 3) REVISAR EM" & vbCrLf
    s = s & "' (i, 4) MARCADO POR  (i, 5) MARCADO EM  (i, 6) True se s" & ChrW(243) & " quem marcou v" & ChrW(234) & "" & vbCrLf
    s = s & "Public Function Lista(ByVal soMinhas As Boolean) As Variant" & vbCrLf
    s = s & "    Dim tb As ListObject" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim n As Long" & vbCrLf
    s = s & "    Dim k As Long" & vbCrLf
    s = s & "    Dim j As Long" & vbCrLf
    s = s & "    Dim r() As Variant" & vbCrLf
    s = s & "    Dim ordem() As Double" & vbCrLf
    s = s & "    Dim saida() As Variant" & vbCrLf
    CodMarcDados7 = s
End Function

Private Function CodMarcDados8() As String
    Dim s As String
    s = s & "    Dim idx() As Long" & vbCrLf
    s = s & "    Dim t As Long" & vbCrLf
    s = s & "    Dim v As Variant" & vbCrLf
    s = s & "    Dim eu As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set tb = Tabela(False)" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Function" & vbCrLf
    s = s & "    If tb.DataBodyRange Is Nothing Then Exit Function" & vbCrLf
    s = s & "    eu = UCase$(Usuario())" & vbCrLf
    s = s & "    ReDim r(1 To tb.ListRows.Count, 0 To 6)" & vbCrLf
    s = s & "    ReDim ordem(1 To tb.ListRows.Count)" & vbCrLf
    s = s & "    For i = 1 To tb.ListRows.Count" & vbCrLf
    s = s & "        If Len(Txt(Cel(tb, i, ""C" & ChrW(211) & "DIGO""))) > 0 Then" & vbCrLf
    s = s & "            If Visivel(tb, i) Then" & vbCrLf
    s = s & "                If (Not soMinhas) Or UCase$(Txt(Cel(tb, i, ""MARCADO POR""))) = eu Then" & vbCrLf
    s = s & "                    n = n + 1" & vbCrLf
    s = s & "                    r(n, 0) = Txt(Cel(tb, i, ""C" & ChrW(211) & "DIGO""))" & vbCrLf
    s = s & "                    r(n, 1) = Txt(Cel(tb, i, ""OBRA""))" & vbCrLf
    s = s & "                    r(n, 2) = Txt(Cel(tb, i, ""O QUE ANALISAR""))" & vbCrLf
    s = s & "                    v = Cel(tb, i, ""REVISAR EM"")" & vbCrLf
    s = s & "                    If IsDate(v) Then" & vbCrLf
    s = s & "                        r(n, 3) = CDate(v)" & vbCrLf
    s = s & "                        ordem(n) = CDbl(CDate(v))" & vbCrLf
    s = s & "                    Else" & vbCrLf
    s = s & "                        r(n, 3) = Empty" & vbCrLf
    s = s & "                        ordem(n) = 2958465" & vbCrLf
    s = s & "                    End If" & vbCrLf
    s = s & "                    r(n, 4) = Txt(Cel(tb, i, ""MARCADO POR""))" & vbCrLf
    s = s & "                    r(n, 5) = Cel(tb, i, ""MARCADO EM"")" & vbCrLf
    s = s & "                    r(n, 6) = (UCase$(Txt(Cel(tb, i, ""QUEM V" & ChrW(202) & """))) = VIS_SO_EU)" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    If n = 0 Then Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' ordem est" & ChrW(225) & "vel por data de revis" & ChrW(227) & "o" & vbCrLf
    s = s & "    ReDim idx(1 To n)" & vbCrLf
    s = s & "    For i = 1 To n" & vbCrLf
    s = s & "        idx(i) = i" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    For i = 2 To n" & vbCrLf
    s = s & "        t = idx(i)" & vbCrLf
    s = s & "        j = i - 1" & vbCrLf
    s = s & "        Do While j >= 1" & vbCrLf
    CodMarcDados8 = s
End Function

Private Function CodMarcDados9() As String
    Dim s As String
    s = s & "            If ordem(idx(j)) <= ordem(t) Then Exit Do" & vbCrLf
    s = s & "            idx(j + 1) = idx(j)" & vbCrLf
    s = s & "            j = j - 1" & vbCrLf
    s = s & "        Loop" & vbCrLf
    s = s & "        idx(j + 1) = t" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ReDim saida(0 To n - 1, 0 To 6)" & vbCrLf
    s = s & "    For i = 1 To n" & vbCrLf
    s = s & "        For k = 0 To 6" & vbCrLf
    s = s & "            saida(i - 1, k) = r(idx(i), k)" & vbCrLf
    s = s & "        Next k" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    Lista = saida" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Resumo(ByRef total As Long, ByRef hoje As Long, ByRef atrasadas As Long)" & vbCrLf
    s = s & "    Dim itens As Variant" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    total = 0" & vbCrLf
    s = s & "    hoje = 0" & vbCrLf
    s = s & "    atrasadas = 0" & vbCrLf
    s = s & "    itens = Lista(False)" & vbCrLf
    s = s & "    If Not IsArray(itens) Then Exit Sub" & vbCrLf
    s = s & "    For i = LBound(itens, 1) To UBound(itens, 1)" & vbCrLf
    s = s & "        total = total + 1" & vbCrLf
    s = s & "        Select Case TipoPrazo(itens(i, 3))" & vbCrLf
    s = s & "            Case PZ_HOJE" & vbCrLf
    s = s & "                hoje = hoje + 1" & vbCrLf
    s = s & "            Case PZ_ATRASADA" & vbCrLf
    s = s & "                atrasadas = atrasadas + 1" & vbCrLf
    s = s & "        End Select" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' follow up" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "Public Function LinhaNoFollowUp(ByVal codObra As String) As Long" & vbCrLf
    s = s & "    LinhaNoFollowUp = LinhaEm(T_Obras, codObra)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function LinhaEm(ByVal nomeTabela As String, ByVal codObra As String) As Long" & vbCrLf
    s = s & "    Dim tb As ListObject" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    CodMarcDados9 = s
End Function

Private Function CodMarcDados10() As String
    Dim s As String
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Set tb = TabelaPorNome(nomeTabela)" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Function" & vbCrLf
    s = s & "    If tb.DataBodyRange Is Nothing Then Exit Function" & vbCrLf
    s = s & "    c = Col(tb, ""C" & ChrW(211) & "DIGO"")" & vbCrLf
    s = s & "    If c = 0 Then Exit Function" & vbCrLf
    s = s & "    For i = 1 To tb.ListRows.Count" & vbCrLf
    s = s & "        If Txt(tb.DataBodyRange.Cells(i, c).Value) = Trim$(codObra) Then" & vbCrLf
    s = s & "            LinhaEm = i" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Function" & vbCrLf
    CodMarcDados10 = s
End Function

Private Function CodigoMarcForm() As String
    Dim s As String
    s = s & CodMarcForm1()
    s = s & CodMarcForm2()
    s = s & CodMarcForm3()
    s = s & CodMarcForm4()
    s = s & CodMarcForm5()
    s = s & CodMarcForm6()
    s = s & CodMarcForm7()
    s = s & CodMarcForm8()
    s = s & CodMarcForm9()
    CodigoMarcForm = s
End Function

Private Function CodMarcForm1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' frmMarcador - marcar a obra para analisar depois (marca-p" & ChrW(225) & "gina)" & vbCrLf
    s = s & "' UserForm VAZIO: os controles s" & ChrW(227) & "o criados por c" & ChrW(243) & "digo. Sem caixas de mensagem." & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const DW As Single = 440" & vbCrLf
    s = s & "Private Const DH As Single = 292" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private WithEvents txtData As MSForms.TextBox" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mMarcas As cMarcadores" & vbCrLf
    s = s & "Private mCliques As Collection" & vbCrLf
    s = s & "Private mFonte As String" & vbCrLf
    s = s & "Private mFonteTit As String" & vbCrLf
    s = s & "Private mCarregando As Boolean" & vbCrLf
    s = s & "Private mCod As String" & vbCrLf
    s = s & "Private mObra As String" & vbCrLf
    s = s & "Private mExiste As Boolean" & vbCrLf
    s = s & "Private mResultado As Long" & vbCrLf
    s = s & "Private mPreset As Long" & vbCrLf
    s = s & "Private mDataOrig As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private lblTitulo As MSForms.Label" & vbCrLf
    s = s & "Private txtTexto As MSForms.TextBox" & vbCrLf
    s = s & "Private mChipF(0 To 3) As MSForms.Label" & vbCrLf
    s = s & "Private mChipT(0 To 3) As MSForms.Label" & vbCrLf
    s = s & "Private lblDataInfo As MSForms.Label" & vbCrLf
    s = s & "Private optTime As MSForms.OptionButton" & vbCrLf
    s = s & "Private optEu As MSForms.OptionButton" & vbCrLf
    s = s & "Private lblQuem As MSForms.Label" & vbCrLf
    s = s & "Private lblSalvarT As MSForms.Label" & vbCrLf
    s = s & "Private lblConcF As MSForms.Label" & vbCrLf
    s = s & "Private lblConcT As MSForms.Label" & vbCrLf
    s = s & "Private lblStatus As MSForms.Label" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' CICLO DE VIDA" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    mFonte = FonteMarcador()" & vbCrLf
    s = s & "    mFonteTit = FonteTituloMarcador()" & vbCrLf
    s = s & "    Set mCliques = New Collection" & vbCrLf
    s = s & "    Set mMarcas = New cMarcadores" & vbCrLf
    s = s & "    mPreset = 0" & vbCrLf
    CodMarcForm1 = s
End Function

Private Function CodMarcForm2() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "    Me.Caption = ""Marcar para analisar""" & vbCrLf
    s = s & "    Me.BackColor = MC_BRANCO" & vbCrLf
    s = s & "    Me.Width = DW + 12" & vbCrLf
    s = s & "    Me.Height = DH + 30" & vbCrLf
    s = s & "    Me.Width = DW + (Me.Width - Me.InsideWidth)" & vbCrLf
    s = s & "    Me.Height = DH + (Me.Height - Me.InsideHeight)" & vbCrLf
    s = s & "    mCarregando = True" & vbCrLf
    s = s & "    MontarTela" & vbCrLf
    s = s & "    mCarregando = False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    Dim ev As cMarcaClique" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If mCliques Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    For Each ev In mCliques" & vbCrLf
    s = s & "        ev.Desconectar" & vbCrLf
    s = s & "    Next ev" & vbCrLf
    s = s & "    Set mCliques = Nothing" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Resultado() As Long" & vbCrLf
    s = s & "    Resultado = mResultado" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Preparar(ByVal codObra As String, ByVal obra As String)" & vbCrLf
    s = s & "    Dim texto As String" & vbCrLf
    s = s & "    Dim revisar As Variant" & vbCrLf
    s = s & "    Dim soEu As Boolean" & vbCrLf
    s = s & "    Dim por As String" & vbCrLf
    s = s & "    Dim em As Variant" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mCod = Trim$(codObra)" & vbCrLf
    s = s & "    mObra = Trim$(obra)" & vbCrLf
    s = s & "    mResultado = 0" & vbCrLf
    s = s & "    lblTitulo.Caption = ""MARCAR PARA ANALISAR  "" & ChrW(183) & ""  "" & UCase$(mObra)" & vbCrLf
    s = s & "    optEu.Caption = ""S" & ChrW(243) & " eu ("" & mMarcas.Usuario() & "")""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mCarregando = True" & vbCrLf
    s = s & "    mExiste = mMarcas.Ler(mCod, texto, revisar, soEu, por, em)" & vbCrLf
    CodMarcForm2 = s
End Function

Private Function CodMarcForm3() As String
    Dim s As String
    s = s & "    If mExiste Then" & vbCrLf
    s = s & "        txtTexto.Text = texto" & vbCrLf
    s = s & "        If IsDate(revisar) Then" & vbCrLf
    s = s & "            txtData.Text = Format$(CDate(revisar), ""dd/mm/yyyy"")" & vbCrLf
    s = s & "            mPreset = -1" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            txtData.Text = """"" & vbCrLf
    s = s & "            mPreset = 0" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        optEu.Value = soEu" & vbCrLf
    s = s & "        optTime.Value = Not soEu" & vbCrLf
    s = s & "        lblQuem.Caption = ""Marcada por "" & por & IIf(IsDate(em), "" em "" & Format$(em, ""dd/mm/yyyy hh:mm""), """")" & vbCrLf
    s = s & "        lblSalvarT.Caption = ""Salvar altera" & ChrW(231) & "" & ChrW(245) & "es""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        txtTexto.Text = """"" & vbCrLf
    s = s & "        txtData.Text = """"" & vbCrLf
    s = s & "        mPreset = 0" & vbCrLf
    s = s & "        optTime.Value = True" & vbCrLf
    s = s & "        lblQuem.Caption = """"" & vbCrLf
    s = s & "        lblSalvarT.Caption = ""Salvar marcador""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    lblConcF.Visible = mExiste" & vbCrLf
    s = s & "    lblConcT.Visible = mExiste" & vbCrLf
    s = s & "    mDataOrig = txtData.Text" & vbCrLf
    s = s & "    mCarregando = False" & vbCrLf
    s = s & "    PintarChips" & vbCrLf
    s = s & "    AtualizarDataInfo" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' TELA" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "Private Function Rot(ByVal nome As String, ByVal texto As String, _" & vbCrLf
    s = s & "                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _" & vbCrLf
    s = s & "                     ByVal cor As Long, ByVal tam As Single, ByVal negrito As Boolean, _" & vbCrLf
    s = s & "                     Optional ByVal fundo As Long = -1, Optional ByVal alinh As Long = 1) As MSForms.Label" & vbCrLf
    s = s & "    Dim c As MSForms.Label" & vbCrLf
    s = s & "    Set c = Me.Controls.Add(""Forms.Label.1"", nome, True)" & vbCrLf
    s = s & "    c.Left = l" & vbCrLf
    s = s & "    c.Top = t" & vbCrLf
    s = s & "    c.Width = w" & vbCrLf
    s = s & "    c.Height = h" & vbCrLf
    s = s & "    c.Font.Name = mFonte" & vbCrLf
    s = s & "    c.Font.Size = tam" & vbCrLf
    s = s & "    c.Font.Bold = negrito" & vbCrLf
    CodMarcForm3 = s
End Function

Private Function CodMarcForm4() As String
    Dim s As String
    s = s & "    c.ForeColor = cor" & vbCrLf
    s = s & "    If fundo >= 0 Then" & vbCrLf
    s = s & "        c.BackStyle = 1" & vbCrLf
    s = s & "        c.BackColor = fundo" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        c.BackStyle = 0" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    c.TextAlign = alinh" & vbCrLf
    s = s & "    c.WordWrap = False" & vbCrLf
    s = s & "    c.Caption = texto" & vbCrLf
    s = s & "    Set Rot = c" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Clique(ByVal rotulo As MSForms.Label, ByVal chave As String)" & vbCrLf
    s = s & "    Dim ev As cMarcaClique" & vbCrLf
    s = s & "    Set ev = New cMarcaClique" & vbCrLf
    s = s & "    ev.Conectar rotulo, chave, Me" & vbCrLf
    s = s & "    mCliques.Add ev" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Botao(ByVal nome As String, ByVal chave As String, ByVal texto As String, _" & vbCrLf
    s = s & "                  ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _" & vbCrLf
    s = s & "                  ByVal fundo As Long, ByVal cor As Long, ByVal borda As Long, _" & vbCrLf
    s = s & "                  ByRef rotFundo As MSForms.Label, ByRef rotTexto As MSForms.Label)" & vbCrLf
    s = s & "    Set rotFundo = Rot(nome & ""F"", """", l, t, w, h, cor, 8, False, fundo)" & vbCrLf
    s = s & "    If borda >= 0 Then" & vbCrLf
    s = s & "        rotFundo.BorderStyle = 1" & vbCrLf
    s = s & "        rotFundo.BorderColor = borda" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Set rotTexto = Rot(nome & ""T"", texto, l, t + (h - 12) / 2 - 0.5, w, 12, cor, 8.5, True, -1, 2)" & vbCrLf
    s = s & "    Clique rotFundo, chave" & vbCrLf
    s = s & "    Clique rotTexto, chave" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub MontarTela()" & vbCrLf
    s = s & "    Dim c As MSForms.Label" & vbCrLf
    s = s & "    Dim f As MSForms.Label" & vbCrLf
    s = s & "    Dim t As MSForms.Label" & vbCrLf
    s = s & "    Dim ic As cMarcaIcone" & vbCrLf
    s = s & "    Dim nomes As Variant" & vbCrLf
    s = s & "    Dim larg As Variant" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim x As Single" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Rot ""lblFaixa"", """", 0, 0, DW, 30, MC_BRANCO, 8, False, MC_NAVY" & vbCrLf
    CodMarcForm4 = s
End Function

Private Function CodMarcForm5() As String
    Dim s As String
    s = s & "    Rot ""lblFaixaAm"", """", 0, 30, DW, 3, MC_BRANCO, 8, False, MC_AMARELO" & vbCrLf
    s = s & "    Set ic = New cMarcaIcone" & vbCrLf
    s = s & "    ic.Criar Me, ""icoTopo"", 14, 8, 1, MC_AMARELO" & vbCrLf
    s = s & "    Set lblTitulo = Rot(""lblTitulo"", """", 32, 8, DW - 44, 16, MC_BRANCO, 10, False)" & vbCrLf
    s = s & "    lblTitulo.Font.Name = mFonteTit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Rot ""lblLab1"", ""O QUE PRECISA SER ANALISADO"", 16, 44, 300, 12, MC_CINZA, 7.5, True" & vbCrLf
    s = s & "    Set txtTexto = Me.Controls.Add(""Forms.TextBox.1"", ""txtTexto"", True)" & vbCrLf
    s = s & "    txtTexto.Left = 16" & vbCrLf
    s = s & "    txtTexto.Top = 58" & vbCrLf
    s = s & "    txtTexto.Width = DW - 32" & vbCrLf
    s = s & "    txtTexto.Height = 52" & vbCrLf
    s = s & "    txtTexto.MultiLine = True" & vbCrLf
    s = s & "    txtTexto.WordWrap = True" & vbCrLf
    s = s & "    txtTexto.EnterKeyBehavior = True" & vbCrLf
    s = s & "    txtTexto.MaxLength = 400" & vbCrLf
    s = s & "    txtTexto.SpecialEffect = 0" & vbCrLf
    s = s & "    txtTexto.BorderStyle = 1" & vbCrLf
    s = s & "    txtTexto.BorderColor = MC_BORDA" & vbCrLf
    s = s & "    txtTexto.BackColor = MC_EDICAO" & vbCrLf
    s = s & "    txtTexto.ForeColor = MC_NAVY" & vbCrLf
    s = s & "    txtTexto.Font.Name = mFonte" & vbCrLf
    s = s & "    txtTexto.Font.Size = 9.5" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Rot ""lblLab2"", ""REVISAR EM"", 16, 120, 200, 12, MC_CINZA, 7.5, True" & vbCrLf
    s = s & "    nomes = Array(""Sem data"", ""Amanh" & ChrW(227) & """, ""Em 2 dias"", ""Em 1 semana"")" & vbCrLf
    s = s & "    larg = Array(58, 58, 62, 74)" & vbCrLf
    s = s & "    x = 16" & vbCrLf
    s = s & "    For i = 0 To 3" & vbCrLf
    s = s & "        Botao ""chip"" & i, ""chip:"" & i, CStr(nomes(i)), x, 134, CSng(larg(i)), 20, MC_BRANCO, MC_NAVY, MC_BORDA, f, t" & vbCrLf
    s = s & "        Set mChipF(i) = f" & vbCrLf
    s = s & "        Set mChipT(i) = t" & vbCrLf
    s = s & "        x = x + CSng(larg(i)) + 6" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    Set txtData = Me.Controls.Add(""Forms.TextBox.1"", ""txtData"", True)" & vbCrLf
    s = s & "    txtData.Left = x" & vbCrLf
    s = s & "    txtData.Top = 134" & vbCrLf
    s = s & "    txtData.Width = 82" & vbCrLf
    s = s & "    txtData.Height = 20" & vbCrLf
    s = s & "    txtData.MaxLength = 10" & vbCrLf
    s = s & "    txtData.SpecialEffect = 0" & vbCrLf
    s = s & "    txtData.BorderStyle = 1" & vbCrLf
    s = s & "    txtData.BorderColor = MC_BORDA" & vbCrLf
    s = s & "    txtData.BackColor = MC_EDICAO" & vbCrLf
    s = s & "    txtData.ForeColor = MC_NAVY" & vbCrLf
    CodMarcForm5 = s
End Function

Private Function CodMarcForm6() As String
    Dim s As String
    s = s & "    txtData.Font.Name = mFonte" & vbCrLf
    s = s & "    txtData.Font.Size = 9" & vbCrLf
    s = s & "    txtData.ControlTipText = ""Outra data: dd/mm/aaaa (ou s" & ChrW(243) & " os n" & ChrW(250) & "meros, 15102026)""" & vbCrLf
    s = s & "    Set lblDataInfo = Rot(""lblDataInfo"", """", 16, 158, DW - 32, 12, MC_CINZA, 8, False)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Rot ""lblLab3"", ""QUEM V" & ChrW(202) & """, 16, 180, 200, 12, MC_CINZA, 7.5, True" & vbCrLf
    s = s & "    Set optTime = NovaOpcao(""optTime"", ""Todo o time"", 16, 193, 110)" & vbCrLf
    s = s & "    Set optEu = NovaOpcao(""optEu"", ""S" & ChrW(243) & " eu"", 130, 193, 220)" & vbCrLf
    s = s & "    Set lblQuem = Rot(""lblQuem"", """", 16, 216, DW - 32, 12, MC_CINZA, 7.5, False)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Botao ""btnSalvar"", ""salvar"", ""Salvar marcador"", 16, 236, 196, 26, MC_NAVY, MC_BRANCO, -1, f, lblSalvarT" & vbCrLf
    s = s & "    Botao ""btnCancelar"", ""cancelar"", ""Cancelar"", 218, 236, 96, 26, MC_BRANCO, MC_NAVY, MC_BORDA, f, t" & vbCrLf
    s = s & "    Botao ""btnConcluir"", ""concluir"", ""Concluir an" & ChrW(225) & "lise"", 320, 236, 104, 26, MC_BRANCO, MC_CORAL, MC_CORAL, lblConcF, lblConcT" & vbCrLf
    s = s & "    lblConcF.ControlTipText = ""Tira a marca" & ChrW(231) & "" & ChrW(227) & "o; fica guardado quem concluiu e quando""" & vbCrLf
    s = s & "    lblConcT.ControlTipText = lblConcF.ControlTipText" & vbCrLf
    s = s & "    Set lblStatus = Rot(""lblStatus"", """", 16, 270, DW - 32, 14, MC_CORAL, 8, False)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function NovaOpcao(ByVal nome As String, ByVal texto As String, ByVal l As Single, ByVal t As Single, ByVal w As Single) As MSForms.OptionButton" & vbCrLf
    s = s & "    Dim o As MSForms.OptionButton" & vbCrLf
    s = s & "    Set o = Me.Controls.Add(""Forms.OptionButton.1"", nome, True)" & vbCrLf
    s = s & "    o.Left = l" & vbCrLf
    s = s & "    o.Top = t" & vbCrLf
    s = s & "    o.Width = w" & vbCrLf
    s = s & "    o.Height = 18" & vbCrLf
    s = s & "    o.Caption = texto" & vbCrLf
    s = s & "    o.BackStyle = 0" & vbCrLf
    s = s & "    o.ForeColor = MC_NAVY" & vbCrLf
    s = s & "    o.Font.Name = mFonte" & vbCrLf
    s = s & "    o.Font.Size = 9" & vbCrLf
    s = s & "    o.GroupName = ""quem""" & vbCrLf
    s = s & "    Set NovaOpcao = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub PintarChips()" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    For i = 0 To 3" & vbCrLf
    s = s & "        If i = mPreset Then" & vbCrLf
    s = s & "            mChipF(i).BackColor = MC_AZUL" & vbCrLf
    s = s & "            mChipF(i).BorderColor = MC_AZUL" & vbCrLf
    s = s & "            mChipT(i).ForeColor = MC_BRANCO" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            mChipF(i).BackColor = MC_BRANCO" & vbCrLf
    s = s & "            mChipF(i).BorderColor = MC_BORDA" & vbCrLf
    s = s & "            mChipT(i).ForeColor = MC_NAVY" & vbCrLf
    CodMarcForm6 = s
End Function

Private Function CodMarcForm7() As String
    Dim s As String
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub AtualizarDataInfo()" & vbCrLf
    s = s & "    Dim d As Date" & vbCrLf
    s = s & "    If Len(Trim$(txtData.Text)) = 0 Then" & vbCrLf
    s = s & "        lblDataInfo.ForeColor = MC_CINZA" & vbCrLf
    s = s & "        lblDataInfo.Caption = ""Sem data de revis" & ChrW(227) & "o: a obra fica marcada at" & ChrW(233) & " algu" & ChrW(233) & "m concluir a an" & ChrW(225) & "lise.""" & vbCrLf
    s = s & "    ElseIf LerData(txtData.Text, d) Then" & vbCrLf
    s = s & "        lblDataInfo.ForeColor = MC_CINZA" & vbCrLf
    s = s & "        lblDataInfo.Caption = ""= "" & Format$(d, ""dd/mm/yyyy"") & "" ("" & DiaDaSemana(d) & "")""" & vbCrLf
    s = s & "        If d < Date Then" & vbCrLf
    s = s & "            lblDataInfo.ForeColor = MC_CORAL" & vbCrLf
    s = s & "            lblDataInfo.Caption = lblDataInfo.Caption & ""  -  esta data j" & ChrW(225) & " passou""" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblDataInfo.ForeColor = MC_CORAL" & vbCrLf
    s = s & "        lblDataInfo.Caption = ""Data inv" & ChrW(225) & "lida. Use dd/mm/aaaa.""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' EVENTOS" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "Private Sub txtData_Change()" & vbCrLf
    s = s & "    If mCarregando Then Exit Sub" & vbCrLf
    s = s & "    mPreset = -1" & vbCrLf
    s = s & "    If Len(Trim$(txtData.Text)) = 0 Then mPreset = 0" & vbCrLf
    s = s & "    PintarChips" & vbCrLf
    s = s & "    AtualizarDataInfo" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub TratarDuploClique(ByVal chave As String)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub TratarClique(ByVal chave As String)" & vbCrLf
    s = s & "    Dim n As Long" & vbCrLf
    s = s & "    If Left$(chave, 5) = ""chip:"" Then" & vbCrLf
    s = s & "        n = CLng(Mid$(chave, 6))" & vbCrLf
    s = s & "        mCarregando = True" & vbCrLf
    s = s & "        Select Case n" & vbCrLf
    s = s & "            Case 0" & vbCrLf
    s = s & "                txtData.Text = """"" & vbCrLf
    s = s & "            Case 1" & vbCrLf
    CodMarcForm7 = s
End Function

Private Function CodMarcForm8() As String
    Dim s As String
    s = s & "                txtData.Text = Format$(Date + 1, ""dd/mm/yyyy"")" & vbCrLf
    s = s & "            Case 2" & vbCrLf
    s = s & "                txtData.Text = Format$(Date + 2, ""dd/mm/yyyy"")" & vbCrLf
    s = s & "            Case 3" & vbCrLf
    s = s & "                txtData.Text = Format$(Date + 7, ""dd/mm/yyyy"")" & vbCrLf
    s = s & "        End Select" & vbCrLf
    s = s & "        mCarregando = False" & vbCrLf
    s = s & "        mPreset = n" & vbCrLf
    s = s & "        PintarChips" & vbCrLf
    s = s & "        AtualizarDataInfo" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Select Case chave" & vbCrLf
    s = s & "        Case ""cancelar""" & vbCrLf
    s = s & "            Me.Hide" & vbCrLf
    s = s & "        Case ""salvar""" & vbCrLf
    s = s & "            Salvar" & vbCrLf
    s = s & "        Case ""concluir""" & vbCrLf
    s = s & "            If mMarcas.Concluir(mCod) Then" & vbCrLf
    s = s & "                mResultado = 2" & vbCrLf
    s = s & "                Me.Hide" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                lblStatus.Caption = mMarcas.UltimoErro" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Salvar()" & vbCrLf
    s = s & "    Dim d As Date" & vbCrLf
    s = s & "    Dim revisar As Variant" & vbCrLf
    s = s & "    Dim texto As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    texto = Trim$(txtTexto.Text)" & vbCrLf
    s = s & "    If Len(texto) = 0 Then" & vbCrLf
    s = s & "        lblStatus.Caption = ""Escreva o que precisa ser analisado.""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    revisar = Empty" & vbCrLf
    s = s & "    If Len(Trim$(txtData.Text)) > 0 Then" & vbCrLf
    s = s & "        If Not LerData(txtData.Text, d) Then" & vbCrLf
    s = s & "            lblStatus.Caption = ""A data de revis" & ChrW(227) & "o n" & ChrW(227) & "o " & ChrW(233) & " v" & ChrW(225) & "lida. Use dd/mm/aaaa ou deixe em branco.""" & vbCrLf
    s = s & "            Exit Sub" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        If d < Date And txtData.Text <> mDataOrig Then" & vbCrLf
    CodMarcForm8 = s
End Function

Private Function CodMarcForm9() As String
    Dim s As String
    s = s & "            lblStatus.Caption = ""A data de revis" & ChrW(227) & "o j" & ChrW(225) & " passou. Escolha hoje ou uma data futura.""" & vbCrLf
    s = s & "            Exit Sub" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        revisar = d" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If mMarcas.Salvar(mCod, mObra, texto, revisar, optEu.Value) Then" & vbCrLf
    s = s & "        mResultado = 1" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblStatus.Caption = mMarcas.UltimoErro" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodMarcForm9 = s
End Function

Private Function CodigoMarcLista() As String
    Dim s As String
    s = s & CodMarcLista1()
    s = s & CodMarcLista2()
    s = s & CodMarcLista3()
    s = s & CodMarcLista4()
    s = s & CodMarcLista5()
    s = s & CodMarcLista6()
    s = s & CodMarcLista7()
    s = s & CodMarcLista8()
    CodigoMarcLista = s
End Function

Private Function CodMarcLista1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' frmMarcadas - obras marcadas para analisar (atrasadas primeiro)" & vbCrLf
    s = s & "' UserForm VAZIO: os controles s" & ChrW(227) & "o criados por c" & ChrW(243) & "digo. Sem caixas de mensagem." & vbCrLf
    s = s & "' Clique seleciona; duplo clique (ou Abrir obra) abre a obra no painel." & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const DW As Single = 720" & vbCrLf
    s = s & "Private Const DH As Single = 424" & vbCrLf
    s = s & "Private Const ALT_LINHA As Single = 36" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mMarcas As cMarcadores" & vbCrLf
    s = s & "Private mCliques As Collection" & vbCrLf
    s = s & "Private mCliquesLinhas As Collection" & vbCrLf
    s = s & "Private mFonte As String" & vbCrLf
    s = s & "Private mFonteTit As String" & vbCrLf
    s = s & "Private mDados As Variant" & vbCrLf
    s = s & "Private mQtd As Long" & vbCrLf
    s = s & "Private mSel As Long" & vbCrLf
    s = s & "Private mSoMinhas As Boolean" & vbCrLf
    s = s & "Private mEscolhida As String" & vbCrLf
    s = s & "Private mFundos() As MSForms.Label" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private fraLista As MSForms.Frame" & vbCrLf
    s = s & "Private lblTitulo As MSForms.Label" & vbCrLf
    s = s & "Private lblMinhasF As MSForms.Label" & vbCrLf
    s = s & "Private lblMinhasT As MSForms.Label" & vbCrLf
    s = s & "Private lblStatus As MSForms.Label" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' CICLO DE VIDA" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    mFonte = FonteMarcador()" & vbCrLf
    s = s & "    mFonteTit = FonteTituloMarcador()" & vbCrLf
    s = s & "    Set mCliques = New Collection" & vbCrLf
    s = s & "    Set mCliquesLinhas = New Collection" & vbCrLf
    s = s & "    Set mMarcas = New cMarcadores" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Me.Caption = ""Obras marcadas para analisar""" & vbCrLf
    s = s & "    Me.BackColor = MC_FUNDO" & vbCrLf
    s = s & "    Me.Width = DW + 12" & vbCrLf
    s = s & "    Me.Height = DH + 30" & vbCrLf
    s = s & "    Me.Width = DW + (Me.Width - Me.InsideWidth)" & vbCrLf
    s = s & "    Me.Height = DH + (Me.Height - Me.InsideHeight)" & vbCrLf
    CodMarcLista1 = s
End Function

Private Function CodMarcLista2() As String
    Dim s As String
    s = s & "    MontarTela" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Desligar mCliques" & vbCrLf
    s = s & "    Desligar mCliquesLinhas" & vbCrLf
    s = s & "    Set mCliques = Nothing" & vbCrLf
    s = s & "    Set mCliquesLinhas = Nothing" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Desligar(ByVal col As Collection)" & vbCrLf
    s = s & "    Dim ev As cMarcaClique" & vbCrLf
    s = s & "    If col Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    For Each ev In col" & vbCrLf
    s = s & "        ev.Desconectar" & vbCrLf
    s = s & "    Next ev" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Escolhida() As String" & vbCrLf
    s = s & "    Escolhida = mEscolhida" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Preparar()" & vbCrLf
    s = s & "    mEscolhida = """"" & vbCrLf
    s = s & "    mSoMinhas = False" & vbCrLf
    s = s & "    PintarMinhas" & vbCrLf
    s = s & "    Carregar" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' TELA" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "Private Function Rot(ByVal pai As Object, ByVal nome As String, ByVal texto As String, _" & vbCrLf
    s = s & "                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _" & vbCrLf
    s = s & "                     ByVal cor As Long, ByVal tam As Single, ByVal negrito As Boolean, _" & vbCrLf
    s = s & "                     Optional ByVal fundo As Long = -1, Optional ByVal alinh As Long = 1) As MSForms.Label" & vbCrLf
    s = s & "    Dim c As MSForms.Label" & vbCrLf
    s = s & "    Set c = pai.Controls.Add(""Forms.Label.1"", nome, True)" & vbCrLf
    s = s & "    c.Left = l" & vbCrLf
    s = s & "    c.Top = t" & vbCrLf
    CodMarcLista2 = s
End Function

Private Function CodMarcLista3() As String
    Dim s As String
    s = s & "    c.Width = w" & vbCrLf
    s = s & "    c.Height = h" & vbCrLf
    s = s & "    c.Font.Name = mFonte" & vbCrLf
    s = s & "    c.Font.Size = tam" & vbCrLf
    s = s & "    c.Font.Bold = negrito" & vbCrLf
    s = s & "    c.ForeColor = cor" & vbCrLf
    s = s & "    If fundo >= 0 Then" & vbCrLf
    s = s & "        c.BackStyle = 1" & vbCrLf
    s = s & "        c.BackColor = fundo" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        c.BackStyle = 0" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    c.TextAlign = alinh" & vbCrLf
    s = s & "    c.WordWrap = False" & vbCrLf
    s = s & "    c.Caption = texto" & vbCrLf
    s = s & "    Set Rot = c" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Clique(ByVal col As Collection, ByVal rotulo As MSForms.Label, ByVal chave As String)" & vbCrLf
    s = s & "    Dim ev As cMarcaClique" & vbCrLf
    s = s & "    Set ev = New cMarcaClique" & vbCrLf
    s = s & "    ev.Conectar rotulo, chave, Me" & vbCrLf
    s = s & "    col.Add ev" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Botao(ByVal nome As String, ByVal chave As String, ByVal texto As String, _" & vbCrLf
    s = s & "                  ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _" & vbCrLf
    s = s & "                  ByVal fundo As Long, ByVal cor As Long, ByVal borda As Long, _" & vbCrLf
    s = s & "                  ByRef rotFundo As MSForms.Label, ByRef rotTexto As MSForms.Label)" & vbCrLf
    s = s & "    Set rotFundo = Rot(Me, nome & ""F"", """", l, t, w, h, cor, 8, False, fundo)" & vbCrLf
    s = s & "    If borda >= 0 Then" & vbCrLf
    s = s & "        rotFundo.BorderStyle = 1" & vbCrLf
    s = s & "        rotFundo.BorderColor = borda" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Set rotTexto = Rot(Me, nome & ""T"", texto, l, t + (h - 12) / 2 - 0.5, w, 12, cor, 8.5, True, -1, 2)" & vbCrLf
    s = s & "    Clique mCliques, rotFundo, chave" & vbCrLf
    s = s & "    Clique mCliques, rotTexto, chave" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Legenda(ByVal nome As String, ByVal texto As String, ByVal x As Single, ByVal cor As Long)" & vbCrLf
    s = s & "    Dim ic As cMarcaIcone" & vbCrLf
    s = s & "    Set ic = New cMarcaIcone" & vbCrLf
    s = s & "    ic.Criar Me, nome, x, 357, 0.7, cor" & vbCrLf
    s = s & "    Rot Me, nome & ""T"", texto, x + 11, 356, 110, 12, MC_CINZA, 8, False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodMarcLista3 = s
End Function

Private Function CodMarcLista4() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "Private Sub MontarTela()" & vbCrLf
    s = s & "    Dim ic As cMarcaIcone" & vbCrLf
    s = s & "    Dim f As MSForms.Label" & vbCrLf
    s = s & "    Dim t As MSForms.Label" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Rot Me, ""lblFaixa"", """", 0, 0, DW, 30, MC_BRANCO, 8, False, MC_NAVY" & vbCrLf
    s = s & "    Rot Me, ""lblFaixaAm"", """", 0, 30, DW, 3, MC_BRANCO, 8, False, MC_AMARELO" & vbCrLf
    s = s & "    Set ic = New cMarcaIcone" & vbCrLf
    s = s & "    ic.Criar Me, ""icoTopo"", 14, 8, 1, MC_AMARELO" & vbCrLf
    s = s & "    Set lblTitulo = Rot(Me, ""lblTitulo"", ""OBRAS MARCADAS"", 32, 8, 400, 16, MC_BRANCO, 10.5, False)" & vbCrLf
    s = s & "    lblTitulo.Font.Name = mFonteTit" & vbCrLf
    s = s & "    Rot Me, ""lblDica"", ""clique seleciona  "" & ChrW(183) & ""  duplo clique abre a obra"", DW - 260, 10, 248, 12, MC_TEXTO_CLARO, 8, False, -1, 3" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Rot Me, ""lblC1"", ""OBRA"", 40, 42, 180, 12, MC_CINZA, 7.5, True" & vbCrLf
    s = s & "    Rot Me, ""lblC2"", ""O QUE ANALISAR"", 232, 42, 260, 12, MC_CINZA, 7.5, True" & vbCrLf
    s = s & "    Rot Me, ""lblC3"", ""REVISAR"", 500, 42, 100, 12, MC_CINZA, 7.5, True" & vbCrLf
    s = s & "    Rot Me, ""lblC4"", ""MARCADA POR"", 600, 42, 100, 12, MC_CINZA, 7.5, True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set fraLista = Me.Controls.Add(""Forms.Frame.1"", ""fraLista"", True)" & vbCrLf
    s = s & "    fraLista.Left = 10" & vbCrLf
    s = s & "    fraLista.Top = 56" & vbCrLf
    s = s & "    fraLista.Width = DW - 20" & vbCrLf
    s = s & "    fraLista.Height = 292" & vbCrLf
    s = s & "    fraLista.Caption = """"" & vbCrLf
    s = s & "    fraLista.SpecialEffect = 0" & vbCrLf
    s = s & "    fraLista.BorderStyle = 1" & vbCrLf
    s = s & "    fraLista.BorderColor = MC_BORDA" & vbCrLf
    s = s & "    fraLista.BackColor = MC_BRANCO" & vbCrLf
    s = s & "    fraLista.ScrollBars = 2" & vbCrLf
    s = s & "    fraLista.KeepScrollBarsVisible = 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Legenda ""legA"", ""atrasada"", 16, MC_CORAL" & vbCrLf
    s = s & "    Legenda ""legH"", ""revisar hoje"", 86, MC_AMARELO" & vbCrLf
    s = s & "    Legenda ""legF"", ""futura ou sem data"", 170, MC_CINZA_CLARO" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Botao ""btnAbrir"", ""abrir"", ""Abrir obra"", 10, 376, 170, 26, MC_NAVY, MC_BRANCO, -1, f, t" & vbCrLf
    s = s & "    Botao ""btnConcluir"", ""concluir"", ""Concluir an" & ChrW(225) & "lise"", 188, 376, 150, 26, MC_BRANCO, MC_CORAL, MC_CORAL, f, t" & vbCrLf
    s = s & "    Botao ""btnMinhas"", ""minhas"", ""S" & ChrW(243) & " as minhas"", 346, 376, 130, 26, MC_BRANCO, MC_NAVY, MC_BORDA, lblMinhasF, lblMinhasT" & vbCrLf
    s = s & "    Botao ""btnFechar"", ""fechar"", ""Fechar"", DW - 110, 376, 100, 26, MC_BRANCO, MC_NAVY, MC_BORDA, f, t" & vbCrLf
    s = s & "    Set lblStatus = Rot(Me, ""lblStatus"", """", 12, 406, DW - 24, 14, MC_CINZA, 8, False)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub PintarMinhas()" & vbCrLf
    s = s & "    If mSoMinhas Then" & vbCrLf
    CodMarcLista4 = s
End Function

Private Function CodMarcLista5() As String
    Dim s As String
    s = s & "        lblMinhasF.BackColor = MC_AZUL" & vbCrLf
    s = s & "        lblMinhasF.BorderColor = MC_AZUL" & vbCrLf
    s = s & "        lblMinhasT.ForeColor = MC_BRANCO" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblMinhasF.BackColor = MC_BRANCO" & vbCrLf
    s = s & "        lblMinhasF.BorderColor = MC_BORDA" & vbCrLf
    s = s & "        lblMinhasT.ForeColor = MC_NAVY" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' LISTA" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "Private Sub Carregar()" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim y As Single" & vbCrLf
    s = s & "    Dim fundo As Long" & vbCrLf
    s = s & "    Dim c As MSForms.Label" & vbCrLf
    s = s & "    Dim ic As cMarcaIcone" & vbCrLf
    s = s & "    Dim chave As String" & vbCrLf
    s = s & "    Dim por As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Desligar mCliquesLinhas" & vbCrLf
    s = s & "    Set mCliquesLinhas = New Collection" & vbCrLf
    s = s & "    fraLista.Controls.Clear" & vbCrLf
    s = s & "    mSel = 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mDados = mMarcas.Lista(mSoMinhas)" & vbCrLf
    s = s & "    If IsArray(mDados) Then" & vbCrLf
    s = s & "        mQtd = UBound(mDados, 1) + 1" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        mQtd = 0" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    lblTitulo.Caption = ""OBRAS MARCADAS ("" & mQtd & "")""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mQtd = 0 Then" & vbCrLf
    s = s & "        Rot fraLista, ""lblVazio"", IIf(mSoMinhas, ""Voc" & ChrW(234) & " n" & ChrW(227) & "o tem obras marcadas."", ""Nenhuma obra marcada para analisar.""), 12, 14, 600, 14, MC_CINZA, 9, False" & vbCrLf
    s = s & "        fraLista.ScrollHeight = 0" & vbCrLf
    s = s & "        Status """", 0" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ReDim mFundos(1 To mQtd)" & vbCrLf
    s = s & "    For i = 1 To mQtd" & vbCrLf
    s = s & "        y = (i - 1) * ALT_LINHA" & vbCrLf
    CodMarcLista5 = s
End Function

Private Function CodMarcLista6() As String
    Dim s As String
    s = s & "        chave = ""linha:"" & i" & vbCrLf
    s = s & "        fundo = FundoDaLinha(i)" & vbCrLf
    s = s & "        Set mFundos(i) = Rot(fraLista, ""lnF"" & i, """", 0, y, DW - 40, ALT_LINHA, MC_NAVY, 8, False, fundo)" & vbCrLf
    s = s & "        Clique mCliquesLinhas, mFundos(i), chave" & vbCrLf
    s = s & "        Set c = Rot(fraLista, ""lnS"" & i, """", 0, y + ALT_LINHA - 1, DW - 40, 1, MC_NAVY, 8, False, MC_BORDA)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        Set ic = New cMarcaIcone" & vbCrLf
    s = s & "        ic.Criar fraLista, ""lnI"" & i, 10, y + 10, 0.9, CorPrazo(mDados(i - 1, 3))" & vbCrLf
    s = s & "        Clique mCliquesLinhas, ic.Corpo, chave" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        Set c = Rot(fraLista, ""lnA"" & i, mDados(i - 1, 0), 30, y + 5, 190, 12, MC_NAVY, 9, True)" & vbCrLf
    s = s & "        Clique mCliquesLinhas, c, chave" & vbCrLf
    s = s & "        Set c = Rot(fraLista, ""lnL"" & i, mDados(i - 1, 1), 30, y + 18, 190, 13, MC_NAVY, 8.5, False)" & vbCrLf
    s = s & "        Clique mCliquesLinhas, c, chave" & vbCrLf
    s = s & "        Set c = Rot(fraLista, ""lnT"" & i, mDados(i - 1, 2), 222, y + 4, 264, 28, MC_NAVY, 8.5, False)" & vbCrLf
    s = s & "        c.WordWrap = True" & vbCrLf
    s = s & "        c.ControlTipText = Left$(mDados(i - 1, 2), 250)" & vbCrLf
    s = s & "        Clique mCliquesLinhas, c, chave" & vbCrLf
    s = s & "        Set c = Rot(fraLista, ""lnR"" & i, TextoPrazo(mDados(i - 1, 3)), 490, y + 11, 98, 13, MC_NAVY, 8.5, False)" & vbCrLf
    s = s & "        Select Case TipoPrazo(mDados(i - 1, 3))" & vbCrLf
    s = s & "            Case PZ_ATRASADA" & vbCrLf
    s = s & "                c.ForeColor = MC_CORAL" & vbCrLf
    s = s & "                c.Font.Bold = True" & vbCrLf
    s = s & "            Case PZ_HOJE" & vbCrLf
    s = s & "                c.Font.Bold = True" & vbCrLf
    s = s & "        End Select" & vbCrLf
    s = s & "        Clique mCliquesLinhas, c, chave" & vbCrLf
    s = s & "        por = mDados(i - 1, 4)" & vbCrLf
    s = s & "        If mDados(i - 1, 6) Then por = por & "" (s" & ChrW(243) & " voc" & ChrW(234) & ")""" & vbCrLf
    s = s & "        Set c = Rot(fraLista, ""lnP"" & i, por, 590, y + 11, 90, 13, MC_NAVY, 8.5, False)" & vbCrLf
    s = s & "        If IsDate(mDados(i - 1, 5)) Then c.ControlTipText = ""Marcada em "" & Format$(mDados(i - 1, 5), ""dd/mm/yyyy hh:mm"")" & vbCrLf
    s = s & "        Clique mCliquesLinhas, c, chave" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    fraLista.ScrollHeight = mQtd * ALT_LINHA" & vbCrLf
    s = s & "    fraLista.ScrollTop = 0" & vbCrLf
    s = s & "    Status """", 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FundoDaLinha(ByVal i As Long) As Long" & vbCrLf
    s = s & "    If i = mSel Then" & vbCrLf
    s = s & "        FundoDaLinha = MC_SELECAO" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Select Case TipoPrazo(mDados(i - 1, 3))" & vbCrLf
    s = s & "        Case PZ_ATRASADA" & vbCrLf
    CodMarcLista6 = s
End Function

Private Function CodMarcLista7() As String
    Dim s As String
    s = s & "            FundoDaLinha = MC_CORAL_FUNDO" & vbCrLf
    s = s & "        Case PZ_HOJE" & vbCrLf
    s = s & "            FundoDaLinha = MC_DESTAQUE" & vbCrLf
    s = s & "        Case Else" & vbCrLf
    s = s & "            FundoDaLinha = MC_BRANCO" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Selecionar(ByVal i As Long)" & vbCrLf
    s = s & "    Dim antes As Long" & vbCrLf
    s = s & "    If i < 1 Or i > mQtd Then Exit Sub" & vbCrLf
    s = s & "    antes = mSel" & vbCrLf
    s = s & "    mSel = i" & vbCrLf
    s = s & "    If antes >= 1 And antes <= mQtd Then mFundos(antes).BackColor = FundoDaLinha(antes)" & vbCrLf
    s = s & "    mFundos(i).BackColor = FundoDaLinha(i)" & vbCrLf
    s = s & "    Status mDados(i - 1, 0) & ""  -  "" & mDados(i - 1, 1), 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Status(ByVal texto As String, ByVal tipo As Long)" & vbCrLf
    s = s & "    If tipo = 2 Then" & vbCrLf
    s = s & "        lblStatus.ForeColor = MC_CORAL" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblStatus.ForeColor = MC_CINZA" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    lblStatus.Caption = texto" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' CLIQUES" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "Public Sub TratarClique(ByVal chave As String)" & vbCrLf
    s = s & "    If Left$(chave, 6) = ""linha:"" Then" & vbCrLf
    s = s & "        Selecionar CLng(Mid$(chave, 7))" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Select Case chave" & vbCrLf
    s = s & "        Case ""fechar""" & vbCrLf
    s = s & "            Me.Hide" & vbCrLf
    s = s & "        Case ""abrir""" & vbCrLf
    s = s & "            AbrirSelecionada" & vbCrLf
    s = s & "        Case ""minhas""" & vbCrLf
    s = s & "            mSoMinhas = Not mSoMinhas" & vbCrLf
    s = s & "            PintarMinhas" & vbCrLf
    s = s & "            Carregar" & vbCrLf
    s = s & "        Case ""concluir""" & vbCrLf
    CodMarcLista7 = s
End Function

Private Function CodMarcLista8() As String
    Dim s As String
    s = s & "            If mSel = 0 Then" & vbCrLf
    s = s & "                Status ""Selecione uma obra na lista."", 2" & vbCrLf
    s = s & "                Exit Sub" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "            If mMarcas.Concluir(CStr(mDados(mSel - 1, 0))) Then" & vbCrLf
    s = s & "                Carregar" & vbCrLf
    s = s & "                Status ""An" & ChrW(225) & "lise conclu" & ChrW(237) & "da; a obra saiu da lista."", 0" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                Status mMarcas.UltimoErro, 2" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub TratarDuploClique(ByVal chave As String)" & vbCrLf
    s = s & "    If Left$(chave, 6) = ""linha:"" Then" & vbCrLf
    s = s & "        Selecionar CLng(Mid$(chave, 7))" & vbCrLf
    s = s & "        AbrirSelecionada" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub AbrirSelecionada()" & vbCrLf
    s = s & "    If mSel = 0 Then" & vbCrLf
    s = s & "        Status ""Selecione uma obra na lista."", 2" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If mMarcas.LinhaNoFollowUp(CStr(mDados(mSel - 1, 0))) = 0 Then" & vbCrLf
    s = s & "        Status ""Esta obra n" & ChrW(227) & "o est" & ChrW(225) & " mais no follow up (foi entregue ou declinada). Conclua a an" & ChrW(225) & "lise se n" & ChrW(227) & "o precisar mais dela."", 2" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    mEscolhida = CStr(mDados(mSel - 1, 0))" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodMarcLista8 = s
End Function

Private Function CodigoNovaModulo() As String
    Dim s As String
    s = s & CodNovaModulo1()
    s = s & CodNovaModulo2()
    s = s & CodNovaModulo3()
    s = s & CodNovaModulo4()
    CodigoNovaModulo = s
End Function

Private Function CodNovaModulo1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' mNovaObra - apoio da " & ChrW(225) & "rea NOVA OBRA do Painel de Obras (cNovaObraPainel):" & vbCrLf
    s = s & "' inclus" & ChrW(227) & "o de obra nova no follow up pelo C" & ChrW(211) & "DIGO, com confer" & ChrW(234) & "ncia" & vbCrLf
    s = s & "' nas bases de onde as f" & ChrW(243) & "rmulas da tb_Obras trazem os dados:" & vbCrLf
    s = s & "'   1" & ChrW(186) & " tb_Cronograma (Cronograma de Obras)" & vbCrLf
    s = s & "'   2" & ChrW(186) & " tb_Carteira (aba CARTEIRA), quando n" & ChrW(227) & "o est" & ChrW(225) & " no cronograma" & vbCrLf
    s = s & "' O C" & ChrW(211) & "DIGO " & ChrW(233) & " gravado como texto (igual " & ChrW(224) & "s bases); o resto vem das f" & ChrW(243) & "rmulas." & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Const NL_NAVY As Long = 3877150" & vbCrLf
    s = s & "Public Const NL_AZUL As Long = 8950797" & vbCrLf
    s = s & "Public Const NL_CORAL As Long = 2500316" & vbCrLf
    s = s & "Public Const NL_AMARELO As Long = 2408443" & vbCrLf
    s = s & "Public Const NL_CINZA As Long = 9139300" & vbCrLf
    s = s & "Public Const NL_FUNDO As Long = 16579320" & vbCrLf
    s = s & "Public Const NL_BORDA As Long = 15788258" & vbCrLf
    s = s & "Public Const NL_BRANCO As Long = 16777215" & vbCrLf
    s = s & "Public Const NL_VERDE As Long = 4891414" & vbCrLf
    s = s & "Public Const NL_DESTAQUE As Long = 15465471" & vbCrLf
    s = s & "Public Const NL_VERDE_FUNDO As Long = 15595240" & vbCrLf
    s = s & "Public Const NL_CORAL_FUNDO As Long = 14869246" & vbCrLf
    s = s & "Public Const NL_AZUL_FUNDO As Long = 16449008" & vbCrLf
    s = s & "Public Const NL_EDICAO As Long = 15465471" & vbCrLf
    s = s & "Public Const NL_DESLIGADO As Long = 12100500" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mFonteTexto As String" & vbCrLf
    s = s & "Private mFonteTitulo As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function TabelaNL(ByVal nome As String) As ListObject" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Dim lo As ListObject" & vbCrLf
    s = s & "    For Each ws In ThisWorkbook.Worksheets" & vbCrLf
    s = s & "        Set lo = Nothing" & vbCrLf
    s = s & "        On Error Resume Next" & vbCrLf
    s = s & "        Set lo = ws.ListObjects(nome)" & vbCrLf
    s = s & "        On Error GoTo 0" & vbCrLf
    s = s & "        If Not lo Is Nothing Then" & vbCrLf
    s = s & "            Set TabelaNL = lo" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next ws" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function NormNL(ByVal s As String) As String" & vbCrLf
    CodNovaModulo1 = s
End Function

Private Function CodNovaModulo2() As String
    Dim s As String
    s = s & "    s = Replace(Replace(Replace(s, vbCr, "" ""), vbLf, "" ""), Chr$(160), "" "")" & vbCrLf
    s = s & "    Do While InStr(1, s, ""  "", vbBinaryCompare) > 0" & vbCrLf
    s = s & "        s = Replace(s, ""  "", "" "")" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "    NormNL = UCase$(Trim$(s))" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function ColNL(ByVal tb As ListObject, ByVal nome As String) As Long" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Function" & vbCrLf
    s = s & "    For i = 1 To tb.ListColumns.Count" & vbCrLf
    s = s & "        If NormNL(tb.ListColumns(i).Name) = NormNL(nome) Then" & vbCrLf
    s = s & "            ColNL = i" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function TextoNL(ByVal v As Variant) As String" & vbCrLf
    s = s & "    If IsError(v) Then Exit Function" & vbCrLf
    s = s & "    If IsEmpty(v) Then Exit Function" & vbCrLf
    s = s & "    If VarType(v) = vbDate Then" & vbCrLf
    s = s & "        TextoNL = Format$(v, ""dd\/mm\/yyyy"")" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        TextoNL = Trim$(Replace(CStr(v), Chr$(160), "" ""))" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Linha da obra (C" & ChrW(211) & "DIGO comparado como texto) na tabela; 0 = n" & ChrW(227) & "o est" & ChrW(225) & "." & vbCrLf
    s = s & "Public Function LinhaNaTabela(ByVal nome As String, ByVal codObra As String) As Long" & vbCrLf
    s = s & "    Dim tb As ListObject" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim v As Variant" & vbCrLf
    s = s & "    Set tb = TabelaNL(nome)" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Function" & vbCrLf
    s = s & "    If tb.DataBodyRange Is Nothing Then Exit Function" & vbCrLf
    s = s & "    c = ColNL(tb, ""C" & ChrW(211) & "DIGO"")" & vbCrLf
    s = s & "    If c = 0 Then Exit Function" & vbCrLf
    s = s & "    codObra = Trim$(codObra)" & vbCrLf
    s = s & "    If Len(codObra) = 0 Then Exit Function" & vbCrLf
    s = s & "    v = tb.ListColumns(c).DataBodyRange.Value" & vbCrLf
    s = s & "    If Not IsArray(v) Then" & vbCrLf
    s = s & "        If TextoNL(v) = codObra Then LinhaNaTabela = 1" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    CodNovaModulo2 = s
End Function

Private Function CodNovaModulo3() As String
    Dim s As String
    s = s & "    End If" & vbCrLf
    s = s & "    For i = 1 To UBound(v, 1)" & vbCrLf
    s = s & "        If TextoNL(v(i, 1)) = codObra Then" & vbCrLf
    s = s & "            LinhaNaTabela = i" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Valor da coluna na linha i da tabela (texto; """" se n" & ChrW(227) & "o houver)." & vbCrLf
    s = s & "Public Function ValorNL(ByVal nome As String, ByVal i As Long, ByVal coluna As String) As String" & vbCrLf
    s = s & "    Dim tb As ListObject" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    If i = 0 Then Exit Function" & vbCrLf
    s = s & "    Set tb = TabelaNL(nome)" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Function" & vbCrLf
    s = s & "    c = ColNL(tb, coluna)" & vbCrLf
    s = s & "    If c = 0 Then Exit Function" & vbCrLf
    s = s & "    ValorNL = TextoNL(tb.DataBodyRange.Cells(i, c).Value)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Fontes da identidade visual (Segoe UI de reserva)" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "Public Function FonteNL() As String" & vbCrLf
    s = s & "    If Len(mFonteTexto) = 0 Then DetectarFontes" & vbCrLf
    s = s & "    FonteNL = mFonteTexto" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function FonteTituloNL() As String" & vbCrLf
    s = s & "    If Len(mFonteTitulo) = 0 Then DetectarFontes" & vbCrLf
    s = s & "    FonteTituloNL = mFonteTitulo" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub DetectarFontes()" & vbCrLf
    s = s & "    Dim achado As String" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    achado = Dir$(Environ$(""WINDIR"") & ""\Fonts\Segoe UI*"")" & vbCrLf
    s = s & "    If Len(achado) = 0 Then achado = Dir$(Environ$(""LOCALAPPDATA"") & ""\Microsoft\Windows\Fonts\Segoe UI*"")" & vbCrLf
    s = s & "    If Len(achado) > 0 Then" & vbCrLf
    s = s & "        mFonteTexto = ""Segoe UI""" & vbCrLf
    s = s & "        mFonteTitulo = ""Segoe UI Black""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        mFonteTexto = ""Segoe UI""" & vbCrLf
    s = s & "        mFonteTitulo = ""Segoe UI Black""" & vbCrLf
    CodNovaModulo3 = s
End Function

Private Function CodNovaModulo4() As String
    Dim s As String
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodNovaModulo4 = s
End Function

Private Function CodigoNovaClique() As String
    Dim s As String
    s = s & CodNovaClique1()
    CodigoNovaClique = s
End Function

Private Function CodNovaClique1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' cNovaClique - repassa " & ChrW(224) & " janela Nova obra o clique de um r" & ChrW(243) & "tulo criado por" & vbCrLf
    s = s & "' c" & ChrW(243) & "digo (bot" & ChrW(245) & "es, c" & ChrW(237) & "rculos do farol): dono.TratarClique chave" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private WithEvents mRotulo As MSForms.Label" & vbCrLf
    s = s & "Private mDono As Object" & vbCrLf
    s = s & "Private mChave As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Conectar(ByVal rotulo As MSForms.Label, ByVal chave As String, ByVal dono As Object)" & vbCrLf
    s = s & "    Set mRotulo = rotulo" & vbCrLf
    s = s & "    mChave = chave" & vbCrLf
    s = s & "    Set mDono = dono" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Desconectar()" & vbCrLf
    s = s & "    Set mRotulo = Nothing" & vbCrLf
    s = s & "    Set mDono = Nothing" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub mRotulo_Click()" & vbCrLf
    s = s & "    If Not mDono Is Nothing Then mDono.TratarClique mChave" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodNovaClique1 = s
End Function

Private Function CodigoNovaClasse() As String
    Dim s As String
    s = s & CodNovaClasse1()
    s = s & CodNovaClasse2()
    s = s & CodNovaClasse3()
    s = s & CodNovaClasse4()
    s = s & CodNovaClasse5()
    s = s & CodNovaClasse6()
    s = s & CodNovaClasse7()
    s = s & CodNovaClasse8()
    s = s & CodNovaClasse9()
    s = s & CodNovaClasse10()
    s = s & CodNovaClasse11()
    CodigoNovaClasse = s
End Function

Private Function CodNovaClasse1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' cNovaObraPainel - " & ChrW(225) & "rea NOVA OBRA dentro do Painel de Obras" & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "' Ocupa a " & ChrW(225) & "rea de dados do painel enquanto a obra nova " & ChrW(233) & " conferida. O C" & ChrW(211) & "DIGO" & vbCrLf
    s = s & "' " & ChrW(233) & " a chave: a " & ChrW(225) & "rea procura a obra no Cronograma de Obras e, se n" & ChrW(227) & "o" & vbCrLf
    s = s & "' achar, na Carteira - as mesmas bases das f" & ChrW(243) & "rmulas da" & vbCrLf
    s = s & "' tb_Obras - e mostra o que as f" & ChrW(243) & "rmulas v" & ChrW(227) & "o trazer antes de incluir." & vbCrLf
    s = s & "' Ao terminar, avisa o painel pelo evento Fechar." & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Event Fechar(ByVal linhaObras As Long)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const T_Obras As String = ""tb_Obras""" & vbCrLf
    s = s & "Private Const T_CRONO As String = ""tb_Cronograma""" & vbCrLf
    s = s & "Private Const T_AUX As String = ""tb_Carteira""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const S_VAZIO As Long = 0" & vbCrLf
    s = s & "Private Const S_CRONO As Long = 1" & vbCrLf
    s = s & "Private Const S_AUX As Long = 2" & vbCrLf
    s = s & "Private Const S_NADA As Long = 3" & vbCrLf
    s = s & "Private Const S_NA_Obras As Long = 4" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private WithEvents txtCod As MSForms.TextBox" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mFra As MSForms.Frame" & vbCrLf
    s = s & "Private mCliques As Collection" & vbCrLf
    s = s & "Private mFonte As String" & vbCrLf
    s = s & "Private mFonteTit As String" & vbCrLf
    s = s & "Private mSituacao As Long" & vbCrLf
    s = s & "Private mCod As String" & vbCrLf
    s = s & "Private mLinhaObras As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private lblIcone As MSForms.Label" & vbCrLf
    s = s & "Private lblTit As MSForms.Label" & vbCrLf
    s = s & "Private lblSub As MSForms.Label" & vbCrLf
    s = s & "Private lblCaixa As MSForms.Label" & vbCrLf
    s = s & "Private lblDesc As MSForms.Label" & vbCrLf
    s = s & "Private mRot(1 To 8) As MSForms.Label" & vbCrLf
    s = s & "Private mVal(1 To 8) As MSForms.Label" & vbCrLf
    s = s & "Private lblPrevTit As MSForms.Label" & vbCrLf
    s = s & "Private lblOkF As MSForms.Label" & vbCrLf
    s = s & "Private lblOkT As MSForms.Label" & vbCrLf
    s = s & "Private lblStatus As MSForms.Label" & vbCrLf
    s = s & "" & vbCrLf
    CodNovaClasse1 = s
End Function

Private Function CodNovaClasse2() As String
    Dim s As String
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' MONTAGEM (dentro do Frame que o painel cria)" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "Private Function Rot(ByVal nome As String, ByVal texto As String, _" & vbCrLf
    s = s & "                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _" & vbCrLf
    s = s & "                     ByVal cor As Long, ByVal tam As Single, ByVal negrito As Boolean, _" & vbCrLf
    s = s & "                     Optional ByVal fundo As Long = -1, Optional ByVal alinh As Long = 1) As MSForms.Label" & vbCrLf
    s = s & "    Dim c As MSForms.Label" & vbCrLf
    s = s & "    Set c = mFra.Controls.Add(""Forms.Label.1"", nome, True)" & vbCrLf
    s = s & "    c.Left = l" & vbCrLf
    s = s & "    c.Top = t" & vbCrLf
    s = s & "    c.Width = w" & vbCrLf
    s = s & "    c.Height = h" & vbCrLf
    s = s & "    c.Font.Name = mFonte" & vbCrLf
    s = s & "    c.Font.Size = tam" & vbCrLf
    s = s & "    c.Font.Bold = negrito" & vbCrLf
    s = s & "    c.ForeColor = cor" & vbCrLf
    s = s & "    If fundo >= 0 Then" & vbCrLf
    s = s & "        c.BackStyle = 1" & vbCrLf
    s = s & "        c.BackColor = fundo" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        c.BackStyle = 0" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    c.TextAlign = alinh" & vbCrLf
    s = s & "    c.WordWrap = False" & vbCrLf
    s = s & "    c.Caption = texto" & vbCrLf
    s = s & "    Set Rot = c" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Clique(ByVal rotulo As MSForms.Label, ByVal chave As String)" & vbCrLf
    s = s & "    Dim ev As cNovaClique" & vbCrLf
    s = s & "    Set ev = New cNovaClique" & vbCrLf
    s = s & "    ev.Conectar rotulo, chave, Me" & vbCrLf
    s = s & "    mCliques.Add ev" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Botao(ByVal nome As String, ByVal chave As String, ByVal texto As String, _" & vbCrLf
    s = s & "                  ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _" & vbCrLf
    s = s & "                  ByVal fundo As Long, ByVal cor As Long, ByVal borda As Long, _" & vbCrLf
    s = s & "                  ByRef rotFundo As MSForms.Label, ByRef rotTexto As MSForms.Label)" & vbCrLf
    s = s & "    Set rotFundo = Rot(nome & ""F"", """", l, t, w, h, cor, 8, False, fundo)" & vbCrLf
    s = s & "    If borda >= 0 Then" & vbCrLf
    s = s & "        rotFundo.BorderStyle = 1" & vbCrLf
    s = s & "        rotFundo.BorderColor = borda" & vbCrLf
    s = s & "    End If" & vbCrLf
    CodNovaClasse2 = s
End Function

Private Function CodNovaClasse3() As String
    Dim s As String
    s = s & "    Set rotTexto = Rot(nome & ""T"", texto, l, t + (h - 12) / 2 - 0.5, w, 12, cor, 8.5, True, -1, 2)" & vbCrLf
    s = s & "    Clique rotFundo, chave" & vbCrLf
    s = s & "    Clique rotTexto, chave" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Criar(ByVal fra As MSForms.Frame)" & vbCrLf
    s = s & "    Dim c As MSForms.Label" & vbCrLf
    s = s & "    Dim f As MSForms.Label" & vbCrLf
    s = s & "    Dim t As MSForms.Label" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim nomes As Variant" & vbCrLf
    s = s & "    Dim w As Single" & vbCrLf
    s = s & "    Dim xd As Single" & vbCrLf
    s = s & "    Dim wd As Single" & vbCrLf
    s = s & "    Dim y As Single" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set mFra = fra" & vbCrLf
    s = s & "    Set mCliques = New Collection" & vbCrLf
    s = s & "    mFonte = FonteNL()" & vbCrLf
    s = s & "    mFonteTit = FonteTituloNL()" & vbCrLf
    s = s & "    w = mFra.InsideWidth" & vbCrLf
    s = s & "    If w < 600 Then w = mFra.Width - 4" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Rot ""nlFaixa"", """", 0, 0, w, 28, NL_BRANCO, 8, False, NL_NAVY" & vbCrLf
    s = s & "    Rot ""nlFaixaAm"", """", 0, 28, w, 3, NL_BRANCO, 8, False, NL_AMARELO" & vbCrLf
    s = s & "    Set c = Rot(""nlTitulo"", ""NOVA OBRA NO FOLLOW UP"", 14, 6, 300, 18, NL_BRANCO, 11.5, False)" & vbCrLf
    s = s & "    c.Font.Name = mFonteTit" & vbCrLf
    s = s & "    Set c = Rot(""nlFechar"", ""Voltar ao painel  "" & ChrW(215), w - 150, 8, 136, 14, NL_BRANCO, 9, True, -1, 3)" & vbCrLf
    s = s & "    Clique c, ""cancelar""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set c = Rot(""nlIntro"", ""O C" & ChrW(211) & "DIGO " & ChrW(233) & " a chave: o resto da linha vem das f" & ChrW(243) & "rmulas, que buscam a obra no Cronograma de Obras e, se n" & ChrW(227) & "o acharem, na Carteira. Confira antes de incluir."", 16, 40, w - 32, 14, NL_CINZA, 8.5, False)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' coluna da esquerda: C" & ChrW(211) & "DIGO e situa" & ChrW(231) & "" & ChrW(227) & "o" & vbCrLf
    s = s & "    Rot ""nlCodRot"", ""C" & ChrW(211) & "DIGO"", 16, 66, 100, 12, NL_CINZA, 7.5, True" & vbCrLf
    s = s & "    Set txtCod = mFra.Controls.Add(""Forms.TextBox.1"", ""nlCod"", True)" & vbCrLf
    s = s & "    txtCod.Left = 16" & vbCrLf
    s = s & "    txtCod.Top = 80" & vbCrLf
    s = s & "    txtCod.Width = 160" & vbCrLf
    s = s & "    txtCod.Height = 28" & vbCrLf
    s = s & "    txtCod.MaxLength = 12" & vbCrLf
    s = s & "    txtCod.SpecialEffect = 0" & vbCrLf
    s = s & "    txtCod.BorderStyle = 1" & vbCrLf
    s = s & "    txtCod.BorderColor = NL_AMARELO" & vbCrLf
    s = s & "    txtCod.BackColor = NL_EDICAO" & vbCrLf
    s = s & "    txtCod.ForeColor = NL_NAVY" & vbCrLf
    CodNovaClasse3 = s
End Function

Private Function CodNovaClasse4() As String
    Dim s As String
    s = s & "    txtCod.Font.Name = mFonte" & vbCrLf
    s = s & "    txtCod.Font.Size = 15" & vbCrLf
    s = s & "    txtCod.Font.Bold = True" & vbCrLf
    s = s & "    Rot ""nlCodDica"", ""a confer" & ChrW(234) & "ncia " & ChrW(233) & " feita enquanto voc" & ChrW(234) & " digita"", 184, 90, 200, 12, NL_CINZA, 8, False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set lblCaixa = Rot(""nlCaixa"", """", 16, 120, 360, 170, NL_NAVY, 8, False, NL_BRANCO)" & vbCrLf
    s = s & "    lblCaixa.BorderStyle = 1" & vbCrLf
    s = s & "    lblCaixa.BorderColor = NL_BORDA" & vbCrLf
    s = s & "    Set lblIcone = Rot(""nlIcone"", """", 26, 130, 18, 18, NL_VERDE, 13, False)" & vbCrLf
    s = s & "    lblIcone.Font.Name = ""Wingdings""" & vbCrLf
    s = s & "    lblIcone.Font.Charset = 2" & vbCrLf
    s = s & "    Set lblTit = Rot(""nlTit"", """", 46, 130, 320, 17, NL_NAVY, 11, True)" & vbCrLf
    s = s & "    Set lblSub = Rot(""nlSub"", """", 46, 152, 320, 130, NL_CINZA, 8.5, False)" & vbCrLf
    s = s & "    lblSub.WordWrap = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Botao ""nlOk"", ""ok"", ""Incluir no follow up"", 16, 300, 360, 30, NL_NAVY, NL_BRANCO, -1, lblOkF, lblOkT" & vbCrLf
    s = s & "    Botao ""nlCancelar"", ""cancelar"", ""Cancelar e voltar ao painel"", 16, 338, 360, 26, NL_BRANCO, NL_NAVY, NL_BORDA, f, t" & vbCrLf
    s = s & "    Set lblStatus = Rot(""nlStatus"", """", 16, 372, 360, 13, NL_CINZA, 8, False)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' coluna da direita: o que as f" & ChrW(243) & "rmulas v" & ChrW(227) & "o trazer" & vbCrLf
    s = s & "    xd = 396" & vbCrLf
    s = s & "    wd = w - xd - 16" & vbCrLf
    s = s & "    Rot ""nlSep"", """", xd - 10, 66, 1, 300, NL_NAVY, 6, False, NL_BORDA" & vbCrLf
    s = s & "    Set lblPrevTit = Rot(""nlPrevTit"", ""O QUE AS F" & ChrW(211) & "RMULAS DO FOLLOW UP V" & ChrW(195) & "O TRAZER"", xd, 66, wd, 12, NL_CINZA, 7.5, True)" & vbCrLf
    s = s & "    Rot ""nlDescRot"", ""NOME"", xd, 84, 120, 12, NL_CINZA, 7.5, True" & vbCrLf
    s = s & "    Set lblDesc = Rot(""nlDesc"", """", xd + 124, 81, wd - 124, 20, NL_NAVY, 10, True, NL_FUNDO)" & vbCrLf
    s = s & "    lblDesc.BorderStyle = 1" & vbCrLf
    s = s & "    lblDesc.BorderColor = NL_BORDA" & vbCrLf
    s = s & "    nomes = Array(""CLIENTE"", ""CRITICIDADE"", ""CIDADE - UF"", ""TIPO DE OBRA"", """ & ChrW(193) & "REA (m" & ChrW(178) & ")"", ""IN" & ChrW(205) & "CIO DA OBRA"", ""FIM DA OBRA"", ""ENTREGA"")" & vbCrLf
    s = s & "    For i = 1 To 8" & vbCrLf
    s = s & "        y = 110 + (i - 1) * 28" & vbCrLf
    s = s & "        Set mRot(i) = Rot(""nlPR"" & i, CStr(nomes(i - 1)), xd, y + 3, 120, 12, NL_CINZA, 7.5, True)" & vbCrLf
    s = s & "        Set mVal(i) = Rot(""nlPV"" & i, """", xd + 124, y, wd - 124, 20, NL_NAVY, 10, False, NL_FUNDO)" & vbCrLf
    s = s & "        mVal(i).BorderStyle = 1" & vbCrLf
    s = s & "        mVal(i).BorderColor = NL_BORDA" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    Rot ""nlNota"", ""Somente o C" & ChrW(211) & "DIGO " & ChrW(233) & " gravado no follow up (como texto, igual " & ChrW(224) & "s bases). Estes dados n" & ChrW(227) & "o s" & ChrW(227) & "o digitados: aparecem pelas f" & ChrW(243) & "rmulas da tabela."", xd, 340, wd, 26, NL_CINZA, 8, False" & vbCrLf
    s = s & "    Me.Limpar" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Limpar()" & vbCrLf
    s = s & "    If txtCod Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    txtCod.Text = """"" & vbCrLf
    s = s & "    Conferir" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodNovaClasse4 = s
End Function

Private Function CodNovaClasse5() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "Public Sub Focar()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    txtCod.SetFocus" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Desligar()" & vbCrLf
    s = s & "    Dim ev As cNovaClique" & vbCrLf
    s = s & "    If mCliques Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    For Each ev In mCliques" & vbCrLf
    s = s & "        ev.Desconectar" & vbCrLf
    s = s & "    Next ev" & vbCrLf
    s = s & "    Set mCliques = Nothing" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' CONFER" & ChrW(202) & "NCIA" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "Private Sub txtCod_Change()" & vbCrLf
    s = s & "    Conferir" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub txtCod_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" & vbCrLf
    s = s & "    If KeyCode = 13 Then" & vbCrLf
    s = s & "        KeyCode = 0" & vbCrLf
    s = s & "        TratarClique ""ok""" & vbCrLf
    s = s & "    ElseIf KeyCode = 27 Then" & vbCrLf
    s = s & "        KeyCode = 0" & vbCrLf
    s = s & "        TratarClique ""cancelar""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function Valor(ByVal iCrono As Long, ByVal iAux As Long, ByVal colCrono As String, ByVal colAux As String) As String" & vbCrLf
    s = s & "    If iCrono > 0 And Len(colCrono) > 0 Then Valor = ValorNL(T_CRONO, iCrono, colCrono)" & vbCrLf
    s = s & "    If Len(Valor) = 0 And iAux > 0 And Len(colAux) > 0 Then Valor = ValorNL(T_AUX, iAux, colAux)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Conferir()" & vbCrLf
    s = s & "    Dim iCrono As Long" & vbCrLf
    s = s & "    Dim iAux As Long" & vbCrLf
    s = s & "    Dim desc As String" & vbCrLf
    s = s & "    Dim cid As String" & vbCrLf
    s = s & "    Dim uf As String" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "" & vbCrLf
    CodNovaClasse5 = s
End Function

Private Function CodNovaClasse6() As String
    Dim s As String
    s = s & "    mCod = Trim$(Replace(txtCod.Text, Chr$(160), """"))" & vbCrLf
    s = s & "    mLinhaObras = 0" & vbCrLf
    s = s & "    lblStatus.Caption = """"" & vbCrLf
    s = s & "    For i = 1 To 8" & vbCrLf
    s = s & "        mVal(i).Caption = """"" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    lblDesc.Caption = """"" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Len(mCod) = 0 Then" & vbCrLf
    s = s & "        Pintar S_VAZIO, ""Digite o C" & ChrW(211) & "DIGO da obra"", ""Somente o C" & ChrW(211) & "DIGO " & ChrW(233) & " digitado; as demais colunas do follow up s" & ChrW(227) & "o preenchidas pelas f" & ChrW(243) & "rmulas.""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    iCrono = LinhaNaTabela(T_CRONO, mCod)" & vbCrLf
    s = s & "    iAux = LinhaNaTabela(T_AUX, mCod)" & vbCrLf
    s = s & "    mLinhaObras = LinhaNaTabela(T_Obras, mCod)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    desc = Valor(iCrono, iAux, ""NOME"", ""NOME"")" & vbCrLf
    s = s & "    If iCrono > 0 Or iAux > 0 Then" & vbCrLf
    s = s & "        mVal(1).Caption = "" "" & Valor(iCrono, iAux, ""CLIENTE"", ""CLIENTE"")" & vbCrLf
    s = s & "        mVal(2).Caption = "" "" & Valor(0, iAux, """", ""CRITICIDADE"")" & vbCrLf
    s = s & "        cid = Valor(iCrono, iAux, ""CIDADE"", ""CIDADE"")" & vbCrLf
    s = s & "        uf = Valor(iCrono, iAux, ""UF"", ""UF"")" & vbCrLf
    s = s & "        If Len(uf) > 0 Then cid = cid & "" - "" & uf" & vbCrLf
    s = s & "        mVal(3).Caption = "" "" & cid" & vbCrLf
    s = s & "        mVal(4).Caption = "" "" & Valor(iCrono, iAux, ""TIPO DE OBRA"", ""TIPO DE OBRA"")" & vbCrLf
    s = s & "        mVal(5).Caption = "" "" & Valor(0, iAux, """", """ & ChrW(193) & "REA (m" & ChrW(178) & ")"")" & vbCrLf
    s = s & "        mVal(6).Caption = "" "" & Valor(iCrono, iAux, ""IN" & ChrW(205) & "CIO DA OBRA"", ""IN" & ChrW(205) & "CIO PREVISTO"")" & vbCrLf
    s = s & "        mVal(7).Caption = "" "" & Valor(iCrono, iAux, ""FIM DA OBRA"", ""FIM PREVISTO"")" & vbCrLf
    s = s & "        mVal(8).Caption = "" "" & Valor(iCrono, iAux, ""ENTREGA PREVISTA"", ""ENTREGA PREVISTA"")" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    For i = 1 To 8" & vbCrLf
    s = s & "        If Len(Trim$(mVal(i).Caption)) = 0 Then mVal(i).Caption = "" "" & ChrW(8212)" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    If iCrono > 0 Or iAux > 0 Then lblDesc.Caption = "" "" & desc" & vbCrLf
    s = s & "    If Len(desc) = 0 Then desc = ""C" & ChrW(211) & "DIGO "" & mCod" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mLinhaObras > 0 Then" & vbCrLf
    s = s & "        Pintar S_NA_Obras, desc, ""Este C" & ChrW(211) & "DIGO j" & ChrW(225) & " est" & ChrW(225) & " no follow up. Nada ser" & ChrW(225) & " inclu" & ChrW(237) & "do para n" & ChrW(227) & "o duplicar a obra; abra a obra no painel.""" & vbCrLf
    s = s & "    ElseIf iCrono > 0 Then" & vbCrLf
    s = s & "        Pintar S_CRONO, desc, ""Cadastrada no Cronograma de Obras. As f" & ChrW(243) & "rmulas do follow up v" & ChrW(227) & "o trazer os dados ao lado.""" & vbCrLf
    s = s & "    ElseIf iAux > 0 Then" & vbCrLf
    s = s & "        Pintar S_AUX, desc, ""Est" & ChrW(225) & " na Carteira, mas ainda n" & ChrW(227) & "o no Cronograma de Obras: os dados ao lado v" & ChrW(234) & "m da Carteira e as datas podem mudar quando entrar no cronograma.""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Pintar S_NADA, ""C" & ChrW(211) & "DIGO "" & mCod & "" n" & ChrW(227) & "o encontrado"", ""N" & ChrW(227) & "o est" & ChrW(225) & " no Cronograma de Obras nem na Carteira. Confira o n" & ChrW(250) & "mero; se a obra " & ChrW(233) & " nova, cadastre-a na base e use ATUALIZAR DADOS no painel antes de incluir.""" & vbCrLf
    CodNovaClasse6 = s
End Function

Private Function CodNovaClasse7() As String
    Dim s As String
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Pintar(ByVal situacao As Long, ByVal titulo As String, ByVal detalhe As String)" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim mostrar As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mSituacao = situacao" & vbCrLf
    s = s & "    lblTit.Caption = titulo" & vbCrLf
    s = s & "    lblSub.Caption = detalhe" & vbCrLf
    s = s & "    mostrar = (situacao = S_CRONO Or situacao = S_AUX Or situacao = S_NA_Obras)" & vbCrLf
    s = s & "    For i = 1 To 8" & vbCrLf
    s = s & "        If Not mostrar Then mVal(i).Caption = """"" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    If Not mostrar Then lblDesc.Caption = """"" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Select Case situacao" & vbCrLf
    s = s & "        Case S_CRONO" & vbCrLf
    s = s & "            Icone Chr$(252), NL_VERDE, NL_VERDE_FUNDO, NL_VERDE" & vbCrLf
    s = s & "            Primario ""Incluir "" & Curto(titulo) & "" no follow up"", True" & vbCrLf
    s = s & "        Case S_AUX" & vbCrLf
    s = s & "            Icone Chr$(252), NL_AMARELO, NL_DESTAQUE, NL_AMARELO" & vbCrLf
    s = s & "            Primario ""Incluir "" & Curto(titulo) & "" no follow up"", True" & vbCrLf
    s = s & "        Case S_NA_Obras" & vbCrLf
    s = s & "            Icone Chr$(105), NL_AZUL, NL_AZUL_FUNDO, NL_AZUL, ""Webdings""" & vbCrLf
    s = s & "            Primario ""Abrir a obra no painel"", True" & vbCrLf
    s = s & "        Case S_NADA" & vbCrLf
    s = s & "            Icone Chr$(251), NL_CORAL, NL_CORAL_FUNDO, NL_CORAL" & vbCrLf
    s = s & "            Primario ""Incluir no follow up (bloqueado)"", False" & vbCrLf
    s = s & "        Case Else" & vbCrLf
    s = s & "            Icone """", NL_CINZA, NL_BRANCO, NL_BORDA" & vbCrLf
    s = s & "            Primario ""Incluir no follow up"", False" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function Curto(ByVal s As String) As String" & vbCrLf
    s = s & "    If Len(s) > 30 Then s = Left$(s, 28) & ""...""" & vbCrLf
    s = s & "    Curto = s" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Icone(ByVal glifo As String, ByVal cor As Long, ByVal fundo As Long, ByVal borda As Long, _" & vbCrLf
    s = s & "                  Optional ByVal fonte As String = ""Wingdings"")" & vbCrLf
    s = s & "    lblIcone.Font.Name = fonte" & vbCrLf
    s = s & "    lblIcone.Font.Charset = 2" & vbCrLf
    s = s & "    lblIcone.Caption = glifo" & vbCrLf
    CodNovaClasse7 = s
End Function

Private Function CodNovaClasse8() As String
    Dim s As String
    s = s & "    lblIcone.ForeColor = cor" & vbCrLf
    s = s & "    lblCaixa.BackColor = fundo" & vbCrLf
    s = s & "    lblCaixa.BorderColor = borda" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Primario(ByVal texto As String, ByVal ativo As Boolean)" & vbCrLf
    s = s & "    lblOkT.Caption = texto" & vbCrLf
    s = s & "    If ativo Then" & vbCrLf
    s = s & "        lblOkF.BackColor = NL_NAVY" & vbCrLf
    s = s & "        lblOkT.ForeColor = NL_BRANCO" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblOkF.BackColor = NL_DESLIGADO" & vbCrLf
    s = s & "        lblOkT.ForeColor = NL_BRANCO" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' CLIQUES" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "Public Sub TratarClique(ByVal chave As String)" & vbCrLf
    s = s & "    Select Case chave" & vbCrLf
    s = s & "        Case ""cancelar""" & vbCrLf
    s = s & "            RaiseEvent Fechar(0)" & vbCrLf
    s = s & "        Case ""ok""" & vbCrLf
    s = s & "            Select Case mSituacao" & vbCrLf
    s = s & "                Case S_CRONO, S_AUX" & vbCrLf
    s = s & "                    Incluir" & vbCrLf
    s = s & "                Case S_NA_Obras" & vbCrLf
    s = s & "                    RaiseEvent Fechar(mLinhaObras)" & vbCrLf
    s = s & "                Case S_NADA" & vbCrLf
    s = s & "                    lblStatus.ForeColor = NL_CORAL" & vbCrLf
    s = s & "                    lblStatus.Caption = ""Obra n" & ChrW(227) & "o encontrada na base: cadastre-a e atualize os dados antes de incluir.""" & vbCrLf
    s = s & "                Case Else" & vbCrLf
    s = s & "                    lblStatus.ForeColor = NL_CORAL" & vbCrLf
    s = s & "                    lblStatus.Caption = ""Digite o C" & ChrW(211) & "DIGO.""" & vbCrLf
    s = s & "            End Select" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Incluir()" & vbCrLf
    s = s & "    Dim tb As ListObject" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim cA As Long" & vbCrLf
    s = s & "    Dim cD As Long" & vbCrLf
    s = s & "    Dim cel As Range" & vbCrLf
    CodNovaClasse8 = s
End Function

Private Function CodNovaClasse9() As String
    Dim s As String
    s = s & "    Dim ev As Boolean" & vbCrLf
    s = s & "    Dim desc As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Falha" & vbCrLf
    s = s & "    Set tb = TabelaNL(T_Obras)" & vbCrLf
    s = s & "    If tb Is Nothing Then" & vbCrLf
    s = s & "        lblStatus.ForeColor = NL_CORAL" & vbCrLf
    s = s & "        lblStatus.Caption = ""A tabela tb_Obras n" & ChrW(227) & "o foi encontrada.""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If LinhaNaTabela(T_Obras, mCod) > 0 Then" & vbCrLf
    s = s & "        Conferir" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    cA = ColNL(tb, ""C" & ChrW(211) & "DIGO"")" & vbCrLf
    s = s & "    cD = ColNL(tb, ""NOME"")" & vbCrLf
    s = s & "    ev = Application.EnableEvents" & vbCrLf
    s = s & "    Application.EnableEvents = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    i = LinhaNova(tb, cA)" & vbCrLf
    s = s & "    If i = 0 Then" & vbCrLf
    s = s & "        Application.EnableEvents = ev" & vbCrLf
    s = s & "        lblStatus.ForeColor = NL_CORAL" & vbCrLf
    s = s & "        lblStatus.Caption = ""O Excel recusou a nova linha: confira se a aba OBRAS n" & ChrW(227) & "o est" & ChrW(225) & " protegida e limpe o filtro da tabela.""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set cel = tb.DataBodyRange.Cells(i, cA)" & vbCrLf
    s = s & "    cel.NumberFormat = ""@""" & vbCrLf
    s = s & "    cel.Value = CStr(mCod)" & vbCrLf
    s = s & "    Application.EnableEvents = ev" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Application.Calculate" & vbCrLf
    s = s & "    On Error GoTo Falha" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If cD > 0 Then desc = TextoNL(tb.DataBodyRange.Cells(i, cD).Value)" & vbCrLf
    s = s & "    lblStatus.ForeColor = NL_CINZA" & vbCrLf
    s = s & "    lblStatus.Caption = """"" & vbCrLf
    s = s & "    RaiseEvent Fechar(i)" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "Falha:" & vbCrLf
    s = s & "    Application.EnableEvents = True" & vbCrLf
    s = s & "    lblStatus.ForeColor = NL_CORAL" & vbCrLf
    s = s & "    lblStatus.Caption = ""N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel incluir: "" & Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodNovaClasse9 = s
End Function

Private Function CodNovaClasse10() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "' Linha para a obra nova: reaproveita uma linha sem C" & ChrW(211) & "DIGO ou cria outra no" & vbCrLf
    s = s & "' fim, repetindo as f" & ChrW(243) & "rmulas da linha de cima (0 = o Excel recusou)." & vbCrLf
    s = s & "Private Function LinhaNova(ByVal tb As ListObject, ByVal cA As Long) As Long" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim antes As Long" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Dim ref As Range" & vbCrLf
    s = s & "    Dim nova As Range" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not tb.DataBodyRange Is Nothing Then" & vbCrLf
    s = s & "        For i = 1 To tb.ListRows.Count" & vbCrLf
    s = s & "            If Len(TextoNL(tb.DataBodyRange.Cells(i, cA).Value)) = 0 Then" & vbCrLf
    s = s & "                If Not tb.DataBodyRange.Cells(i, cA).HasFormula Then" & vbCrLf
    s = s & "                    LinhaNova = i" & vbCrLf
    s = s & "                    Exit Function" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        Next i" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    antes = tb.ListRows.Count" & vbCrLf
    s = s & "    If antes > 0 Then Set ref = tb.ListRows(antes).Range" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    tb.ListRows.Add" & vbCrLf
    s = s & "    If tb.ListRows.Count <= antes Then" & vbCrLf
    s = s & "        Err.Clear" & vbCrLf
    s = s & "        tb.Resize tb.Range.Resize(tb.Range.Rows.Count + 1)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "    If tb.ListRows.Count <= antes Then Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set nova = tb.ListRows(tb.ListRows.Count).Range" & vbCrLf
    s = s & "    If Not ref Is Nothing Then" & vbCrLf
    s = s & "        For c = 1 To tb.ListColumns.Count" & vbCrLf
    s = s & "            If c <> cA Then" & vbCrLf
    s = s & "                If ref.Cells(1, c).HasFormula And Not nova.Cells(1, c).HasFormula Then" & vbCrLf
    s = s & "                    On Error Resume Next" & vbCrLf
    s = s & "                    nova.Cells(1, c).Formula2 = ref.Cells(1, c).Formula2" & vbCrLf
    s = s & "                    If Err.Number <> 0 Then" & vbCrLf
    s = s & "                        Err.Clear" & vbCrLf
    s = s & "                        nova.Cells(1, c).Formula = ref.Cells(1, c).Formula" & vbCrLf
    s = s & "                    End If" & vbCrLf
    s = s & "                    On Error GoTo 0" & vbCrLf
    s = s & "                End If" & vbCrLf
    CodNovaClasse10 = s
End Function

Private Function CodNovaClasse11() As String
    Dim s As String
    s = s & "            End If" & vbCrLf
    s = s & "        Next c" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    LinhaNova = tb.ListRows.Count" & vbCrLf
    s = s & "End Function" & vbCrLf
    CodNovaClasse11 = s
End Function

Private Function CodigoIaModulo() As String
    Dim s As String
    s = s & CodIaModulo1()
    s = s & CodIaModulo2()
    CodigoIaModulo = s
End Function

Private Function CodIaModulo1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' mEntregues - obra entregue: do follow up para ENTREGUES" & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "'   - " & ChrW(225) & "rea OBRA ENTREGUE dentro do Painel de Obras (cEntregaPainel)" & vbCrLf
    s = s & "'   - ArquivarEntregues: bot" & ChrW(227) & "o em lote da aba OBRAS (SITUA" & ChrW(199) & "" & ChrW(195) & "O = ENTREGUE)," & vbCrLf
    s = s & "'     agora copiando as colunas pelo nome" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Const IA_NAVY As Long = 3877150" & vbCrLf
    s = s & "Public Const IA_AZUL As Long = 8950797" & vbCrLf
    s = s & "Public Const IA_CORAL As Long = 2500316" & vbCrLf
    s = s & "Public Const IA_AMARELO As Long = 2408443" & vbCrLf
    s = s & "Public Const IA_CINZA As Long = 9139300" & vbCrLf
    s = s & "Public Const IA_FUNDO As Long = 16579320" & vbCrLf
    s = s & "Public Const IA_BORDA As Long = 15788258" & vbCrLf
    s = s & "Public Const IA_BRANCO As Long = 16777215" & vbCrLf
    s = s & "Public Const IA_VERDE As Long = 4891414" & vbCrLf
    s = s & "Public Const IA_DESTAQUE As Long = 15465471" & vbCrLf
    s = s & "Public Const IA_VERDE_FUNDO As Long = 15595240" & vbCrLf
    s = s & "Public Const IA_CORAL_FUNDO As Long = 14869246" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mFonteTexto As String" & vbCrLf
    s = s & "Private mFonteTitulo As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Bot" & ChrW(227) & "o em lote da aba OBRAS: todas as obras com SITUA" & ChrW(199) & "" & ChrW(195) & "O = ENTREGUE." & vbCrLf
    s = s & "Public Sub ArquivarEntregues()" & vbCrLf
    s = s & "    Dim arq As cArquivoObras" & vbCrLf
    s = s & "    Dim lista As Collection" & vbCrLf
    s = s & "    Dim v As Variant" & vbCrLf
    s = s & "    Dim ok As Long" & vbCrLf
    s = s & "    Dim falhas As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set arq = New cArquivoObras" & vbCrLf
    s = s & "    If Not arq.TemArquivo() Then" & vbCrLf
    s = s & "        MsgBox ""A tabela de entregues (tb_Entregues) n" & ChrW(227) & "o foi encontrada."", vbExclamation, ""Arquivar entregues""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Set lista = arq.CodObraEntregues()" & vbCrLf
    s = s & "    If lista.Count = 0 Then" & vbCrLf
    s = s & "        MsgBox ""Nenhuma obra do follow up est" & ChrW(225) & " com SITUA" & ChrW(199) & "" & ChrW(195) & "O ENTREGUE."" & vbCrLf & vbCrLf & ""Para mover uma obra espec" & ChrW(237) & "fica, use o bot" & ChrW(227) & "o OBRA ENTREGUE do Painel de Obras."", vbInformation, ""Arquivar entregues""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If MsgBox(lista.Count & "" obra(s) com SITUA" & ChrW(199) & "" & ChrW(195) & "O ENTREGUE v" & ChrW(227) & "o para ENTREGUES (colunas casadas pelo nome) e sair" & ChrW(227) & "o do follow up."" & vbCrLf & vbCrLf & ""Continuar?"", vbQuestion + vbYesNo, ""Arquivar entregues"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    CodIaModulo1 = s
End Function

Private Function CodIaModulo2() As String
    Dim s As String
    s = s & "    For Each v In lista" & vbCrLf
    s = s & "        If arq.Entregar(CStr(v)) Then" & vbCrLf
    s = s & "            ok = ok + 1" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            falhas = falhas & vbCrLf & ""C" & ChrW(211) & "DIGO "" & v & "": "" & arq.UltimoErro" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next v" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Len(falhas) = 0 Then" & vbCrLf
    s = s & "        MsgBox ok & "" obra(s) movida(s) para ENTREGUES."" & vbCrLf & arq.AvisoFiltro, vbInformation, ""Arquivar entregues""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        MsgBox ok & "" obra(s) movida(s). N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel mover:"" & falhas, vbExclamation, ""Arquivar entregues""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function FonteIA() As String" & vbCrLf
    s = s & "    If Len(mFonteTexto) = 0 Then DetectarFontes" & vbCrLf
    s = s & "    FonteIA = mFonteTexto" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function FonteTituloIA() As String" & vbCrLf
    s = s & "    If Len(mFonteTitulo) = 0 Then DetectarFontes" & vbCrLf
    s = s & "    FonteTituloIA = mFonteTitulo" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub DetectarFontes()" & vbCrLf
    s = s & "    Dim achado As String" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    achado = Dir$(Environ$(""WINDIR"") & ""\Fonts\Segoe UI*"")" & vbCrLf
    s = s & "    If Len(achado) = 0 Then achado = Dir$(Environ$(""LOCALAPPDATA"") & ""\Microsoft\Windows\Fonts\Segoe UI*"")" & vbCrLf
    s = s & "    If Len(achado) > 0 Then" & vbCrLf
    s = s & "        mFonteTexto = ""Segoe UI""" & vbCrLf
    s = s & "        mFonteTitulo = ""Segoe UI Black""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        mFonteTexto = ""Segoe UI""" & vbCrLf
    s = s & "        mFonteTitulo = ""Segoe UI Black""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodIaModulo2 = s
End Function

Private Function CodigoIaClique() As String
    Dim s As String
    s = s & CodIaClique1()
    CodigoIaClique = s
End Function

Private Function CodIaClique1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' cEntregaClique - repassa " & ChrW(224) & " " & ChrW(225) & "rea Obra entregue o clique de um r" & ChrW(243) & "tulo criado por" & vbCrLf
    s = s & "' c" & ChrW(243) & "digo (bot" & ChrW(245) & "es, c" & ChrW(237) & "rculos do farol): dono.TratarClique chave" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private WithEvents mRotulo As MSForms.Label" & vbCrLf
    s = s & "Private mDono As Object" & vbCrLf
    s = s & "Private mChave As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Conectar(ByVal rotulo As MSForms.Label, ByVal chave As String, ByVal dono As Object)" & vbCrLf
    s = s & "    Set mRotulo = rotulo" & vbCrLf
    s = s & "    mChave = chave" & vbCrLf
    s = s & "    Set mDono = dono" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Desconectar()" & vbCrLf
    s = s & "    Set mRotulo = Nothing" & vbCrLf
    s = s & "    Set mDono = Nothing" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub mRotulo_Click()" & vbCrLf
    s = s & "    If Not mDono Is Nothing Then mDono.TratarClique mChave" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodIaClique1 = s
End Function

Private Function CodigoIaArquivo() As String
    Dim s As String
    s = s & CodIaArquivo1()
    s = s & CodIaArquivo2()
    s = s & CodIaArquivo3()
    s = s & CodIaArquivo4()
    s = s & CodIaArquivo5()
    s = s & CodIaArquivo6()
    s = s & CodIaArquivo7()
    s = s & CodIaArquivo8()
    s = s & CodIaArquivo9()
    s = s & CodIaArquivo10()
    s = s & CodIaArquivo11()
    s = s & CodIaArquivo12()
    CodigoIaArquivo = s
End Function

Private Function CodIaArquivo1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' cArquivoObras - leva obras entregues da tb_Obras para ENTREGUES" & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "'   - a linha inteira vai para a tabela tb_Entregues, casando as colunas pelo" & vbCrLf
    s = s & "'     NOME (f" & ChrW(243) & "rmulas entram como valor; C" & ChrW(211) & "DIGO como texto); colunas que n" & ChrW(227) & "o" & vbCrLf
    s = s & "'     existem l" & ChrW(225) & " ficam de fora e s" & ChrW(227) & "o informadas" & vbCrLf
    s = s & "'   - registra no hist" & ChrW(243) & "rico de altera" & ChrW(231) & "" & ChrW(245) & "es quem moveu e quando" & vbCrLf
    s = s & "'   - conclui o marcador ""para analisar"" da obra, se houver" & vbCrLf
    s = s & "'   - filtros e segmenta" & ChrW(231) & "" & ChrW(245) & "es s" & ChrW(227) & "o guardados, limpos e reaplicados; a sa" & ChrW(237) & "da" & vbCrLf
    s = s & "'     da obra do follow up " & ChrW(233) & " conferida" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const T_Obras As String = ""tb_Obras""" & vbCrLf
    s = s & "Private Const T_ENTREGA As String = ""tb_Entregues""" & vbCrLf
    s = s & "Private Const T_HIST As String = ""tb_Alteracoes""" & vbCrLf
    s = s & "Private Const T_MARC As String = ""tb_Marcadores""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mErro As String" & vbCrLf
    s = s & "Private mAvisoFiltro As String" & vbCrLf
    s = s & "Private mFiltros As Collection" & vbCrLf
    s = s & "Private mSlicers As Collection" & vbCrLf
    s = s & "Private mEventos As Boolean" & vbCrLf
    s = s & "Private mTela As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get UltimoErro() As String" & vbCrLf
    s = s & "    UltimoErro = mErro" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Observa" & ChrW(231) & "" & ChrW(227) & "o sobre filtros que n" & ChrW(227) & "o puderam ser refeitos (vazio = tudo certo)." & vbCrLf
    s = s & "Public Property Get AvisoFiltro() As String" & vbCrLf
    s = s & "    AvisoFiltro = mAvisoFiltro" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Tabelas e colunas" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Tabelas e colunas" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "Public Function Tabela(ByVal nome As String) As ListObject" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Dim lo As ListObject" & vbCrLf
    s = s & "    For Each ws In ThisWorkbook.Worksheets" & vbCrLf
    s = s & "        Set lo = Nothing" & vbCrLf
    CodIaArquivo1 = s
End Function

Private Function CodIaArquivo2() As String
    Dim s As String
    s = s & "        On Error Resume Next" & vbCrLf
    s = s & "        Set lo = ws.ListObjects(nome)" & vbCrLf
    s = s & "        On Error GoTo 0" & vbCrLf
    s = s & "        If Not lo Is Nothing Then" & vbCrLf
    s = s & "            Set Tabela = lo" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next ws" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function Norm(ByVal s As String) As String" & vbCrLf
    s = s & "    s = Replace(Replace(Replace(s, vbCr, "" ""), vbLf, "" ""), Chr$(160), "" "")" & vbCrLf
    s = s & "    Do While InStr(1, s, ""  "", vbBinaryCompare) > 0" & vbCrLf
    s = s & "        s = Replace(s, ""  "", "" "")" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "    Norm = UCase$(Trim$(s))" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function Col(ByVal tb As ListObject, ByVal cabecalho As String) As Long" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim alvo As String" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Function" & vbCrLf
    s = s & "    alvo = Norm(cabecalho)" & vbCrLf
    s = s & "    For i = 1 To tb.ListColumns.Count" & vbCrLf
    s = s & "        If Norm(tb.ListColumns(i).Name) = alvo Then" & vbCrLf
    s = s & "            Col = i" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function Chave(ByVal v As Variant) As String" & vbCrLf
    s = s & "    If IsError(v) Then Exit Function" & vbCrLf
    s = s & "    If IsEmpty(v) Then Exit Function" & vbCrLf
    s = s & "    Chave = Trim$(Replace(CStr(v), Chr$(160), "" ""))" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Linha da obra na tabela, procurando o C" & ChrW(211) & "DIGO como texto (0 = n" & ChrW(227) & "o achou)." & vbCrLf
    s = s & "Public Function LinhaDe(ByVal tb As ListObject, ByVal codObra As String) As Long" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim alvo As String" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Function" & vbCrLf
    s = s & "    If tb.ListRows.Count = 0 Then Exit Function" & vbCrLf
    s = s & "    c = Col(tb, ""C" & ChrW(211) & "DIGO"")" & vbCrLf
    CodIaArquivo2 = s
End Function

Private Function CodIaArquivo3() As String
    Dim s As String
    s = s & "    If c = 0 Then Exit Function" & vbCrLf
    s = s & "    alvo = Trim$(codObra)" & vbCrLf
    s = s & "    If Len(alvo) = 0 Then Exit Function" & vbCrLf
    s = s & "    For i = 1 To tb.ListRows.Count" & vbCrLf
    s = s & "        If Chave(tb.DataBodyRange.Cells(i, c).Value) = alvo Then" & vbCrLf
    s = s & "            LinhaDe = i" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Consultas" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Valor da coluna na obra do follow up, como texto de exibi" & ChrW(231) & "" & ChrW(227) & "o." & vbCrLf
    s = s & "Public Function Texto(ByVal codObra As String, ByVal coluna As String) As String" & vbCrLf
    s = s & "    Dim tb As ListObject" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Dim v As Variant" & vbCrLf
    s = s & "    On Error GoTo Fim" & vbCrLf
    s = s & "    Set tb = Tabela(T_Obras)" & vbCrLf
    s = s & "    i = LinhaDe(tb, codObra)" & vbCrLf
    s = s & "    c = Col(tb, coluna)" & vbCrLf
    s = s & "    If i = 0 Or c = 0 Then Exit Function" & vbCrLf
    s = s & "    v = tb.DataBodyRange.Cells(i, c).Value" & vbCrLf
    s = s & "    If IsError(v) Then Exit Function" & vbCrLf
    s = s & "    If IsEmpty(v) Then Exit Function" & vbCrLf
    s = s & "    If VarType(v) = vbDate Then" & vbCrLf
    s = s & "        Texto = Format$(v, ""dd\/mm\/yyyy"")" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Texto = Trim$(CStr(v))" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "Fim:" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Data de entrega da obra (Empty se n" & ChrW(227) & "o houver)." & vbCrLf
    s = s & "Public Function DataEntrega(ByVal codObra As String) As Variant" & vbCrLf
    s = s & "    Dim tb As ListObject" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Dim v As Variant" & vbCrLf
    s = s & "    DataEntrega = Empty" & vbCrLf
    s = s & "    On Error GoTo Fim" & vbCrLf
    s = s & "    Set tb = Tabela(T_Obras)" & vbCrLf
    CodIaArquivo3 = s
End Function

Private Function CodIaArquivo4() As String
    Dim s As String
    s = s & "    i = LinhaDe(tb, codObra)" & vbCrLf
    s = s & "    c = Col(tb, ""ENTREGA"")" & vbCrLf
    s = s & "    If i = 0 Or c = 0 Then Exit Function" & vbCrLf
    s = s & "    v = tb.DataBodyRange.Cells(i, c).Value" & vbCrLf
    s = s & "    If IsDate(v) Then DataEntrega = CDate(v)" & vbCrLf
    s = s & "Fim:" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Quantas colunas do follow up existem (pelo nome) no arquivo de entregues." & vbCrLf
    s = s & "Public Function ColunasQueVao() As Long" & vbCrLf
    s = s & "    Dim tbF As ListObject" & vbCrLf
    s = s & "    Dim tbI As ListObject" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Set tbF = Tabela(T_Obras)" & vbCrLf
    s = s & "    Set tbI = Tabela(T_ENTREGA)" & vbCrLf
    s = s & "    If tbF Is Nothing Or tbI Is Nothing Then Exit Function" & vbCrLf
    s = s & "    For c = 1 To tbF.ListColumns.Count" & vbCrLf
    s = s & "        If Col(tbI, tbF.ListColumns(c).Name) > 0 Then ColunasQueVao = ColunasQueVao + 1" & vbCrLf
    s = s & "    Next c" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Colunas do follow up que n" & ChrW(227) & "o existem no arquivo ("""" = nenhuma)." & vbCrLf
    s = s & "Public Function ColunasFora() As String" & vbCrLf
    s = s & "    Dim tbF As ListObject" & vbCrLf
    s = s & "    Dim tbI As ListObject" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Set tbF = Tabela(T_Obras)" & vbCrLf
    s = s & "    Set tbI = Tabela(T_ENTREGA)" & vbCrLf
    s = s & "    If tbF Is Nothing Or tbI Is Nothing Then Exit Function" & vbCrLf
    s = s & "    For c = 1 To tbF.ListColumns.Count" & vbCrLf
    s = s & "        If Col(tbI, tbF.ListColumns(c).Name) = 0 Then" & vbCrLf
    s = s & "            If Len(ColunasFora) > 0 Then ColunasFora = ColunasFora & "", """ & vbCrLf
    s = s & "            ColunasFora = ColunasFora & tbF.ListColumns(c).Name" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next c" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function TemArquivo() As Boolean" & vbCrLf
    s = s & "    TemArquivo = Not (Tabela(T_ENTREGA) Is Nothing)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Quantas entradas com data (""dd/mm/aaaa ..."") a HIST" & ChrW(211) & "RICO da obra tem." & vbCrLf
    s = s & "Public Function QtdObservacoes(ByVal codObra As String) As Long" & vbCrLf
    s = s & "    Dim s As String" & vbCrLf
    s = s & "    Dim partes() As String" & vbCrLf
    CodIaArquivo4 = s
End Function

Private Function CodIaArquivo5() As String
    Dim s As String
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    s = Replace(Texto(codObra, ""HIST" & ChrW(211) & "RICO""), vbCr, """")" & vbCrLf
    s = s & "    If Len(s) = 0 Then Exit Function" & vbCrLf
    s = s & "    partes = Split(s, vbLf)" & vbCrLf
    s = s & "    For i = LBound(partes) To UBound(partes)" & vbCrLf
    s = s & "        If Trim$(partes(i)) Like ""##/##/####*"" Then QtdObservacoes = QtdObservacoes + 1" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' A obra tem marcador ""para analisar"" aberto?" & vbCrLf
    s = s & "Public Function TemMarcador(ByVal codObra As String) As Boolean" & vbCrLf
    s = s & "    TemMarcador = (LinhasMarcador(codObra, False) > 0)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Conta (e, se concluir = True, conclui) os marcadores abertos da obra." & vbCrLf
    s = s & "Private Function LinhasMarcador(ByVal codObra As String, ByVal fechar As Boolean) As Long" & vbCrLf
    s = s & "    Dim tb As ListObject" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim cA As Long" & vbCrLf
    s = s & "    Dim cS As Long" & vbCrLf
    s = s & "    Dim cP As Long" & vbCrLf
    s = s & "    Dim cE As Long" & vbCrLf
    s = s & "    On Error GoTo Fim" & vbCrLf
    s = s & "    Set tb = Tabela(T_MARC)" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Function" & vbCrLf
    s = s & "    If tb.DataBodyRange Is Nothing Then Exit Function" & vbCrLf
    s = s & "    cA = Col(tb, ""C" & ChrW(211) & "DIGO"")" & vbCrLf
    s = s & "    cS = Col(tb, ""SITUA" & ChrW(199) & "" & ChrW(195) & "O"")" & vbCrLf
    s = s & "    cP = Col(tb, ""CONCLU" & ChrW(205) & "DO POR"")" & vbCrLf
    s = s & "    cE = Col(tb, ""CONCLU" & ChrW(205) & "DO EM"")" & vbCrLf
    s = s & "    If cA = 0 Or cS = 0 Then Exit Function" & vbCrLf
    s = s & "    For i = 1 To tb.ListRows.Count" & vbCrLf
    s = s & "        If Chave(tb.DataBodyRange.Cells(i, cA).Value) = Trim$(codObra) Then" & vbCrLf
    s = s & "            If UCase$(Chave(tb.DataBodyRange.Cells(i, cS).Value)) = ""ABERTA"" Then" & vbCrLf
    s = s & "                LinhasMarcador = LinhasMarcador + 1" & vbCrLf
    s = s & "                If fechar Then" & vbCrLf
    s = s & "                    tb.DataBodyRange.Cells(i, cS).Value = ""CONCLU" & ChrW(205) & "DA""" & vbCrLf
    s = s & "                    If cP > 0 Then tb.DataBodyRange.Cells(i, cP).Value = Environ$(""Username"")" & vbCrLf
    s = s & "                    If cE > 0 Then" & vbCrLf
    s = s & "                        tb.DataBodyRange.Cells(i, cE).NumberFormat = ""dd/mm/yyyy hh:mm""" & vbCrLf
    s = s & "                        tb.DataBodyRange.Cells(i, cE).Value = Now" & vbCrLf
    s = s & "                    End If" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    CodIaArquivo5 = s
End Function

Private Function CodIaArquivo6() As String
    Dim s As String
    s = s & "    Next i" & vbCrLf
    s = s & "Fim:" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' C" & ChrW(211) & "DIGO das obras com SITUA" & ChrW(199) & "" & ChrW(195) & "O = ENTREGUE (bot" & ChrW(227) & "o em lote da aba OBRAS)." & vbCrLf
    s = s & "Public Function CodObraEntregues() As Collection" & vbCrLf
    s = s & "    Dim tb As ListObject" & vbCrLf
    s = s & "    Dim res As Collection" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim cS As Long" & vbCrLf
    s = s & "    Dim cA As Long" & vbCrLf
    s = s & "    Set res = New Collection" & vbCrLf
    s = s & "    Set CodObraEntregues = res" & vbCrLf
    s = s & "    Set tb = Tabela(T_Obras)" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Function" & vbCrLf
    s = s & "    If tb.ListRows.Count = 0 Then Exit Function" & vbCrLf
    s = s & "    cS = Col(tb, ""SITUA" & ChrW(199) & "" & ChrW(195) & "O"")" & vbCrLf
    s = s & "    cA = Col(tb, ""C" & ChrW(211) & "DIGO"")" & vbCrLf
    s = s & "    If cS = 0 Or cA = 0 Then Exit Function" & vbCrLf
    s = s & "    For i = 1 To tb.ListRows.Count" & vbCrLf
    s = s & "        If UCase$(Chave(tb.DataBodyRange.Cells(i, cS).Value)) = ""ENTREGUE"" Then" & vbCrLf
    s = s & "            If Len(Chave(tb.DataBodyRange.Cells(i, cA).Value)) > 0 Then res.Add Chave(tb.DataBodyRange.Cells(i, cA).Value)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Mover para Entregues" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "Public Function Entregar(ByVal codObra As String) As Boolean" & vbCrLf
    s = s & "    Dim tbF As ListObject" & vbCrLf
    s = s & "    Dim tbI As ListObject" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim etapa As String" & vbCrLf
    s = s & "    Dim ok As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Falha" & vbCrLf
    s = s & "    mErro = """"" & vbCrLf
    s = s & "    mAvisoFiltro = """"" & vbCrLf
    s = s & "    codObra = Trim$(codObra)" & vbCrLf
    s = s & "    Set tbF = Tabela(T_Obras)" & vbCrLf
    s = s & "    Set tbI = Tabela(T_ENTREGA)" & vbCrLf
    s = s & "    If tbI Is Nothing Then" & vbCrLf
    s = s & "        mErro = ""A tabela de entregues ("" & T_ENTREGA & "") n" & ChrW(227) & "o foi encontrada.""" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    CodIaArquivo6 = s
End Function

Private Function CodIaArquivo7() As String
    Dim s As String
    s = s & "    End If" & vbCrLf
    s = s & "    If LinhaDe(tbF, codObra) = 0 Then" & vbCrLf
    s = s & "        mErro = ""Obra n" & ChrW(227) & "o encontrada no follow up.""" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Preparar" & vbCrLf
    s = s & "    etapa = ""limpeza dos filtros""" & vbCrLf
    s = s & "    Set mFiltros = New Collection" & vbCrLf
    s = s & "    Set mSlicers = New Collection" & vbCrLf
    s = s & "    GuardarFiltros tbF" & vbCrLf
    s = s & "    GuardarFiltros tbI" & vbCrLf
    s = s & "    i = LinhaDe(tbF, codObra)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    etapa = ""c" & ChrW(243) & "pia para ENTREGUES""" & vbCrLf
    s = s & "    CopiarPorNome tbF, i, tbI" & vbCrLf
    s = s & "    If LinhaDe(tbI, codObra) = 0 Then Err.Raise vbObjectError + 610, , ""a obra n" & ChrW(227) & "o apareceu no arquivo de entregues""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    etapa = ""registro no hist" & ChrW(243) & "rico de altera" & ChrW(231) & "" & ChrW(245) & "es""" & vbCrLf
    s = s & "    Registrar tbF, i, ""MOVIMENTA" & ChrW(199) & "" & ChrW(195) & "O"", ""OBRAS"", ""ENTREGUES""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    etapa = ""marcador da obra""" & vbCrLf
    s = s & "    LinhasMarcador codObra, True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    etapa = ""sa" & ChrW(237) & "da da tb_Obras""" & vbCrLf
    s = s & "    tbF.ListRows(i).Delete" & vbCrLf
    s = s & "    If LinhaDe(tbF, codObra) > 0 Then" & vbCrLf
    s = s & "        mErro = ""A obra foi gravada em Entregues, mas n" & ChrW(227) & "o saiu do follow up. Apague a linha do follow up manualmente.""" & vbCrLf
    s = s & "        GoTo Fim" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    ok = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fim:" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    RestaurarFiltros" & vbCrLf
    s = s & "    Concluir" & vbCrLf
    s = s & "    Entregar = ok" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "Falha:" & vbCrLf
    s = s & "    mErro = ""A movimenta" & ChrW(231) & "" & ChrW(227) & "o parou na etapa "" & etapa & "": "" & Err.Description" & vbCrLf
    s = s & "    Resume Fim" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Linha i do follow up para uma linha nova do arquivo, coluna a coluna pelo nome." & vbCrLf
    s = s & "Private Sub CopiarPorNome(ByVal origem As ListObject, ByVal i As Long, ByVal destino As ListObject)" & vbCrLf
    CodIaArquivo7 = s
End Function

Private Function CodIaArquivo8() As String
    Dim s As String
    s = s & "    Dim nova As ListRow" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Dim d As Long" & vbCrLf
    s = s & "    Dim cel As Range" & vbCrLf
    s = s & "    Dim nome As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set nova = LinhaLivre(destino)" & vbCrLf
    s = s & "    For c = 1 To origem.ListColumns.Count" & vbCrLf
    s = s & "        nome = origem.ListColumns(c).Name" & vbCrLf
    s = s & "        d = Col(destino, nome)" & vbCrLf
    s = s & "        If d > 0 Then" & vbCrLf
    s = s & "            Set cel = nova.Range.Cells(1, d)" & vbCrLf
    s = s & "            If Not cel.HasFormula Then" & vbCrLf
    s = s & "                If Norm(nome) = ""C" & ChrW(211) & "DIGO"" Then" & vbCrLf
    s = s & "                    cel.NumberFormat = ""@""" & vbCrLf
    s = s & "                    cel.Value = Chave(origem.DataBodyRange.Cells(i, c).Value)" & vbCrLf
    s = s & "                Else" & vbCrLf
    s = s & "                    cel.NumberFormat = origem.DataBodyRange.Cells(i, c).NumberFormat" & vbCrLf
    s = s & "                    cel.Value = Seguro(origem.DataBodyRange.Cells(i, c).Value)" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next c" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Texto iniciado por = + - @ n" & ChrW(227) & "o pode virar f" & ChrW(243) & "rmula." & vbCrLf
    s = s & "Private Function Seguro(ByVal v As Variant) As Variant" & vbCrLf
    s = s & "    If VarType(v) = vbString Then" & vbCrLf
    s = s & "        If Len(v) > 0 Then" & vbCrLf
    s = s & "            If InStr(1, ""=+-@"", Left$(v, 1), vbBinaryCompare) > 0 Then v = ""'"" & v" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Seguro = v" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function LinhaLivre(ByVal tb As ListObject) As ListRow" & vbCrLf
    s = s & "    If tb.ListRows.Count = 1 Then" & vbCrLf
    s = s & "        If Application.WorksheetFunction.CountA(tb.ListRows(1).Range) = 0 Then" & vbCrLf
    s = s & "            Set LinhaLivre = tb.ListRows(1)" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Set LinhaLivre = tb.ListRows.Add" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    CodIaArquivo8 = s
End Function

Private Function CodIaArquivo9() As String
    Dim s As String
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Hist" & ChrW(243) & "rico de altera" & ChrW(231) & "" & ChrW(245) & "es" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "Private Sub Registrar(ByVal tbOrigem As ListObject, ByVal i As Long, ByVal campo As String, ByVal anterior As String, ByVal novo As String)" & vbCrLf
    s = s & "    Dim tbH As ListObject" & vbCrLf
    s = s & "    Dim lr As ListRow" & vbCrLf
    s = s & "    Dim nomes As Variant" & vbCrLf
    s = s & "    Dim k As Long" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Dim o As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Fim" & vbCrLf
    s = s & "    Set tbH = Tabela(T_HIST)" & vbCrLf
    s = s & "    If tbH Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    Set lr = LinhaLivre(tbH)" & vbCrLf
    s = s & "    nomes = Array(""C" & ChrW(211) & "DIGO"", ""NOME"", ""CLIENTE"", ""UF"", ""CIDADE"")" & vbCrLf
    s = s & "    For k = 0 To 4" & vbCrLf
    s = s & "        c = Col(tbH, CStr(nomes(k)))" & vbCrLf
    s = s & "        o = Col(tbOrigem, CStr(nomes(k)))" & vbCrLf
    s = s & "        If c > 0 And o > 0 Then" & vbCrLf
    s = s & "            If k = 0 Then lr.Range.Cells(1, c).NumberFormat = ""@""" & vbCrLf
    s = s & "            lr.Range.Cells(1, c).Value = Chave(tbOrigem.DataBodyRange.Cells(i, o).Value)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next k" & vbCrLf
    s = s & "    EscreverHist tbH, lr, ""CAMPO"", campo" & vbCrLf
    s = s & "    EscreverHist tbH, lr, ""ANTES"", anterior" & vbCrLf
    s = s & "    EscreverHist tbH, lr, ""DEPOIS"", novo" & vbCrLf
    s = s & "    EscreverHist tbH, lr, ""USU" & ChrW(193) & "RIO"", Environ$(""Username"")" & vbCrLf
    s = s & "    c = Col(tbH, ""DATA/HORA"")" & vbCrLf
    s = s & "    If c > 0 Then lr.Range.Cells(1, c).Value = Now" & vbCrLf
    s = s & "Fim:" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub EscreverHist(ByVal tbH As ListObject, ByVal lr As ListRow, ByVal coluna As String, ByVal valor As String)" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    c = Col(tbH, coluna)" & vbCrLf
    s = s & "    If c > 0 Then" & vbCrLf
    s = s & "        lr.Range.Cells(1, c).NumberFormat = ""@""" & vbCrLf
    s = s & "        lr.Range.Cells(1, c).Value = valor" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Filtros e segmenta" & ChrW(231) & "" & ChrW(245) & "es" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    CodIaArquivo9 = s
End Function

Private Function CodIaArquivo10() As String
    Dim s As String
    s = s & "Private Sub Preparar()" & vbCrLf
    s = s & "    mEventos = Application.EnableEvents" & vbCrLf
    s = s & "    mTela = Application.ScreenUpdating" & vbCrLf
    s = s & "    Application.EnableEvents = False" & vbCrLf
    s = s & "    Application.ScreenUpdating = False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Concluir()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Application.EnableEvents = mEventos" & vbCrLf
    s = s & "    Application.ScreenUpdating = mTela" & vbCrLf
    s = s & "    Application.EnableEvents = True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Guarda filtros e segmenta" & ChrW(231) & "" & ChrW(245) & "es da tabela e deixa todas as linhas vis" & ChrW(237) & "veis." & vbCrLf
    s = s & "Private Sub GuardarFiltros(ByVal tb As ListObject)" & vbCrLf
    s = s & "    Dim sc As Object" & vbCrLf
    s = s & "    Dim it As Object" & vbCrLf
    s = s & "    Dim sel As String" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim op As Long" & vbCrLf
    s = s & "    Dim c1 As Variant" & vbCrLf
    s = s & "    Dim c2 As Variant" & vbCrLf
    s = s & "    Dim nomeTabela As String" & vbCrLf
    s = s & "    Dim filtrado As Boolean" & vbCrLf
    s = s & "    Dim ligado As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    ' segmenta" & ChrW(231) & "" & ChrW(245) & "es (slicers) ligadas a esta tabela" & vbCrLf
    s = s & "    ' (o nome da tabela vai para uma vari" & ChrW(225) & "vel antes do If: com Resume Next," & vbCrLf
    s = s & "    ' um erro dentro da condi" & ChrW(231) & "" & ChrW(227) & "o faria o VBA entrar no bloco)" & vbCrLf
    s = s & "    For Each sc In ThisWorkbook.SlicerCaches" & vbCrLf
    s = s & "        nomeTabela = """"" & vbCrLf
    s = s & "        filtrado = False" & vbCrLf
    s = s & "        nomeTabela = sc.ListObject.Name" & vbCrLf
    s = s & "        filtrado = Not sc.FilterCleared" & vbCrLf
    s = s & "        Err.Clear" & vbCrLf
    s = s & "        If nomeTabela = tb.Name And filtrado Then" & vbCrLf
    s = s & "            sel = ""|""" & vbCrLf
    s = s & "            For Each it In sc.SlicerItems" & vbCrLf
    s = s & "                If it.Selected Then sel = sel & it.Name & ""|""" & vbCrLf
    s = s & "            Next it" & vbCrLf
    s = s & "            mSlicers.Add Array(sc, sel)" & vbCrLf
    s = s & "            sc.ClearManualFilter" & vbCrLf
    s = s & "        End If" & vbCrLf
    CodIaArquivo10 = s
End Function

Private Function CodIaArquivo11() As String
    Dim s As String
    s = s & "        Err.Clear" & vbCrLf
    s = s & "    Next sc" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' filtros comuns" & vbCrLf
    s = s & "    ligado = False" & vbCrLf
    s = s & "    ligado = tb.ShowAutoFilter" & vbCrLf
    s = s & "    If ligado Then ligado = tb.AutoFilter.FilterMode" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    If ligado Then" & vbCrLf
    s = s & "        For i = 1 To tb.AutoFilter.Filters.Count" & vbCrLf
    s = s & "            filtrado = False" & vbCrLf
    s = s & "            filtrado = tb.AutoFilter.Filters(i).On" & vbCrLf
    s = s & "            Err.Clear" & vbCrLf
    s = s & "            If filtrado Then" & vbCrLf
    s = s & "                op = 0" & vbCrLf
    s = s & "                c1 = Empty" & vbCrLf
    s = s & "                c2 = Empty" & vbCrLf
    s = s & "                op = tb.AutoFilter.Filters(i).Operator" & vbCrLf
    s = s & "                c1 = tb.AutoFilter.Filters(i).Criteria1" & vbCrLf
    s = s & "                If op = xlAnd Or op = xlOr Then c2 = tb.AutoFilter.Filters(i).Criteria2" & vbCrLf
    s = s & "                If Not IsEmpty(c1) Then mFiltros.Add Array(tb, i, op, c1, c2)" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "            Err.Clear" & vbCrLf
    s = s & "        Next i" & vbCrLf
    s = s & "        tb.AutoFilter.ShowAllData" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Reaplica o que foi guardado; o que n" & ChrW(227) & "o der para refazer vira aviso." & vbCrLf
    s = s & "Private Sub RestaurarFiltros()" & vbCrLf
    s = s & "    Dim item As Variant" & vbCrLf
    s = s & "    Dim tb As ListObject" & vbCrLf
    s = s & "    Dim sc As Object" & vbCrLf
    s = s & "    Dim it As Object" & vbCrLf
    s = s & "    Dim sel As String" & vbCrLf
    s = s & "    Dim falhou As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    If Not mFiltros Is Nothing Then" & vbCrLf
    s = s & "        For Each item In mFiltros" & vbCrLf
    s = s & "            Set tb = item(0)" & vbCrLf
    s = s & "            Err.Clear" & vbCrLf
    s = s & "            Select Case item(2)" & vbCrLf
    s = s & "                Case 0" & vbCrLf
    CodIaArquivo11 = s
End Function

Private Function CodIaArquivo12() As String
    Dim s As String
    s = s & "                    tb.Range.AutoFilter Field:=item(1), Criteria1:=item(3)" & vbCrLf
    s = s & "                Case xlAnd, xlOr" & vbCrLf
    s = s & "                    tb.Range.AutoFilter Field:=item(1), Criteria1:=item(3), Operator:=item(2), Criteria2:=item(4)" & vbCrLf
    s = s & "                Case Else" & vbCrLf
    s = s & "                    tb.Range.AutoFilter Field:=item(1), Criteria1:=item(3), Operator:=item(2)" & vbCrLf
    s = s & "            End Select" & vbCrLf
    s = s & "            If Err.Number <> 0 Then falhou = True" & vbCrLf
    s = s & "        Next item" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not mSlicers Is Nothing Then" & vbCrLf
    s = s & "        For Each item In mSlicers" & vbCrLf
    s = s & "            Set sc = item(0)" & vbCrLf
    s = s & "            sel = item(1)" & vbCrLf
    s = s & "            Err.Clear" & vbCrLf
    s = s & "            For Each it In sc.SlicerItems" & vbCrLf
    s = s & "                If InStr(1, sel, ""|"" & it.Name & ""|"", vbBinaryCompare) > 0 Then it.Selected = True" & vbCrLf
    s = s & "            Next it" & vbCrLf
    s = s & "            For Each it In sc.SlicerItems" & vbCrLf
    s = s & "                If InStr(1, sel, ""|"" & it.Name & ""|"", vbBinaryCompare) = 0 Then it.Selected = False" & vbCrLf
    s = s & "            Next it" & vbCrLf
    s = s & "            If Err.Number <> 0 Then falhou = True" & vbCrLf
    s = s & "        Next item" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If falhou Then mAvisoFiltro = ""Algum filtro ou segmenta" & ChrW(231) & "" & ChrW(227) & "o n" & ChrW(227) & "o p" & ChrW(244) & "de ser refeito; confira os filtros da aba.""" & vbCrLf
    s = s & "    Set mFiltros = Nothing" & vbCrLf
    s = s & "    Set mSlicers = Nothing" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodIaArquivo12 = s
End Function

Private Function CodigoIaClasse() As String
    Dim s As String
    s = s & CodIaClasse1()
    s = s & CodIaClasse2()
    s = s & CodIaClasse3()
    s = s & CodIaClasse4()
    s = s & CodIaClasse5()
    s = s & CodIaClasse6()
    s = s & CodIaClasse7()
    s = s & CodIaClasse8()
    CodigoIaClasse = s
End Function

Private Function CodIaClasse1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' cEntregaPainel - " & ChrW(225) & "rea OBRA ENTREGUE dentro do Painel de Obras" & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "' Ocupa a " & ChrW(225) & "rea de dados do painel: mostra a obra, a data de entrega e o" & vbCrLf
    s = s & "' que vai para ENTREGUES; pede confirma" & ChrW(231) & "" & ChrW(227) & "o ali mesmo e move." & vbCrLf
    s = s & "' Ao terminar, avisa o painel pelo evento Fechar (moveu = True se a obra" & vbCrLf
    s = s & "' saiu do follow up)." & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Event Fechar(ByVal moveu As Boolean)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' etapas da " & ChrW(225) & "rea" & vbCrLf
    s = s & "Private Const E_INICIO As Long = 0" & vbCrLf
    s = s & "Private Const E_CONFIRMA As Long = 1" & vbCrLf
    s = s & "Private Const E_FEITO As Long = 2" & vbCrLf
    s = s & "Private Const E_ERRO As Long = 3" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mFra As MSForms.Frame" & vbCrLf
    s = s & "Private mArq As cArquivoObras" & vbCrLf
    s = s & "Private mCliques As Collection" & vbCrLf
    s = s & "Private mFonte As String" & vbCrLf
    s = s & "Private mFonteTit As String" & vbCrLf
    s = s & "Private mCod As String" & vbCrLf
    s = s & "Private mNome As String" & vbCrLf
    s = s & "Private mEtapa As Long" & vbCrLf
    s = s & "Private mMoveu As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private lblTitulo As MSForms.Label" & vbCrLf
    s = s & "Private lblObra As MSForms.Label" & vbCrLf
    s = s & "Private lblLocal As MSForms.Label" & vbCrLf
    s = s & "Private lblCaixa As MSForms.Label" & vbCrLf
    s = s & "Private lblIcone As MSForms.Label" & vbCrLf
    s = s & "Private lblTit As MSForms.Label" & vbCrLf
    s = s & "Private lblSub As MSForms.Label" & vbCrLf
    s = s & "Private lblOkF As MSForms.Label" & vbCrLf
    s = s & "Private lblOkT As MSForms.Label" & vbCrLf
    s = s & "Private lblNaoF As MSForms.Label" & vbCrLf
    s = s & "Private lblNaoT As MSForms.Label" & vbCrLf
    s = s & "Private lblStatus As MSForms.Label" & vbCrLf
    s = s & "Private mG(1 To 7) As MSForms.Label" & vbCrLf
    s = s & "Private mT(1 To 7) As MSForms.Label" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' MONTAGEM" & vbCrLf
    CodIaClasse1 = s
End Function

Private Function CodIaClasse2() As String
    Dim s As String
    s = s & "'===============================================================================" & vbCrLf
    s = s & "Private Function Rot(ByVal nome As String, ByVal texto As String, _" & vbCrLf
    s = s & "                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _" & vbCrLf
    s = s & "                     ByVal cor As Long, ByVal tam As Single, ByVal negrito As Boolean, _" & vbCrLf
    s = s & "                     Optional ByVal fundo As Long = -1, Optional ByVal alinh As Long = 1) As MSForms.Label" & vbCrLf
    s = s & "    Dim c As MSForms.Label" & vbCrLf
    s = s & "    Set c = mFra.Controls.Add(""Forms.Label.1"", nome, True)" & vbCrLf
    s = s & "    c.Left = l" & vbCrLf
    s = s & "    c.Top = t" & vbCrLf
    s = s & "    c.Width = w" & vbCrLf
    s = s & "    c.Height = h" & vbCrLf
    s = s & "    c.Font.Name = mFonte" & vbCrLf
    s = s & "    c.Font.Size = tam" & vbCrLf
    s = s & "    c.Font.Bold = negrito" & vbCrLf
    s = s & "    c.ForeColor = cor" & vbCrLf
    s = s & "    If fundo >= 0 Then" & vbCrLf
    s = s & "        c.BackStyle = 1" & vbCrLf
    s = s & "        c.BackColor = fundo" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        c.BackStyle = 0" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    c.TextAlign = alinh" & vbCrLf
    s = s & "    c.WordWrap = False" & vbCrLf
    s = s & "    c.Caption = texto" & vbCrLf
    s = s & "    Set Rot = c" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Clique(ByVal rotulo As MSForms.Label, ByVal chave As String)" & vbCrLf
    s = s & "    Dim ev As cEntregaClique" & vbCrLf
    s = s & "    Set ev = New cEntregaClique" & vbCrLf
    s = s & "    ev.Conectar rotulo, chave, Me" & vbCrLf
    s = s & "    mCliques.Add ev" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Botao(ByVal nome As String, ByVal chave As String, ByVal texto As String, _" & vbCrLf
    s = s & "                  ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _" & vbCrLf
    s = s & "                  ByVal fundo As Long, ByVal cor As Long, ByVal borda As Long, _" & vbCrLf
    s = s & "                  ByRef rotFundo As MSForms.Label, ByRef rotTexto As MSForms.Label)" & vbCrLf
    s = s & "    Set rotFundo = Rot(nome & ""F"", """", l, t, w, h, cor, 8, False, fundo)" & vbCrLf
    s = s & "    If borda >= 0 Then" & vbCrLf
    s = s & "        rotFundo.BorderStyle = 1" & vbCrLf
    s = s & "        rotFundo.BorderColor = borda" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Set rotTexto = Rot(nome & ""T"", texto, l, t + (h - 12) / 2 - 0.5, w, 12, cor, 8.5, True, -1, 2)" & vbCrLf
    s = s & "    Clique rotFundo, chave" & vbCrLf
    CodIaClasse2 = s
End Function

Private Function CodIaClasse3() As String
    Dim s As String
    s = s & "    Clique rotTexto, chave" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Criar(ByVal fra As MSForms.Frame)" & vbCrLf
    s = s & "    Dim c As MSForms.Label" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim w As Single" & vbCrLf
    s = s & "    Dim xd As Single" & vbCrLf
    s = s & "    Dim wd As Single" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set mFra = fra" & vbCrLf
    s = s & "    Set mArq = New cArquivoObras" & vbCrLf
    s = s & "    Set mCliques = New Collection" & vbCrLf
    s = s & "    mFonte = FonteIA()" & vbCrLf
    s = s & "    mFonteTit = FonteTituloIA()" & vbCrLf
    s = s & "    w = mFra.InsideWidth" & vbCrLf
    s = s & "    If w < 600 Then w = mFra.Width - 4" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Rot ""iaFaixa"", """", 0, 0, w, 28, IA_BRANCO, 8, False, IA_NAVY" & vbCrLf
    s = s & "    Rot ""iaFaixaVd"", """", 0, 28, w, 3, IA_BRANCO, 8, False, IA_VERDE" & vbCrLf
    s = s & "    Set lblTitulo = Rot(""iaTitulo"", ""OBRA ENTREGUE"", 14, 6, w - 180, 18, IA_BRANCO, 11.5, False)" & vbCrLf
    s = s & "    lblTitulo.Font.Name = mFonteTit" & vbCrLf
    s = s & "    Set c = Rot(""iaFechar"", ""Voltar ao painel  "" & ChrW(215), w - 150, 8, 136, 14, IA_BRANCO, 9, True, -1, 3)" & vbCrLf
    s = s & "    Clique c, ""voltar""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Rot ""iaIntro"", ""A linha inteira vai para ENTREGUES (tabela tb_Entregues), com as colunas casadas pelo nome, e a obra sai do follow up."", 16, 40, w - 32, 14, IA_CINZA, 8.5, False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' esquerda: obra e situa" & ChrW(231) & "" & ChrW(227) & "o" & vbCrLf
    s = s & "    Rot ""iaObraRot"", ""OBRA"", 16, 64, 100, 12, IA_CINZA, 7.5, True" & vbCrLf
    s = s & "    Set lblObra = Rot(""iaObra"", """", 16, 78, 360, 20, IA_NAVY, 13, True)" & vbCrLf
    s = s & "    lblObra.Font.Name = mFonteTit" & vbCrLf
    s = s & "    lblObra.Font.Bold = False" & vbCrLf
    s = s & "    Set lblLocal = Rot(""iaLocal"", """", 16, 100, 360, 13, IA_CINZA, 8.5, False)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set lblCaixa = Rot(""iaCaixa"", """", 16, 120, 360, 170, IA_NAVY, 8, False, IA_BRANCO)" & vbCrLf
    s = s & "    lblCaixa.BorderStyle = 1" & vbCrLf
    s = s & "    lblCaixa.BorderColor = IA_BORDA" & vbCrLf
    s = s & "    Set lblIcone = Rot(""iaIcone"", """", 26, 130, 18, 18, IA_VERDE, 13, False)" & vbCrLf
    s = s & "    lblIcone.Font.Name = ""Wingdings""" & vbCrLf
    s = s & "    lblIcone.Font.Charset = 2" & vbCrLf
    s = s & "    Set lblTit = Rot(""iaTit"", """", 46, 130, 320, 17, IA_NAVY, 11, True)" & vbCrLf
    s = s & "    Set lblSub = Rot(""iaSub"", """", 46, 152, 320, 130, IA_CINZA, 8.5, False)" & vbCrLf
    s = s & "    lblSub.WordWrap = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Botao ""iaOk"", ""ok"", ""Mover para Entregues"", 16, 300, 360, 30, IA_NAVY, IA_BRANCO, -1, lblOkF, lblOkT" & vbCrLf
    CodIaClasse3 = s
End Function

Private Function CodIaClasse4() As String
    Dim s As String
    s = s & "    Botao ""iaNao"", ""voltar"", ""Cancelar e voltar ao painel"", 16, 338, 360, 26, IA_BRANCO, IA_NAVY, IA_BORDA, lblNaoF, lblNaoT" & vbCrLf
    s = s & "    Set lblStatus = Rot(""iaStatus"", """", 16, 372, 360, 13, IA_CINZA, 8, False)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' direita: o que vai para o arquivo e o que acontece depois" & vbCrLf
    s = s & "    xd = 396" & vbCrLf
    s = s & "    wd = w - xd - 16" & vbCrLf
    s = s & "    Rot ""iaSep"", """", xd - 10, 66, 1, 300, IA_NAVY, 6, False, IA_BORDA" & vbCrLf
    s = s & "    Rot ""iaDirTit"", ""O QUE VAI PARA ENTREGUES"", xd, 66, wd, 12, IA_CINZA, 7.5, True" & vbCrLf
    s = s & "    For i = 1 To 7" & vbCrLf
    s = s & "        Set mG(i) = Rot(""iaG"" & i, """", xd, 84 + (i - 1) * 24 + IIf(i >= 4, 16, 0), 16, 16, IA_VERDE, 11, False)" & vbCrLf
    s = s & "        mG(i).Font.Name = ""Wingdings""" & vbCrLf
    s = s & "        mG(i).Font.Charset = 2" & vbCrLf
    s = s & "        Set mT(i) = Rot(""iaT"" & i, """", xd + 20, 86 + (i - 1) * 24 + IIf(i >= 4, 16, 0), wd - 20, 14, IA_NAVY, 9, False)" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    Rot ""iaDepois"", ""DEPOIS"", xd, 160, wd, 12, IA_CINZA, 7.5, True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Desligar()" & vbCrLf
    s = s & "    Dim ev As cEntregaClique" & vbCrLf
    s = s & "    If mCliques Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    For Each ev In mCliques" & vbCrLf
    s = s & "        ev.Desconectar" & vbCrLf
    s = s & "    Next ev" & vbCrLf
    s = s & "    Set mCliques = Nothing" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' CONTE" & ChrW(218) & "DO" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' glifo ""*"" = marcador simples (bolinha); os demais s" & ChrW(227) & "o da fonte Wingdings" & vbCrLf
    s = s & "Private Sub Linha(ByVal i As Long, ByVal glifo As String, ByVal cor As Long, ByVal texto As String)" & vbCrLf
    s = s & "    If glifo = ""*"" Then" & vbCrLf
    s = s & "        mG(i).Font.Name = mFonte" & vbCrLf
    s = s & "        mG(i).Font.Charset = 0" & vbCrLf
    s = s & "        glifo = ChrW(8226)" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        mG(i).Font.Name = ""Wingdings""" & vbCrLf
    s = s & "        mG(i).Font.Charset = 2" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    mG(i).Caption = glifo" & vbCrLf
    s = s & "    mG(i).ForeColor = cor" & vbCrLf
    s = s & "    mT(i).Caption = texto" & vbCrLf
    s = s & "    mT(i).ControlTipText = texto" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    CodIaClasse4 = s
End Function

Private Function CodIaClasse5() As String
    Dim s As String
    s = s & "Private Sub Caixa(ByVal glifo As String, ByVal cor As Long, ByVal fundo As Long, ByVal titulo As String, ByVal detalhe As String)" & vbCrLf
    s = s & "    lblIcone.Caption = glifo" & vbCrLf
    s = s & "    lblIcone.ForeColor = cor" & vbCrLf
    s = s & "    lblCaixa.BackColor = fundo" & vbCrLf
    s = s & "    lblCaixa.BorderColor = cor" & vbCrLf
    s = s & "    lblTit.Caption = titulo" & vbCrLf
    s = s & "    lblSub.Caption = detalhe" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Botoes(ByVal textoOk As String, ByVal fundoOk As Long, ByVal textoNao As String)" & vbCrLf
    s = s & "    lblOkT.Caption = textoOk" & vbCrLf
    s = s & "    lblOkF.BackColor = fundoOk" & vbCrLf
    s = s & "    lblOkT.ForeColor = IA_BRANCO" & vbCrLf
    s = s & "    lblNaoF.Visible = (Len(textoNao) > 0)" & vbCrLf
    s = s & "    lblNaoT.Visible = (Len(textoNao) > 0)" & vbCrLf
    s = s & "    lblNaoT.Caption = textoNao" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Abre a " & ChrW(225) & "rea para a obra do painel." & vbCrLf
    s = s & "Public Sub Preparar(ByVal codObra As String)" & vbCrLf
    s = s & "    Dim d As Variant" & vbCrLf
    s = s & "    Dim fora As String" & vbCrLf
    s = s & "    Dim qtd As Long" & vbCrLf
    s = s & "    Dim lugar As String" & vbCrLf
    s = s & "    Dim v As String" & vbCrLf
    s = s & "    Dim dias As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mCod = Trim$(codObra)" & vbCrLf
    s = s & "    mMoveu = False" & vbCrLf
    s = s & "    mEtapa = E_INICIO" & vbCrLf
    s = s & "    lblStatus.Caption = """"" & vbCrLf
    s = s & "    mNome = mArq.Texto(mCod, ""NOME"")" & vbCrLf
    s = s & "    If Len(mNome) = 0 Then mNome = ""C" & ChrW(211) & "DIGO "" & mCod" & vbCrLf
    s = s & "    lblTitulo.Caption = ""OBRA ENTREGUE  "" & ChrW(183) & ""  "" & UCase$(mNome)" & vbCrLf
    s = s & "    lblObra.Caption = mNome" & vbCrLf
    s = s & "    lugar = ""C" & ChrW(211) & "DIGO "" & mCod" & vbCrLf
    s = s & "    v = mArq.Texto(mCod, ""CIDADE"")" & vbCrLf
    s = s & "    If Len(v) > 0 Then lugar = lugar & ""  "" & ChrW(183) & ""  "" & v" & vbCrLf
    s = s & "    v = mArq.Texto(mCod, ""UF"")" & vbCrLf
    s = s & "    If Len(v) > 0 Then lugar = lugar & "" - "" & v" & vbCrLf
    s = s & "    v = mArq.Texto(mCod, ""ETAPA"")" & vbCrLf
    s = s & "    If Len(v) > 0 Then lugar = lugar & ""  "" & ChrW(183) & ""  Etapa: "" & v" & vbCrLf
    s = s & "    lblLocal.Caption = lugar" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    d = mArq.DataEntrega(mCod)" & vbCrLf
    CodIaClasse5 = s
End Function

Private Function CodIaClasse6() As String
    Dim s As String
    s = s & "    If Not mArq.TemArquivo() Then" & vbCrLf
    s = s & "        Caixa Chr$(251), IA_CORAL, IA_CORAL_FUNDO, ""Tabela de entregues n" & ChrW(227) & "o encontrada"", ""A tabela tb_Entregues (aba ENTREGUES) n" & ChrW(227) & "o est" & ChrW(225) & " nesta planilha; nada pode ser movido.""" & vbCrLf
    s = s & "        Botoes ""Mover para Entregues (bloqueado)"", IA_BORDA, ""Voltar ao painel""" & vbCrLf
    s = s & "        mEtapa = E_ERRO" & vbCrLf
    s = s & "    ElseIf IsEmpty(d) Then" & vbCrLf
    s = s & "        Caixa Chr$(105), IA_AMARELO, IA_DESTAQUE, ""Sem data de entrega"", ""A coluna ENTREGA desta obra est" & ChrW(225) & " vazia (ela vem do Cronograma de Obras). Confirme que a obra foi entregue antes de mover.""" & vbCrLf
    s = s & "        Botoes ""Mover "" & Curto(mNome) & "" para Entregues"", IA_NAVY, ""Cancelar e voltar ao painel""" & vbCrLf
    s = s & "    ElseIf CDate(d) <= Date Then" & vbCrLf
    s = s & "        Caixa Chr$(252), IA_VERDE, IA_VERDE_FUNDO, ""Entregue em "" & Format$(d, ""dd\/mm\/yyyy""), ""Data de entrega do follow up (vem do Cronograma de Obras). A obra pode ir para o arquivo.""" & vbCrLf
    s = s & "        Botoes ""Mover "" & Curto(mNome) & "" para Entregues"", IA_NAVY, ""Cancelar e voltar ao painel""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        dias = CLng(CDate(d) - Date)" & vbCrLf
    s = s & "        Caixa Chr$(105), IA_AMARELO, IA_DESTAQUE, ""Entrega prevista para "" & Format$(d, ""dd\/mm\/yyyy""), ""A data do follow up ainda n" & ChrW(227) & "o chegou (faltam "" & dias & "" dia(s)). Se a obra j" & ChrW(225) & " foi entregue, pode mover; a data no arquivo ser" & ChrW(225) & " a que est" & ChrW(225) & " no follow up.""" & vbCrLf
    s = s & "        Botoes ""Mover "" & Curto(mNome) & "" para Entregues"", IA_NAVY, ""Cancelar e voltar ao painel""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If lblIcone.Caption = Chr$(105) Then" & vbCrLf
    s = s & "        lblIcone.Font.Name = ""Webdings""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblIcone.Font.Name = ""Wingdings""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    qtd = mArq.QtdObservacoes(mCod)" & vbCrLf
    s = s & "    fora = mArq.ColunasFora()" & vbCrLf
    s = s & "    Linha 1, Chr$(252), IA_VERDE, mArq.ColunasQueVao() & "" colunas casadas pelo nome (f" & ChrW(243) & "rmulas entram como valor)""" & vbCrLf
    s = s & "    Linha 2, Chr$(252), IA_VERDE, ""HIST" & ChrW(211) & "RICO ("" & qtd & "" registro(s)) e " & ChrW(218) & "LTIMO REGISTRO v" & ChrW(227) & "o junto""" & vbCrLf
    s = s & "    If Len(fora) > 0 Then" & vbCrLf
    s = s & "        Linha 3, Chr$(251), IA_CINZA, ""Ficam de fora (n" & ChrW(227) & "o existem no arquivo): "" & fora" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Linha 3, Chr$(252), IA_VERDE, ""Todas as colunas do follow up existem no arquivo""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Linha 4, ""*"", IA_AZUL, ""A obra sai do follow up e as de baixo sobem""" & vbCrLf
    s = s & "    Linha 5, ""*"", IA_AZUL, ""Fica no hist" & ChrW(243) & "rico de altera" & ChrW(231) & "" & ChrW(245) & "es quem moveu e quando""" & vbCrLf
    s = s & "    Linha 6, ""*"", IA_AZUL, ""Filtros da aba OBRAS s" & ChrW(227) & "o guardados e reaplicados no fim""" & vbCrLf
    s = s & "    If mArq.TemMarcador(mCod) Then" & vbCrLf
    s = s & "        Linha 7, ""*"", IA_AZUL, ""O marcador "" & ChrW(8220) & ""para analisar"" & ChrW(8221) & "" da obra " & ChrW(233) & " conclu" & ChrW(237) & "do""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Linha 7, ""*"", IA_AZUL, """"" & vbCrLf
    s = s & "        mG(7).Caption = """"" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function Curto(ByVal s As String) As String" & vbCrLf
    s = s & "    If Len(s) > 28 Then s = Left$(s, 26) & ""...""" & vbCrLf
    s = s & "    Curto = s" & vbCrLf
    s = s & "End Function" & vbCrLf
    CodIaClasse6 = s
End Function

Private Function CodIaClasse7() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' CLIQUES" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "Public Sub TratarClique(ByVal chave As String)" & vbCrLf
    s = s & "    Select Case chave" & vbCrLf
    s = s & "        Case ""voltar""" & vbCrLf
    s = s & "            RaiseEvent Fechar(mMoveu)" & vbCrLf
    s = s & "        Case ""ok""" & vbCrLf
    s = s & "            Select Case mEtapa" & vbCrLf
    s = s & "                Case E_INICIO" & vbCrLf
    s = s & "                    mEtapa = E_CONFIRMA" & vbCrLf
    s = s & "                    Caixa Chr$(105), IA_AMARELO, IA_DESTAQUE, ""Confirma a movimenta" & ChrW(231) & "" & ChrW(227) & "o?"", _" & vbCrLf
    s = s & "                        Curto(mNome) & "" (C" & ChrW(211) & "DIGO "" & mCod & "") vai para ENTREGUES e sai do follow up."" & vbLf & vbLf & ""N" & ChrW(227) & "o d" & ChrW(225) & " para desfazer com Ctrl+Z.""" & vbCrLf
    s = s & "                    lblIcone.Font.Name = ""Webdings""" & vbCrLf
    s = s & "                    Botoes ""Confirmar e mover para Entregues"", IA_VERDE, ""Cancelar e voltar ao painel""" & vbCrLf
    s = s & "                Case E_CONFIRMA" & vbCrLf
    s = s & "                    Mover" & vbCrLf
    s = s & "                Case E_FEITO, E_ERRO" & vbCrLf
    s = s & "                    RaiseEvent Fechar(mMoveu)" & vbCrLf
    s = s & "            End Select" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Mover()" & vbCrLf
    s = s & "    Dim ok As Boolean" & vbCrLf
    s = s & "    Dim aviso As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblStatus.ForeColor = IA_CINZA" & vbCrLf
    s = s & "    lblStatus.Caption = ""Movendo...""" & vbCrLf
    s = s & "    DoEvents" & vbCrLf
    s = s & "    ok = mArq.Entregar(mCod)" & vbCrLf
    s = s & "    aviso = mArq.AvisoFiltro" & vbCrLf
    s = s & "    lblStatus.Caption = """"" & vbCrLf
    s = s & "    lblIcone.Font.Name = ""Wingdings""" & vbCrLf
    s = s & "    If ok Then" & vbCrLf
    s = s & "        mMoveu = True" & vbCrLf
    s = s & "        mEtapa = E_FEITO" & vbCrLf
    s = s & "        Caixa Chr$(252), IA_VERDE, IA_VERDE_FUNDO, ""Movida para Entregues"", _" & vbCrLf
    s = s & "            Curto(mNome) & "" est" & ChrW(225) & " em ENTREGUES e saiu do follow up (conferido)."" & vbLf & _" & vbCrLf
    s = s & "            ""Registrado no hist" & ChrW(243) & "rico de altera" & ChrW(231) & "" & ChrW(245) & "es."" & IIf(Len(aviso) > 0, vbLf & vbLf & aviso, """")" & vbCrLf
    s = s & "        Botoes ""Voltar ao painel"", IA_NAVY, """"" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        mMoveu = (Len(mArq.UltimoErro) > 0 And InStr(1, mArq.UltimoErro, ""n" & ChrW(227) & "o saiu do follow up"", vbTextCompare) > 0)" & vbCrLf
    s = s & "        mEtapa = E_ERRO" & vbCrLf
    CodIaClasse7 = s
End Function

Private Function CodIaClasse8() As String
    Dim s As String
    s = s & "        Caixa Chr$(251), IA_CORAL, IA_CORAL_FUNDO, ""N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel concluir"", mArq.UltimoErro & IIf(Len(aviso) > 0, vbLf & vbLf & aviso, """")" & vbCrLf
    s = s & "        Botoes ""Voltar ao painel"", IA_NAVY, """"" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodIaClasse8 = s
End Function

Private Function CodigoBase() As String
    Dim s As String
    s = s & CodBase1()
    s = s & CodBase2()
    s = s & CodBase3()
    s = s & CodBase4()
    s = s & CodBase5()
    s = s & CodBase6()
    s = s & CodBase7()
    s = s & CodBase8()
    s = s & CodBase9()
    s = s & CodBase10()
    s = s & CodBase11()
    s = s & CodBase12()
    s = s & CodBase13()
    s = s & CodBase14()
    s = s & CodBase15()
    s = s & CodBase16()
    s = s & CodBase17()
    s = s & CodBase18()
    s = s & CodBase19()
    s = s & CodBase20()
    s = s & CodBase21()
    s = s & CodBase22()
    CodigoBase = s
End Function

Private Function CodBase1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' PAINEL - M" & ChrW(211) & "DULO PRINCIPAL" & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "' Abrir o painel ......: AbrirPainelObras" & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "' Regras:" & vbCrLf
    s = s & "'  - Colunas da tb_Obras ANTES do primeiro cabe" & ChrW(231) & "alho amarelo" & vbCrLf
    s = s & "'    = dados da obra (somente leitura no painel)" & vbCrLf
    s = s & "'  - Colunas com cabe" & ChrW(231) & "alho AMARELO = datas edit" & ChrW(225) & "veis (painel de datas)" & vbCrLf
    s = s & "'  - HIST" & ChrW(211) & "RICO = hist" & ChrW(243) & "rico / " & ChrW(218) & "LTIMO REGISTRO = " & ChrW(250) & "ltima observa" & ChrW(231) & "" & ChrW(227) & "o" & vbCrLf
    s = s & "'  - Toda grava" & ChrW(231) & "" & ChrW(227) & "o feita pelo painel " & ChrW(233) & " registrada na" & vbCrLf
    s = s & "'    tb_Alteracoes (com valor anterior correto)" & vbCrLf
    s = s & "'  - C" & ChrW(233) & "lulas com f" & ChrW(243) & "rmula nunca s" & ChrW(227) & "o sobrescritas" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'----------------------------- CONFIGURA" & ChrW(199) & "" & ChrW(195) & "O -------------------------" & vbCrLf
    s = s & "Public Const PO_TABELA As String = ""tb_Obras""" & vbCrLf
    s = s & "Public Const PO_TABELA_HIST As String = ""tb_Alteracoes""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Const PO_COL_CODIGO As String = ""C" & ChrW(211) & "DIGO""" & vbCrLf
    s = s & "Public Const PO_COL_DESCRICAO As String = ""NOME""" & vbCrLf
    s = s & "Public Const PO_COL_CIDADE As String = ""CIDADE""" & vbCrLf
    s = s & "Public Const PO_COL_UF As String = ""UF""" & vbCrLf
    s = s & "Public Const PO_COL_OBS As String = ""HIST" & ChrW(211) & "RICO""" & vbCrLf
    s = s & "Public Const PO_COL_ULTIMA As String = """ & ChrW(218) & "LTIMO REGISTRO""" & vbCrLf
    s = s & "Public Const PO_COL_MODALIDADE As String = ""TIPO DE OBRA""" & vbCrLf
    s = s & "Public Const PO_COL_FASE As String = ""ETAPA""" & vbCrLf
    s = s & "Public Const PO_COL_STATUS As String = ""SITUA" & ChrW(199) & "" & ChrW(195) & "O""" & vbCrLf
    s = s & "Public Const PO_COL_ENTREGA As String = ""ENTREGA""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Const PO_TAM_SEPARADOR As Long = 60" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Cores do painel" & vbCrLf
    s = s & "Public Const PO_COR_AZUL As Long = 3877150        'RGB(30, 41, 59)" & vbCrLf
    s = s & "Public Const PO_COR_AMARELO As Long = 2408443       'RGB(251, 191, 36)" & vbCrLf
    s = s & "Public Const PO_COR_FUNDO As Long = 16579320      'RGB(248, 250, 252)" & vbCrLf
    s = s & "Public Const PO_COR_LEITURA As Long = 15921906    'RGB(242, 242, 242)" & vbCrLf
    s = s & "Public Const PO_COR_EDICAO As Long = 14745599     'RGB(255, 255, 224)" & vbCrLf
    s = s & "Public Const PO_COR_ERRO As Long = 14935295       'RGB(255, 228, 227)" & vbCrLf
    s = s & "Public Const PO_COR_BORDA As Long = 14141640      'RGB(200, 200, 215)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    CodBase1 = s
End Function

Private Function CodBase2() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' LOCALIZA" & ChrW(199) & "" & ChrW(195) & "O DE TABELA / COLUNAS" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function PO_TabelaPorNome(ByVal nome As String) As ListObject" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Dim lo As ListObject" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For Each ws In ThisWorkbook.Worksheets" & vbCrLf
    s = s & "        For Each lo In ws.ListObjects" & vbCrLf
    s = s & "            If UCase$(lo.Name) = UCase$(nome) Then" & vbCrLf
    s = s & "                Set PO_TabelaPorNome = lo" & vbCrLf
    s = s & "                Exit Function" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        Next lo" & vbCrLf
    s = s & "    Next ws" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'" & ChrW(205) & "ndice da coluna pelo nome (0 = n" & ChrW(227) & "o existe)" & vbCrLf
    s = s & "Public Function PO_ColIdx(ByVal tb As ListObject, ByVal nome As String) As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim alvo As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    alvo = PO_Norm(nome)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 1 To tb.ListColumns.Count" & vbCrLf
    s = s & "        If PO_Norm(tb.ListColumns(i).Name) = alvo Then" & vbCrLf
    s = s & "            PO_ColIdx = i" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Cabe" & ChrW(231) & "alho pintado de amarelo?" & vbCrLf
    s = s & "Public Function PO_EhEditavel(ByVal tb As ListObject, ByVal idx As Long) As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim cor As Long" & vbCrLf
    s = s & "    Dim r As Long, g As Long, b As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Fim" & vbCrLf
    CodBase2 = s
End Function

Private Function CodBase3() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "    cor = tb.HeaderRowRange.Cells(1, idx).Interior.Color" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    r = cor Mod 256" & vbCrLf
    s = s & "    g = (cor \ 256) Mod 256" & vbCrLf
    s = s & "    b = (cor \ 65536) Mod 256" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    'Aceita tons de amarelo / ouro (FFC000, FFFF00, FFD966...)" & vbCrLf
    s = s & "    PO_EhEditavel = (r >= 230 And g >= 170 And b <= 210 And (r - b) >= 45)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fim:" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Lista (1..n) com os " & ChrW(237) & "ndices das colunas amarelas. Empty se n" & ChrW(227) & "o houver." & vbCrLf
    s = s & "Public Function PO_ColunasEditaveis(ByVal tb As ListObject) As Variant" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim i As Long, n As Long" & vbCrLf
    s = s & "    Dim arr() As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ReDim arr(1 To tb.ListColumns.Count)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 1 To tb.ListColumns.Count" & vbCrLf
    s = s & "        If PO_EhEditavel(tb, i) Then" & vbCrLf
    s = s & "            n = n + 1" & vbCrLf
    s = s & "            arr(n) = i" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If n = 0 Then" & vbCrLf
    s = s & "        PO_ColunasEditaveis = Empty" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        ReDim Preserve arr(1 To n)" & vbCrLf
    s = s & "        PO_ColunasEditaveis = arr" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Lista (1..n) com as colunas de ""dados da obra"" (antes do 1" & ChrW(186) & " amarelo)" & vbCrLf
    s = s & "Public Function PO_ColunasInfo(ByVal tb As ListObject) As Variant" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim i As Long, n As Long, ultima As Long" & vbCrLf
    s = s & "    Dim arr() As Long" & vbCrLf
    s = s & "    Dim ed As Variant" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ed = PO_ColunasEditaveis(tb)" & vbCrLf
    CodBase3 = s
End Function

Private Function CodBase4() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "    If IsArray(ed) Then" & vbCrLf
    s = s & "        ultima = ed(1) - 1" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        ultima = tb.ListColumns.Count" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If ultima < 1 Then" & vbCrLf
    s = s & "        PO_ColunasInfo = Empty" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ReDim arr(1 To ultima)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 1 To ultima" & vbCrLf
    s = s & "        If i <> PO_ColIdx(tb, PO_COL_OBS) And i <> PO_ColIdx(tb, PO_COL_ULTIMA) Then" & vbCrLf
    s = s & "            n = n + 1" & vbCrLf
    s = s & "            arr(n) = i" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If n = 0 Then" & vbCrLf
    s = s & "        PO_ColunasInfo = Empty" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        ReDim Preserve arr(1 To n)" & vbCrLf
    s = s & "        PO_ColunasInfo = arr" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Nome curto da etapa: ""DT PREV. ALVAR" & ChrW(193) & """ -> ""ALVAR" & ChrW(193) & """" & vbCrLf
    s = s & "Public Function PO_NomeEtapa(ByVal cabecalho As String) As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim s As String" & vbCrLf
    s = s & "    Dim mudou As Boolean" & vbCrLf
    s = s & "    Dim prefixos As Variant" & vbCrLf
    s = s & "    Dim p As Variant" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    prefixos = Array(""DT. "", ""DT "", ""PREV. "", ""PREV "", ""PREVIS" & ChrW(195) & "O "", ""PREVISTO "")" & vbCrLf
    s = s & "    s = UCase$(Trim$(cabecalho))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Do" & vbCrLf
    s = s & "        mudou = False" & vbCrLf
    s = s & "        For Each p In prefixos" & vbCrLf
    s = s & "            If Left$(s, Len(p)) = p Then" & vbCrLf
    CodBase4 = s
End Function

Private Function CodBase5() As String
    Dim s As String
    s = s & "                s = Trim$(Mid$(s, Len(p) + 1))" & vbCrLf
    s = s & "                mudou = True" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        Next p" & vbCrLf
    s = s & "    Loop While mudou" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If s = """" Then s = cabecalho" & vbCrLf
    s = s & "    PO_NomeEtapa = s" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' LEITURA / FORMATA" & ChrW(199) & "" & ChrW(195) & "O DE VALORES" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "Public Function PO_Valor(ByVal tb As ListObject, ByVal lr As ListRow, ByVal nomeCol As String) As Variant" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim idx As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    idx = PO_ColIdx(tb, nomeCol)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If idx = 0 Then" & vbCrLf
    s = s & "        PO_Valor = Empty" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        PO_Valor = lr.Range.Cells(1, idx).value" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function PO_TextoCampo(ByVal tb As ListObject, ByVal lr As ListRow, ByVal nomeCol As String) As String" & vbCrLf
    s = s & "    PO_TextoCampo = PO_TextoValor(PO_Valor(tb, lr, nomeCol))" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Converte qualquer valor de c" & ChrW(233) & "lula em texto de exibi" & ChrW(231) & "" & ChrW(227) & "o" & vbCrLf
    s = s & "Public Function PO_TextoValor(ByVal v As Variant) As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo vazio" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If IsError(v) Then GoTo vazio" & vbCrLf
    s = s & "    If IsEmpty(v) Then GoTo vazio" & vbCrLf
    s = s & "    If IsNull(v) Then GoTo vazio" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If VarType(v) = vbDate Then" & vbCrLf
    s = s & "        If CDbl(v) <= 1 Then GoTo vazio" & vbCrLf
    s = s & "        If CDbl(v) = Int(CDbl(v)) Then" & vbCrLf
    CodBase5 = s
End Function

Private Function CodBase6() As String
    Dim s As String
    s = s & "            PO_TextoValor = Format$(v, ""dd\/mm\/yyyy"")" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            PO_TextoValor = Format$(v, ""dd\/mm\/yyyy hh:mm"")" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    PO_TextoValor = Trim$(CStr(v))" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "vazio:" & vbCrLf
    s = s & "    PO_TextoValor = """"" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Mai" & ChrW(250) & "sculas, sem acento, sem quebras" & vbCrLf
    s = s & "Public Function PO_Norm(ByVal s As String) As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    s = Replace(s, vbCr, "" "")" & vbCrLf
    s = s & "    s = Replace(s, vbLf, "" "")" & vbCrLf
    s = s & "    PO_Norm = UCase$(Trim$(PO_SemAcento(s)))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function PO_SemAcento(ByVal s As String) As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Const COM As String = """ & ChrW(193) & "" & ChrW(192) & "" & ChrW(194) & "" & ChrW(195) & "" & ChrW(196) & "" & ChrW(201) & "" & ChrW(200) & "" & ChrW(202) & "" & ChrW(203) & "" & ChrW(205) & "" & ChrW(204) & "" & ChrW(206) & "" & ChrW(207) & "" & ChrW(211) & "" & ChrW(210) & "" & ChrW(212) & "" & ChrW(213) & "" & ChrW(214) & "" & ChrW(218) & "" & ChrW(217) & "" & ChrW(219) & "" & ChrW(220) & "" & ChrW(199) & "" & ChrW(225) & "" & ChrW(224) & "" & ChrW(226) & "" & ChrW(227) & "" & ChrW(228) & "" & ChrW(233) & "" & ChrW(232) & "" & ChrW(234) & "" & ChrW(235) & "" & ChrW(237) & "" & ChrW(236) & "" & ChrW(238) & "" & ChrW(239) & "" & ChrW(243) & "" & ChrW(242) & "" & ChrW(244) & "" & ChrW(245) & "" & ChrW(246) & "" & ChrW(250) & "" & ChrW(249) & "" & ChrW(251) & "" & ChrW(252) & "" & ChrW(231) & """" & vbCrLf
    s = s & "    Const SEM As String = ""AAAAAEEEEIIIIOOOOOUUUUCaaaaaeeeeiiiiooooouuuuc""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 1 To Len(COM)" & vbCrLf
    s = s & "        s = Replace(s, Mid$(COM, i, 1), Mid$(SEM, i, 1))" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    PO_SemAcento = s" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Quebras de linha no padr" & ChrW(227) & "o do TextBox / da c" & ChrW(233) & "lula (vbCrLf)" & vbCrLf
    s = s & "Public Function PO_ParaCaixa(ByVal s As String) As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    s = Replace(s, vbCrLf, vbLf)" & vbCrLf
    s = s & "    s = Replace(s, vbCr, vbLf)" & vbCrLf
    s = s & "    PO_ParaCaixa = Replace(s, vbLf, vbCrLf)" & vbCrLf
    CodBase6 = s
End Function

Private Function CodBase7() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Remove espa" & ChrW(231) & "os e quebras de linha no fim do texto" & vbCrLf
    s = s & "Public Function PO_TrimFim(ByVal s As String) As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Do While Len(s) > 0" & vbCrLf
    s = s & "        Select Case Right$(s, 1)" & vbCrLf
    s = s & "            Case vbCr, vbLf, "" "", vbTab" & vbCrLf
    s = s & "                s = Left$(s, Len(s) - 1)" & vbCrLf
    s = s & "            Case Else" & vbCrLf
    s = s & "                Exit Do" & vbCrLf
    s = s & "        End Select" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    PO_TrimFim = s" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Texto ""dd/mm/aaaa"", ""ddmmaaaa"", ""ddmmaa"" -> Date.  """" ou ""-"" = vazio." & vbCrLf
    s = s & "Public Function PO_TextoParaData(ByVal s As String, ByRef d As Date, ByRef vazio As Boolean) As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim p() As String" & vbCrLf
    s = s & "    Dim dd As Long, mm As Long, yy As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Invalido" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    s = Trim$(s)" & vbCrLf
    s = s & "    vazio = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If s = """" Or s = ""-"" Then" & vbCrLf
    s = s & "        vazio = True" & vbCrLf
    s = s & "        PO_TextoParaData = True" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If InStr(s, ""/"") > 0 Then" & vbCrLf
    s = s & "        p = Split(s, ""/"")" & vbCrLf
    s = s & "        If UBound(p) <> 2 Then GoTo Invalido" & vbCrLf
    s = s & "        If Not (p(0) Like ""#"" Or p(0) Like ""##"") Then GoTo Invalido" & vbCrLf
    s = s & "        If Not (p(1) Like ""#"" Or p(1) Like ""##"") Then GoTo Invalido" & vbCrLf
    s = s & "        If Not (p(2) Like ""##"" Or p(2) Like ""####"") Then GoTo Invalido" & vbCrLf
    s = s & "        dd = CLng(p(0)): mm = CLng(p(1)): yy = CLng(p(2))" & vbCrLf
    s = s & "    ElseIf s Like ""########"" Then" & vbCrLf
    s = s & "        dd = CLng(Left$(s, 2)): mm = CLng(Mid$(s, 3, 2)): yy = CLng(Right$(s, 4))" & vbCrLf
    CodBase7 = s
End Function

Private Function CodBase8() As String
    Dim s As String
    s = s & "    ElseIf s Like ""######"" Then" & vbCrLf
    s = s & "        dd = CLng(Left$(s, 2)): mm = CLng(Mid$(s, 3, 2)): yy = CLng(Right$(s, 2))" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        GoTo Invalido" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If yy < 100 Then yy = yy + 2000" & vbCrLf
    s = s & "    If yy < 1990 Or yy > 2100 Then GoTo Invalido" & vbCrLf
    s = s & "    If mm < 1 Or mm > 12 Then GoTo Invalido" & vbCrLf
    s = s & "    If dd < 1 Or dd > 31 Then GoTo Invalido" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    d = DateSerial(yy, mm, dd)" & vbCrLf
    s = s & "    If Day(d) <> dd Or Month(d) <> mm Then GoTo Invalido" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    PO_TextoParaData = True" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Invalido:" & vbCrLf
    s = s & "    PO_TextoParaData = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' GRAVA" & ChrW(199) & "" & ChrW(195) & "O NA C" & ChrW(201) & "LULA + LOG" & vbCrLf
    s = s & "'   Retorno:  1 = alterou | 0 = j" & ChrW(225) & " era igual | -1 = n" & ChrW(227) & "o gravou" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "Public Function PO_Gravar(ByVal tb As ListObject, ByVal lr As ListRow, _" & vbCrLf
    s = s & "                          ByVal idx As Long, ByVal novo As Variant) As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim cel As Range" & vbCrLf
    s = s & "    Dim antigo As Variant" & vbCrLf
    s = s & "    Dim evt As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    evt = Application.EnableEvents" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Falha" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set cel = lr.Range.Cells(1, idx)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If cel.HasFormula Then" & vbCrLf
    s = s & "        MsgBox ""O campo '"" & tb.ListColumns(idx).Name & ""' cont" & ChrW(233) & "m f" & ChrW(243) & "rmula e n" & ChrW(227) & "o foi alterado."", _" & vbCrLf
    s = s & "               vbExclamation, ""painel""" & vbCrLf
    s = s & "        PO_Gravar = -1" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    CodBase8 = s
End Function

Private Function CodBase9() As String
    Dim s As String
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    antigo = cel.value" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If PO_Iguais(antigo, novo) Then" & vbCrLf
    s = s & "        PO_Gravar = 0" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    'Eventos desligados: o log " & ChrW(233) & " feito aqui com o valor anterior correto" & vbCrLf
    s = s & "    Application.EnableEvents = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If IsEmpty(novo) Then" & vbCrLf
    s = s & "        cel.ClearContents" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        If VarType(novo) = vbDate Then" & vbCrLf
    s = s & "            If cel.NumberFormat = ""General"" Or cel.NumberFormat = ""Geral"" Then" & vbCrLf
    s = s & "                cel.NumberFormat = ""dd/mm/yyyy""" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        cel.value = novo" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Application.EnableEvents = evt" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    PO_RegistrarLog tb, lr, idx, antigo, novo" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    PO_Gravar = 1" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Falha:" & vbCrLf
    s = s & "    Application.EnableEvents = evt" & vbCrLf
    s = s & "    PO_Gravar = -1" & vbCrLf
    s = s & "    MsgBox ""N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel gravar o campo '"" & tb.ListColumns(idx).Name & ""'."" & _" & vbCrLf
    s = s & "           vbCrLf & vbCrLf & Err.Description, vbCritical, ""painel""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function PO_Vazio(ByVal v As Variant) As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If IsError(v) Then Exit Function" & vbCrLf
    s = s & "    If IsEmpty(v) Then" & vbCrLf
    s = s & "        PO_Vazio = True" & vbCrLf
    s = s & "    ElseIf VarType(v) = vbString Then" & vbCrLf
    s = s & "        PO_Vazio = (Trim$(v) = """" Or Trim$(v) = ""-"")" & vbCrLf
    CodBase9 = s
End Function

Private Function CodBase10() As String
    Dim s As String
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function PO_Iguais(ByVal a As Variant, ByVal b As Variant) As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Diferente" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If IsError(a) Or IsError(b) Then GoTo Diferente" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If PO_Vazio(a) And PO_Vazio(b) Then" & vbCrLf
    s = s & "        PO_Iguais = True" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If PO_Vazio(a) Or PO_Vazio(b) Then GoTo Diferente" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If VarType(a) = vbDate Or VarType(b) = vbDate Then" & vbCrLf
    s = s & "        If IsDate(a) And IsDate(b) Then" & vbCrLf
    s = s & "            PO_Iguais = (Int(CDbl(CDate(a))) = Int(CDbl(CDate(b))))" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    PO_Iguais = (CStr(a) = CStr(b))" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Diferente:" & vbCrLf
    s = s & "    PO_Iguais = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub PO_RegistrarLog(ByVal tb As ListObject, ByVal lr As ListRow, _" & vbCrLf
    s = s & "                            ByVal idx As Long, ByVal antigo As Variant, ByVal novo As Variant)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim th As ListObject" & vbCrLf
    s = s & "    Dim nl As ListRow" & vbCrLf
    s = s & "    Dim evt As Boolean" & vbCrLf
    s = s & "    Dim campos As Variant" & vbCrLf
    s = s & "    Dim c As Variant" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    evt = Application.EnableEvents" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Fim" & vbCrLf
    s = s & "" & vbCrLf
    CodBase10 = s
End Function

Private Function CodBase11() As String
    Dim s As String
    s = s & "    Set th = PO_TabelaPorNome(PO_TABELA_HIST)" & vbCrLf
    s = s & "    If th Is Nothing Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Application.EnableEvents = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set nl = th.ListRows.Add" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    campos = Array(""C" & ChrW(211) & "DIGO"", ""NOME"", ""CLIENTE"", ""UF"", ""CIDADE"")" & vbCrLf
    s = s & "    For Each c In campos" & vbCrLf
    s = s & "        PO_SetLog th, nl, CStr(c), PO_ValorLog(PO_Valor(tb, lr, CStr(c)))" & vbCrLf
    s = s & "    Next c" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    PO_SetLog th, nl, ""CAMPO"", tb.ListColumns(idx).Name" & vbCrLf
    s = s & "    PO_SetLog th, nl, ""ANTES"", PO_ValorLog(antigo)" & vbCrLf
    s = s & "    PO_SetLog th, nl, ""DEPOIS"", PO_ValorLog(novo)" & vbCrLf
    s = s & "    PO_SetLog th, nl, ""USU" & ChrW(193) & "RIO"", Environ$(""Username"")" & vbCrLf
    s = s & "    PO_SetLog th, nl, ""DATA/HORA"", Now" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fim:" & vbCrLf
    s = s & "    Application.EnableEvents = evt" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub PO_SetLog(ByVal th As ListObject, ByVal nl As ListRow, ByVal nomeCol As String, ByVal v As Variant)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim idx As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    idx = PO_ColIdx(th, nomeCol)" & vbCrLf
    s = s & "    If idx > 0 Then nl.Range.Cells(1, idx).value = v" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function PO_ValorLog(ByVal v As Variant) As Variant" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If IsError(v) Then" & vbCrLf
    s = s & "        PO_ValorLog = Empty" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        PO_ValorLog = v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' LEITURA DO HIST" & ChrW(211) & "RICO DE OBSERVA" & ChrW(199) & "" & ChrW(213) & "ES" & vbCrLf
    CodBase11 = s
End Function

Private Function CodBase12() As String
    Dim s As String
    s = s & "'   Cada item da cole" & ChrW(231) & "" & ChrW(227) & "o = Array(data As Date|Empty, rotulo, texto)" & vbCrLf
    s = s & "'   Uma entrada come" & ChrW(231) & "a com ""dd/mm/aaaa ..."" ; linhas ""-----"" separam." & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "Public Function PO_ParseObs(ByVal texto As String) As Collection" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim col As New Collection" & vbCrLf
    s = s & "    Dim linhas() As String" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim ln As String" & vbCrLf
    s = s & "    Dim temAtual As Boolean" & vbCrLf
    s = s & "    Dim dAtual As Variant" & vbCrLf
    s = s & "    Dim rotAtual As String" & vbCrLf
    s = s & "    Dim txtAtual As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    texto = Replace(texto, vbCrLf, vbLf)" & vbCrLf
    s = s & "    texto = Replace(texto, vbCr, vbLf)" & vbCrLf
    s = s & "    linhas = Split(texto, vbLf)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = LBound(linhas) To UBound(linhas)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        ln = Trim$(linhas(i))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If PO_EhSeparador(ln) Then" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            If temAtual Then PO_AddEntrada col, dAtual, rotAtual, txtAtual" & vbCrLf
    s = s & "            temAtual = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        ElseIf ln Like ""##/##/####*"" Then" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            If temAtual Then PO_AddEntrada col, dAtual, rotAtual, txtAtual" & vbCrLf
    s = s & "            PO_SepararCabecalho ln, dAtual, rotAtual, txtAtual" & vbCrLf
    s = s & "            temAtual = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        ElseIf temAtual Then" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            If txtAtual = """" Then" & vbCrLf
    s = s & "                txtAtual = ln" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                txtAtual = txtAtual & vbLf & ln" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        ElseIf ln <> """" Then" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            'Texto sem data no in" & ChrW(237) & "cio" & vbCrLf
    s = s & "            dAtual = Empty" & vbCrLf
    CodBase12 = s
End Function

Private Function CodBase13() As String
    Dim s As String
    s = s & "            rotAtual = """"" & vbCrLf
    s = s & "            txtAtual = ln" & vbCrLf
    s = s & "            temAtual = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If temAtual Then PO_AddEntrada col, dAtual, rotAtual, txtAtual" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set PO_ParseObs = col" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function PO_EhSeparador(ByVal ln As String) As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim s As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Len(ln) < 5 Then Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    s = ln" & vbCrLf
    s = s & "    s = Replace(s, ""-"", """")" & vbCrLf
    s = s & "    s = Replace(s, ""_"", """")" & vbCrLf
    s = s & "    s = Replace(s, ""="", """")" & vbCrLf
    s = s & "    s = Replace(s, ChrW(8211), """")" & vbCrLf
    s = s & "    s = Replace(s, ChrW(8212), """")" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    PO_EhSeparador = (Len(Trim$(s)) = 0)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub PO_SepararCabecalho(ByVal ln As String, ByRef d As Variant, _" & vbCrLf
    s = s & "                                ByRef rotulo As String, ByRef corpo As String)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim dt As Date" & vbCrLf
    s = s & "    Dim vz As Boolean" & vbCrLf
    s = s & "    Dim resto As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If PO_TextoParaData(Left$(ln, 10), dt, vz) And Not vz Then" & vbCrLf
    s = s & "        d = dt" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        d = Empty" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    rotulo = Left$(ln, 10)" & vbCrLf
    CodBase13 = s
End Function

Private Function CodBase14() As String
    Dim s As String
    s = s & "    resto = PO_TirarPrefixo(Mid$(ln, 11))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If resto Like ""##:##*"" Then" & vbCrLf
    s = s & "        rotulo = rotulo & "" "" & Left$(resto, 5)" & vbCrLf
    s = s & "        resto = Mid$(resto, 6)" & vbCrLf
    s = s & "        If resto Like "":##*"" Then resto = Mid$(resto, 4)   'segundos" & vbCrLf
    s = s & "        resto = PO_TirarPrefixo(resto)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    corpo = resto" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function PO_TirarPrefixo(ByVal s As String) As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim chars As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    chars = "" -:."" & ChrW(8211) & ChrW(8212) & vbTab" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Do While Len(s) > 0" & vbCrLf
    s = s & "        If InStr(1, chars, Left$(s, 1)) > 0 Then" & vbCrLf
    s = s & "            s = Mid$(s, 2)" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            Exit Do" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    PO_TirarPrefixo = s" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub PO_AddEntrada(ByVal col As Collection, ByVal d As Variant, _" & vbCrLf
    s = s & "                          ByVal rotulo As String, ByVal texto As String)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    texto = PO_TrimFim(texto)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If texto = """" And rotulo = """" Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    col.Add Array(d, rotulo, texto)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' ATA EM WORD" & vbCrLf
    CodBase14 = s
End Function

Private Function CodBase15() As String
    Dim s As String
    s = s & "'   lista = array (1..qtd) com os " & ChrW(237) & "ndices das linhas (ListRows)" & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "'   A ata traz, de cada obra, TUDO o que est" & ChrW(225) & " na coluna " & ChrW(218) & "LTIMA" & vbCrLf
    s = s & "'   ATUALIZA" & ChrW(199) & "" & ChrW(195) & "O (o texto inteiro da c" & ChrW(233) & "lula, com todas as linhas) e nada" & vbCrLf
    s = s & "'   do hist" & ChrW(243) & "rico. Obra com a " & ChrW(218) & "LTIMO REGISTRO em branco vai para a" & vbCrLf
    s = s & "'   lista ""Projetos sem atualiza" & ChrW(231) & "" & ChrW(227) & "o""." & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "Public Sub PO_GerarAta(ByVal tb As ListObject, ByVal lista As Variant, ByVal qtd As Long)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim i As Long, g As Long" & vbCrLf
    s = s & "    Dim lr As ListRow" & vbCrLf
    s = s & "    Dim txt() As String          'texto da " & ChrW(250) & "ltima atualiza" & ChrW(231) & "" & ChrW(227) & "o (sem a data inicial)" & vbCrLf
    s = s & "    Dim rot() As String          'data/hora inicial, quando houver (""dd/mm/aaaa hh:mm"")" & vbCrLf
    s = s & "    Dim grupos As New Collection" & vbCrLf
    s = s & "    Dim nomeGrupo As String" & vbCrLf
    s = s & "    Dim existe As Boolean" & vbCrLf
    s = s & "    Dim semAtual As String" & vbCrLf
    s = s & "    Dim qtdCom As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim wd As Object, doc As Object, r As Object" & vbCrLf
    s = s & "    Dim linhaInfo As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If qtd < 1 Then" & vbCrLf
    s = s & "        MsgBox ""Nenhuma obra na lista do painel."", vbExclamation, ""Gerar ata""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If PO_ColIdx(tb, PO_COL_ULTIMA) = 0 Then" & vbCrLf
    s = s & "        MsgBox ""A coluna '"" & PO_COL_ULTIMA & ""' n" & ChrW(227) & "o foi encontrada na tabela '"" & PO_TABELA & ""'."", _" & vbCrLf
    s = s & "               vbExclamation, ""Gerar ata""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    '--------------------------------------------- l" & ChrW(234) & " a " & ChrW(250) & "ltima atualiza" & ChrW(231) & "" & ChrW(227) & "o" & vbCrLf
    s = s & "    ReDim txt(1 To qtd)" & vbCrLf
    s = s & "    ReDim rot(1 To qtd)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 1 To qtd" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        Set lr = tb.ListRows(lista(i))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        PO_UltimaParaAta PO_TextoCampo(tb, lr, PO_COL_ULTIMA), rot(i), txt(i)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If txt(i) <> """" Or rot(i) <> """" Then qtdCom = qtdCom + 1" & vbCrLf
    s = s & "" & vbCrLf
    CodBase15 = s
End Function

Private Function CodBase16() As String
    Dim s As String
    s = s & "        'grupos (TIPO DE OBRA) na ordem em que aparecem" & vbCrLf
    s = s & "        nomeGrupo = PO_GrupoAta(tb, lr)" & vbCrLf
    s = s & "        existe = False" & vbCrLf
    s = s & "        For g = 1 To grupos.Count" & vbCrLf
    s = s & "            If grupos(g) = nomeGrupo Then existe = True: Exit For" & vbCrLf
    s = s & "        Next g" & vbCrLf
    s = s & "        If Not existe Then grupos.Add nomeGrupo" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If qtdCom = 0 Then" & vbCrLf
    s = s & "        MsgBox ""Nenhuma obra da lista tem texto na coluna '"" & PO_COL_ULTIMA & ""'."", vbInformation, ""Gerar ata""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    '--------------------------------------------- Word" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Set wd = CreateObject(""Word.Application"")" & vbCrLf
    s = s & "    On Error GoTo Falha" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If wd Is Nothing Then" & vbCrLf
    s = s & "        MsgBox ""N" & ChrW(227) & "o foi poss" & ChrW(237) & "vel abrir o Microsoft Word."", vbCritical, ""Gerar ata""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Application.StatusBar = ""Gerando ata no Word...""" & vbCrLf
    s = s & "    Application.Cursor = xlWait" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set doc = wd.Documents.Add" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    With doc.PageSetup" & vbCrLf
    s = s & "        .TopMargin = 56.7" & vbCrLf
    s = s & "        .BottomMargin = 56.7" & vbCrLf
    s = s & "        .LeftMargin = 56.7" & vbCrLf
    s = s & "        .RightMargin = 56.7" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    doc.Styles(-1).Font.Name = ""Calibri""     'Normal" & vbCrLf
    s = s & "    doc.Styles(-1).Font.Size = 10.5" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    PO_WAdd doc, ""ATA DE REUNI" & ChrW(195) & "O " & ChrW(8211) & " FOLLOW UP DE OBRAS"", -63               'T" & ChrW(237) & "tulo" & vbCrLf
    s = s & "    PO_WAdd doc, ""Data: "" & Format$(Date, ""dd\/mm\/yyyy"") & ""     Hor" & ChrW(225) & "rio: "" & Format$(Time, ""hh:mm""), -1" & vbCrLf
    s = s & "    PO_WAdd doc, ""Base: " & ChrW(250) & "ltima atualiza" & ChrW(231) & "" & ChrW(227) & "o registrada de cada projeto"", -1" & vbCrLf
    s = s & "    PO_WAdd doc, ""Projetos com atualiza" & ChrW(231) & "" & ChrW(227) & "o: "" & qtdCom & "" de "" & qtd, -1" & vbCrLf
    s = s & "    PO_WAdd doc, ""Participantes: "", -1" & vbCrLf
    CodBase16 = s
End Function

Private Function CodBase17() As String
    Dim s As String
    s = s & "    PO_WAdd doc, """", -1" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For g = 1 To grupos.Count" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        's" & ChrW(243) & " cria o t" & ChrW(237) & "tulo do grupo se houver obra com conte" & ChrW(250) & "do" & vbCrLf
    s = s & "        existe = False" & vbCrLf
    s = s & "        For i = 1 To qtd" & vbCrLf
    s = s & "            If txt(i) <> """" Or rot(i) <> """" Then" & vbCrLf
    s = s & "                If PO_GrupoAta(tb, tb.ListRows(lista(i))) = grupos(g) Then existe = True: Exit For" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If existe Then" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            PO_WAdd doc, grupos(g), -2                                       'T" & ChrW(237) & "tulo 1" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            For i = 1 To qtd" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "                Set lr = tb.ListRows(lista(i))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "                If (txt(i) <> """" Or rot(i) <> """") And PO_GrupoAta(tb, lr) = grupos(g) Then" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "                    PO_WAdd doc, PO_TextoCampo(tb, lr, PO_COL_DESCRICAO) & _" & vbCrLf
    s = s & "                                 ""  (C" & ChrW(211) & "DIGO "" & PO_TextoCampo(tb, lr, PO_COL_CODIGO) & "")"", -3   'T" & ChrW(237) & "tulo 2" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "                    linhaInfo = PO_LinhaInfoAta(tb, lr)" & vbCrLf
    s = s & "                    If linhaInfo <> """" Then" & vbCrLf
    s = s & "                        Set r = PO_WAdd(doc, linhaInfo, -1)" & vbCrLf
    s = s & "                        r.Font.Italic = True" & vbCrLf
    s = s & "                        r.Font.Size = 9" & vbCrLf
    s = s & "                        r.Font.Color = 8421504" & vbCrLf
    s = s & "                    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "                    'um " & ChrW(250) & "nico item com o texto inteiro (as quebras de linha s" & ChrW(227) & "o mantidas)" & vbCrLf
    s = s & "                    If rot(i) <> """" And txt(i) <> """" Then" & vbCrLf
    s = s & "                        Set r = PO_WAdd(doc, rot(i) & "" "" & ChrW(8211) & "" "" & PO_ParaWord(txt(i)), -49)" & vbCrLf
    s = s & "                        doc.Range(r.Start, r.Start + Len(rot(i))).Font.Bold = True" & vbCrLf
    s = s & "                    ElseIf rot(i) <> """" Then" & vbCrLf
    s = s & "                        Set r = PO_WAdd(doc, rot(i), -49)" & vbCrLf
    s = s & "                        r.Font.Bold = True" & vbCrLf
    s = s & "                    Else" & vbCrLf
    s = s & "                        PO_WAdd doc, PO_ParaWord(txt(i)), -49" & vbCrLf
    s = s & "                    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "                End If" & vbCrLf
    CodBase17 = s
End Function

Private Function CodBase18() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "            Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Next g" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    'Obras sem texto na " & ChrW(250) & "ltima atualiza" & ChrW(231) & "" & ChrW(227) & "o" & vbCrLf
    s = s & "    semAtual = """"" & vbCrLf
    s = s & "    For i = 1 To qtd" & vbCrLf
    s = s & "        If txt(i) = """" And rot(i) = """" Then" & vbCrLf
    s = s & "            Set lr = tb.ListRows(lista(i))" & vbCrLf
    s = s & "            If semAtual <> """" Then semAtual = semAtual & ""; """ & vbCrLf
    s = s & "            semAtual = semAtual & PO_TextoCampo(tb, lr, PO_COL_DESCRICAO)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If semAtual <> """" Then" & vbCrLf
    s = s & "        PO_WAdd doc, ""Projetos sem atualiza" & ChrW(231) & "" & ChrW(227) & "o"", -2" & vbCrLf
    s = s & "        PO_WAdd doc, semAtual, -1" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    doc.Paragraphs(doc.Paragraphs.Count).Style = doc.Styles(-1)" & vbCrLf
    s = s & "    doc.Content.LanguageID = 1046      'Portugu" & ChrW(234) & "s (Brasil)" & vbCrLf
    s = s & "    On Error GoTo Falha" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    wd.Visible = True" & vbCrLf
    s = s & "    doc.Activate" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    wd.Activate" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Application.StatusBar = False" & vbCrLf
    s = s & "    Application.Cursor = xlDefault" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Falha:" & vbCrLf
    s = s & "    Application.StatusBar = False" & vbCrLf
    s = s & "    Application.Cursor = xlDefault" & vbCrLf
    s = s & "    MsgBox ""Erro ao gerar a ata:"" & vbCrLf & Err.Description, vbCritical, ""Gerar ata""" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    If Not wd Is Nothing Then wd.Visible = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodBase18 = s
End Function

Private Function CodBase19() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "'Prepara o texto da " & ChrW(218) & "LTIMO REGISTRO para a ata: separa a data/hora do" & vbCrLf
    s = s & "'in" & ChrW(237) & "cio (se houver) para sair em negrito e devolve TODO o restante do" & vbCrLf
    s = s & "'texto, com todas as linhas, exatamente como est" & ChrW(225) & " na c" & ChrW(233) & "lula." & vbCrLf
    s = s & "Private Sub PO_UltimaParaAta(ByVal bruto As String, ByRef rotulo As String, ByRef corpo As String)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim s As String" & vbCrLf
    s = s & "    Dim primeira As String, resto As String" & vbCrLf
    s = s & "    Dim p As Long" & vbCrLf
    s = s & "    Dim d As Variant" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    rotulo = """"" & vbCrLf
    s = s & "    corpo = """"" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    s = Replace(bruto, vbCrLf, vbLf)" & vbCrLf
    s = s & "    s = Replace(s, vbCr, vbLf)" & vbCrLf
    s = s & "    s = PO_TrimFim(s)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    'tira linhas em branco do in" & ChrW(237) & "cio" & vbCrLf
    s = s & "    Do While Left$(s, 1) = vbLf Or Left$(s, 1) = "" """ & vbCrLf
    s = s & "        s = Mid$(s, 2)" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If s = """" Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If s Like ""##/##/####*"" Then" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        p = InStr(s, vbLf)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If p > 0 Then" & vbCrLf
    s = s & "            primeira = Left$(s, p - 1)" & vbCrLf
    s = s & "            resto = Mid$(s, p + 1)" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            primeira = s" & vbCrLf
    s = s & "            resto = """"" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        PO_SepararCabecalho Trim$(primeira), d, rotulo, corpo" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If resto <> """" Then" & vbCrLf
    s = s & "            If corpo <> """" Then" & vbCrLf
    s = s & "                corpo = corpo & vbLf & resto" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                corpo = resto" & vbCrLf
    s = s & "            End If" & vbCrLf
    CodBase19 = s
End Function

Private Function CodBase20() As String
    Dim s As String
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        corpo = s" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function PO_GrupoAta(ByVal tb As ListObject, ByVal lr As ListRow) As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim s As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    s = UCase$(PO_TextoCampo(tb, lr, PO_COL_MODALIDADE))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Select Case PO_SemAcento(s)" & vbCrLf
    s = s & "        Case ""CONSTRUCAO"": s = ""CONSTRU" & ChrW(199) & "" & ChrW(195) & "O""" & vbCrLf
    s = s & "        Case ""AMPLIACAO"": s = ""AMPLIA" & ChrW(199) & "" & ChrW(195) & "O""" & vbCrLf
    s = s & "        Case """": s = ""OUTRAS OBRAS""" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    PO_GrupoAta = s" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function PO_LinhaInfoAta(ByVal tb As ListObject, ByVal lr As ListRow) As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim partes As String" & vbCrLf
    s = s & "    Dim v As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    v = PO_TextoCampo(tb, lr, PO_COL_CIDADE)" & vbCrLf
    s = s & "    If v <> """" Then" & vbCrLf
    s = s & "        If PO_TextoCampo(tb, lr, PO_COL_UF) <> """" Then v = v & ""/"" & PO_TextoCampo(tb, lr, PO_COL_UF)" & vbCrLf
    s = s & "        partes = partes & ""  |  "" & v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    v = PO_TextoCampo(tb, lr, PO_COL_FASE)" & vbCrLf
    s = s & "    If v <> """" Then partes = partes & ""  |  Etapa: "" & v" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    v = PO_TextoCampo(tb, lr, PO_COL_STATUS)" & vbCrLf
    s = s & "    If v <> """" Then partes = partes & ""  |  Situa" & ChrW(231) & "" & ChrW(227) & "o: "" & v" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    v = PO_TextoCampo(tb, lr, PO_COL_ENTREGA)" & vbCrLf
    s = s & "    If v <> """" Then partes = partes & ""  |  Entrega prevista: "" & v" & vbCrLf
    CodBase20 = s
End Function

Private Function CodBase21() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "    If Len(partes) > 5 Then partes = Mid$(partes, 6)" & vbCrLf
    s = s & "    PO_LinhaInfoAta = partes" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Acrescenta um par" & ChrW(225) & "grafo no fim do documento e devolve o Range dele" & vbCrLf
    s = s & "Private Function PO_WAdd(ByVal doc As Object, ByVal texto As String, ByVal estilo As Long) As Object" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim r As Object" & vbCrLf
    s = s & "    Dim pos As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    pos = doc.Content.End - 1" & vbCrLf
    s = s & "    Set r = doc.Range(pos, pos)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    r.InsertAfter texto" & vbCrLf
    s = s & "    r.Style = doc.Styles(estilo)" & vbCrLf
    s = s & "    r.Font.Reset" & vbCrLf
    s = s & "    r.InsertParagraphAfter" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set PO_WAdd = doc.Range(r.Start, r.Start + Len(texto))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function PO_ParaWord(ByVal s As String) As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    s = Replace(s, vbCrLf, vbLf)" & vbCrLf
    s = s & "    s = Replace(s, vbCr, vbLf)" & vbCrLf
    s = s & "    PO_ParaWord = Replace(s, vbLf, Chr(11))     'quebra de linha manual" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' PAINEL DE OBRAS (tabela tb_Obras, aba OBRAS)" & vbCrLf
    s = s & "'   Abrir o painel ......: AbrirPainelObras" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "Public Sub AbrirPainelObras()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim tb As ListObject" & vbCrLf
    s = s & "    Dim idx As Long" & vbCrLf
    s = s & "    Dim f As Object" & vbCrLf
    s = s & "    Dim painel As Object" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set tb = PO_TabelaPorNome(""tb_Obras"")" & vbCrLf
    CodBase21 = s
End Function

Private Function CodBase22() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "    If tb Is Nothing Then" & vbCrLf
    s = s & "        MsgBox ""A tabela 'tb_Obras' n" & ChrW(227) & "o foi encontrada."", vbExclamation, ""Painel de Obras""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    'Se a c" & ChrW(233) & "lula ativa estiver dentro da tabela, abre naquela obra" & vbCrLf
    s = s & "    idx = 0" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    If ActiveSheet.Name = tb.Parent.Name Then" & vbCrLf
    s = s & "        If Not Intersect(ActiveCell, tb.DataBodyRange) Is Nothing Then" & vbCrLf
    s = s & "            idx = ActiveCell.row - tb.DataBodyRange.row + 1" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    'O formul" & ChrW(225) & "rio " & ChrW(233) & " localizado pelo nome, para este m" & ChrW(243) & "dulo continuar" & vbCrLf
    s = s & "    'compilando mesmo antes de o frmPainelObras ser criado." & vbCrLf
    s = s & "    For Each f In VBA.UserForms" & vbCrLf
    s = s & "        If f.Name = ""frmPainelObras"" Then Set painel = f" & vbCrLf
    s = s & "    Next f" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If painel Is Nothing Then" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        On Error Resume Next" & vbCrLf
    s = s & "        Set painel = VBA.UserForms.Add(""frmPainelObras"")" & vbCrLf
    s = s & "        On Error GoTo 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If painel Is Nothing Then" & vbCrLf
    s = s & "            MsgBox ""O formul" & ChrW(225) & "rio 'frmPainelObras' ainda n" & ChrW(227) & "o existe neste arquivo."" & vbCrLf & vbCrLf & _" & vbCrLf
    s = s & "                   ""Crie um UserForm com esse nome e cole nele o c" & ChrW(243) & "digo do painel."", _" & vbCrLf
    s = s & "                   vbExclamation, ""Painel de Obras""" & vbCrLf
    s = s & "            Exit Sub" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    painel.IrParaLinha idx" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not painel.Visible Then painel.Show vbModeless" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodBase22 = s
End Function

Private Function CodigoObra() As String
    Dim s As String
    s = s & CodObra1()
    s = s & CodObra2()
    s = s & CodObra3()
    s = s & CodObra4()
    s = s & CodObra5()
    CodigoObra = s
End Function

Private Function CodObra1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' PAINEL - CLASSE OBRA" & vbCrLf
    s = s & "' Representa UMA linha da tb_Obras." & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "'   Dim obra As New clsObra" & vbCrLf
    s = s & "'   obra.Carregar tb, 7" & vbCrLf
    s = s & "'   obra.Nome / obra.CodObra / obra.Texto(""CIDADE"")" & vbCrLf
    s = s & "'   obra.AdicionarObservacao ""texto""      -> grava HIST" & ChrW(211) & "RICO + " & ChrW(218) & "LTIMO REGISTRO" & vbCrLf
    s = s & "'   obra.SubstituirHistorico ""texto""      -> corre" & ChrW(231) & "" & ChrW(227) & "o do hist" & ChrW(243) & "rico" & vbCrLf
    s = s & "'   obra.GravarData idxColuna, ""15/11/2026""" & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "' Toda grava" & ChrW(231) & "" & ChrW(227) & "o passa por PO_Gravar (log na tb_Alteracoes)." & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mTb As ListObject" & vbCrLf
    s = s & "Private mIdx As Long" & vbCrLf
    s = s & "Private mLr As ListRow" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'------------------------------------------------------------ carga" & vbCrLf
    s = s & "Public Sub Carregar(ByVal tabelaOrigem As ListObject, ByVal indiceLinha As Long)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set mTb = tabelaOrigem" & vbCrLf
    s = s & "    mIdx = indiceLinha" & vbCrLf
    s = s & "    Set mLr = mTb.ListRows(mIdx)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Tabela() As ListObject" & vbCrLf
    s = s & "    Set Tabela = mTb" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Indice() As Long" & vbCrLf
    s = s & "    Indice = mIdx" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Linha() As ListRow" & vbCrLf
    s = s & "    Set Linha = mLr" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'------------------------------------------------------------ leitura" & vbCrLf
    s = s & "Public Function valor(ByVal nomeColuna As String) As Variant" & vbCrLf
    CodObra1 = s
End Function

Private Function CodObra2() As String
    Dim s As String
    s = s & "    valor = PO_Valor(mTb, mLr, nomeColuna)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function texto(ByVal nomeColuna As String) As String" & vbCrLf
    s = s & "    texto = PO_TextoCampo(mTb, mLr, nomeColuna)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function TextoColuna(ByVal idxColuna As Long) As String" & vbCrLf
    s = s & "    TextoColuna = PO_TextoValor(mLr.Range.Cells(1, idxColuna).value)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get nome() As String" & vbCrLf
    s = s & "    nome = texto(PO_COL_DESCRICAO)" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get CodObra() As String" & vbCrLf
    s = s & "    CodObra = texto(PO_COL_CODIGO)" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get rotulo() As String" & vbCrLf
    s = s & "    rotulo = CodObra & "" - "" & nome" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Historico() As String" & vbCrLf
    s = s & "    Historico = texto(PO_COL_OBS)" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get UltimaObservacao() As String" & vbCrLf
    s = s & "    UltimaObservacao = texto(PO_COL_ULTIMA)" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Visivel() As Boolean" & vbCrLf
    s = s & "    Visivel = Not mLr.Range.EntireRow.Hidden" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'""C" & ChrW(211) & "DIGO 2801 | CLIENTE ATLAS | CAMPINAS/SP | CONSTRU" & ChrW(199) & "" & ChrW(195) & "O | Etapa: ... | Situa" & ChrW(231) & "" & ChrW(227) & "o: ...""" & vbCrLf
    s = s & "Public Property Get Resumo() As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim s As String, v As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    s = ""C" & ChrW(211) & "DIGO "" & CodObra" & vbCrLf
    s = s & "    v = texto(""CLIENTE""): If v <> """" Then s = s & ""   |   "" & v" & vbCrLf
    s = s & "    v = texto(PO_COL_CIDADE)" & vbCrLf
    s = s & "    If v <> """" Then" & vbCrLf
    s = s & "        s = s & ""   |   "" & v" & vbCrLf
    CodObra2 = s
End Function

Private Function CodObra3() As String
    Dim s As String
    s = s & "        If texto(PO_COL_UF) <> """" Then s = s & ""/"" & texto(PO_COL_UF)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    v = texto(PO_COL_MODALIDADE): If v <> """" Then s = s & ""   |   "" & v" & vbCrLf
    s = s & "    v = texto(PO_COL_FASE):       If v <> """" Then s = s & ""   |   Etapa: "" & v" & vbCrLf
    s = s & "    v = texto(PO_COL_STATUS):     If v <> """" Then s = s & ""   |   Situa" & ChrW(231) & "" & ChrW(227) & "o: "" & v" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Resumo = s" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Entradas do hist" & ChrW(243) & "rico: cada item = Array(data|Empty, r" & ChrW(243) & "tulo, texto)" & vbCrLf
    s = s & "Public Function Observacoes() As Collection" & vbCrLf
    s = s & "    Set Observacoes = PO_ParseObs(Historico)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'------------------------------------------------------------ grava" & ChrW(231) & "" & ChrW(227) & "o" & vbCrLf
    s = s & "'Retorno: 1 alterou | 0 igual | -1 falhou" & vbCrLf
    s = s & "Public Function GravarColuna(ByVal idxColuna As Long, ByVal novo As Variant) As Long" & vbCrLf
    s = s & "    GravarColuna = PO_Gravar(mTb, mLr, idxColuna, novo)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Texto dd/mm/aaaa (ou vazio para limpar). -2 = data inv" & ChrW(225) & "lida" & vbCrLf
    s = s & "Public Function GravarData(ByVal idxColuna As Long, ByVal textoData As String) As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim d As Date, vz As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not PO_TextoParaData(textoData, d, vz) Then" & vbCrLf
    s = s & "        GravarData = -2" & vbCrLf
    s = s & "    ElseIf vz Then" & vbCrLf
    s = s & "        GravarData = GravarColuna(idxColuna, Empty)" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        GravarData = GravarColuna(idxColuna, d)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Acrescenta ""dd/mm/aaaa hh:mm - texto"" ao hist" & ChrW(243) & "rico e " & ChrW(224) & " " & ChrW(250) & "ltima atualiza" & ChrW(231) & "" & ChrW(227) & "o." & vbCrLf
    s = s & "'historicoBase: se informado, " & ChrW(233) & " usado no lugar do hist" & ChrW(243) & "rico atual (edi" & ChrW(231) & "" & ChrW(227) & "o manual)." & vbCrLf
    s = s & "Public Function AdicionarObservacao(ByVal textoNovo As String, _" & vbCrLf
    s = s & "                                    Optional ByVal historicoBase As Variant) As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim idxObs As Long, idxUlt As Long" & vbCrLf
    s = s & "    Dim base As String, entrada As String" & vbCrLf
    s = s & "" & vbCrLf
    CodObra3 = s
End Function

Private Function CodObra4() As String
    Dim s As String
    s = s & "    idxObs = PO_ColIdx(mTb, PO_COL_OBS)" & vbCrLf
    s = s & "    idxUlt = PO_ColIdx(mTb, PO_COL_ULTIMA)" & vbCrLf
    s = s & "    If idxObs = 0 Or idxUlt = 0 Then Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    textoNovo = Trim$(PO_TrimFim(PO_ParaCaixa(textoNovo)))" & vbCrLf
    s = s & "    If textoNovo = """" Then Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If IsMissing(historicoBase) Then" & vbCrLf
    s = s & "        base = Historico" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        base = CStr(historicoBase)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    entrada = Format$(Now, ""dd\/mm\/yyyy hh:mm"") & "" - "" & textoNovo" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    base = PO_TrimFim(base)" & vbCrLf
    s = s & "    If base <> """" Then base = base & vbCrLf & String(PO_TAM_SEPARADOR, ""-"") & vbCrLf" & vbCrLf
    s = s & "    base = base & entrada" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If GravarColuna(idxObs, base) < 0 Then Exit Function" & vbCrLf
    s = s & "    If GravarColuna(idxUlt, entrada) < 0 Then Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    AdicionarObservacao = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Substitui o hist" & ChrW(243) & "rico inteiro (corre" & ChrW(231) & "" & ChrW(245) & "es feitas em ""Editar hist" & ChrW(243) & "rico"")" & vbCrLf
    s = s & "Public Function SubstituirHistorico(ByVal NovoTexto As String) As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim idxObs As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    idxObs = PO_ColIdx(mTb, PO_COL_OBS)" & vbCrLf
    s = s & "    If idxObs = 0 Then Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SubstituirHistorico = (GravarColuna(idxObs, PO_TrimFim(PO_ParaCaixa(NovoTexto))) >= 0)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Seleciona a linha da obra na planilha" & vbCrLf
    s = s & "Public Sub IrParaPlanilha()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    mTb.Parent.Activate" & vbCrLf
    s = s & "    mLr.Range.Cells(1, 1).Select" & vbCrLf
    s = s & "    ActiveWindow.ScrollRow = mLr.Range.row" & vbCrLf
    CodObra4 = s
End Function

Private Function CodObra5() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodObra5 = s
End Function

Private Function CodigoCampoData() As String
    Dim s As String
    s = s & CodCampoData1()
    s = s & CodCampoData2()
    CodigoCampoData = s
End Function

Private Function CodCampoData1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' PAINEL - comportamento das caixas de data" & vbCrLf
    s = s & "'  - aceita s" & ChrW(243) & " n" & ChrW(250) & "meros e ""/""  (barras inseridas automaticamente)" & vbCrLf
    s = s & "'  - fica vermelha se a data estiver inv" & ChrW(225) & "lida" & vbCrLf
    s = s & "'  - duplo clique = data de hoje" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public WithEvents Caixa As MSForms.TextBox" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Caixa_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim k As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    k = KeyAscii" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If k < 32 Then Exit Sub          'teclas de controle (Backspace, Ctrl+V...)" & vbCrLf
    s = s & "    If k = 47 Then Exit Sub          '""/""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If k >= 48 And k <= 57 Then" & vbCrLf
    s = s & "        With Caixa" & vbCrLf
    s = s & "            If .SelLength = 0 And .SelStart = Len(.Text) Then" & vbCrLf
    s = s & "                If Len(.Text) = 2 Or Len(.Text) = 5 Then" & vbCrLf
    s = s & "                    .Text = .Text & ""/""" & vbCrLf
    s = s & "                    .SelStart = Len(.Text)" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End With" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    KeyAscii = 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Caixa_Change()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim d As Date" & vbCrLf
    s = s & "    Dim vz As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If PO_TextoParaData(Caixa.Text, d, vz) Then" & vbCrLf
    s = s & "        Caixa.BackColor = vbWhite" & vbCrLf
    s = s & "    ElseIf Len(Caixa.Text) >= 8 Then" & vbCrLf
    s = s & "        Caixa.BackColor = PO_COR_ERRO" & vbCrLf
    CodCampoData1 = s
End Function

Private Function CodCampoData2() As String
    Dim s As String
    s = s & "    Else" & vbCrLf
    s = s & "        Caixa.BackColor = vbWhite" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Caixa_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Caixa.Text = Format$(Date, ""dd\/mm\/yyyy"")" & vbCrLf
    s = s & "    Cancel = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodCampoData2 = s
End Function

Private Function CodigoPainel() As String
    Dim s As String
    s = s & CodPainel1()
    s = s & CodPainel2()
    s = s & CodPainel3()
    s = s & CodPainel4()
    s = s & CodPainel5()
    s = s & CodPainel6()
    s = s & CodPainel7()
    s = s & CodPainel8()
    s = s & CodPainel9()
    s = s & CodPainel10()
    s = s & CodPainel11()
    s = s & CodPainel12()
    s = s & CodPainel13()
    s = s & CodPainel14()
    s = s & CodPainel15()
    s = s & CodPainel16()
    s = s & CodPainel17()
    s = s & CodPainel18()
    s = s & CodPainel19()
    s = s & CodPainel20()
    s = s & CodPainel21()
    s = s & CodPainel22()
    s = s & CodPainel23()
    s = s & CodPainel24()
    s = s & CodPainel25()
    s = s & CodPainel26()
    s = s & CodPainel27()
    s = s & CodPainel28()
    s = s & CodPainel29()
    s = s & CodPainel30()
    s = s & CodPainel31()
    s = s & CodPainel32()
    s = s & CodPainel33()
    s = s & CodPainel34()
    s = s & CodPainel35()
    s = s & CodPainel36()
    s = s & CodPainel37()
    s = s & CodPainel38()
    s = s & CodPainel39()
    s = s & CodPainel40()
    s = s & CodPainel41()
    s = s & CodPainel42()
    CodigoPainel = s
End Function

Private Function CodPainel1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'---------------- API: tamanho real da tela (para ajustar o painel)" & vbCrLf
    s = s & "Private Type RECT_TELA" & vbCrLf
    s = s & "    Esq As Long" & vbCrLf
    s = s & "    Topo As Long" & vbCrLf
    s = s & "    Direita As Long" & vbCrLf
    s = s & "    Baixo As Long" & vbCrLf
    s = s & "End Type" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "#If VBA7 Then" & vbCrLf
    s = s & "    Private Declare PtrSafe Function SystemParametersInfoA Lib ""user32"" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long" & vbCrLf
    s = s & "    Private Declare PtrSafe Function GetDC Lib ""user32"" (ByVal hwnd As LongPtr) As LongPtr" & vbCrLf
    s = s & "    Private Declare PtrSafe Function ReleaseDC Lib ""user32"" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long" & vbCrLf
    s = s & "    Private Declare PtrSafe Function GetDeviceCaps Lib ""gdi32"" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long" & vbCrLf
    s = s & "    Private Declare PtrSafe Function CreateFontA Lib ""gdi32"" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As LongPtr" & vbCrLf
    s = s & "    Private Declare PtrSafe Function SelectObject Lib ""gdi32"" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr" & vbCrLf
    s = s & "    Private Declare PtrSafe Function DeleteObject Lib ""gdi32"" (ByVal hObject As LongPtr) As Long" & vbCrLf
    s = s & "    Private Declare PtrSafe Function GetTextFaceA Lib ""gdi32"" (ByVal hdc As LongPtr, ByVal nCount As Long, ByVal lpFaceName As String) As Long" & vbCrLf
    s = s & "#Else" & vbCrLf
    s = s & "    Private Declare Function SystemParametersInfoA Lib ""user32"" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long" & vbCrLf
    s = s & "    Private Declare Function GetDC Lib ""user32"" (ByVal hwnd As Long) As Long" & vbCrLf
    s = s & "    Private Declare Function ReleaseDC Lib ""user32"" (ByVal hwnd As Long, ByVal hdc As Long) As Long" & vbCrLf
    s = s & "    Private Declare Function GetDeviceCaps Lib ""gdi32"" (ByVal hdc As Long, ByVal nIndex As Long) As Long" & vbCrLf
    s = s & "    Private Declare Function CreateFontA Lib ""gdi32"" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long" & vbCrLf
    s = s & "    Private Declare Function SelectObject Lib ""gdi32"" (ByVal hdc As Long, ByVal hObject As Long) As Long" & vbCrLf
    s = s & "    Private Declare Function DeleteObject Lib ""gdi32"" (ByVal hObject As Long) As Long" & vbCrLf
    s = s & "    Private Declare Function GetTextFaceA Lib ""gdi32"" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFaceName As String) As Long" & vbCrLf
    s = s & "#End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' PAINEL DE OBRAS - FORMUL" & ChrW(193) & "RIO PRINCIPAL" & vbCrLf
    s = s & "' Todos os controles s" & ChrW(227) & "o criados por c" & ChrW(243) & "digo (n" & ChrW(227) & "o precisa desenhar nada)." & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "' Trabalha sobre a tabela tb_Obras (aba OBRAS) e usa as rotinas do m" & ChrW(243) & "dulo" & vbCrLf
    s = s & "' mPainelBase e das classes clsObra e clsCampoData." & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "'   SALVAR ............ grava as datas alteradas e a nova observa" & ChrW(231) & "" & ChrW(227) & "o" & vbCrLf
    s = s & "'   GERAR ATA ......... Word, com o texto inteiro da " & ChrW(218) & "LTIMO REGISTRO" & vbCrLf
    s = s & "'   ATUALIZAR DADOS ... atualiza as consultas e recarrega o painel" & vbCrLf
    s = s & "'   NOVA OBRA ......... inclui uma linha pelo C" & ChrW(211) & "DIGO; as f" & ChrW(243) & "rmulas da" & vbCrLf
    s = s & "'                       tabela puxam os demais dados" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const TB_OBRAS As String = ""tb_Obras""" & vbCrLf
    CodPainel1 = s
End Function

Private Function CodPainel2() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "'---------------- IDENTIDADE VISUAL (paleta institucional, sem o rosa)" & vbCrLf
    s = s & "Private Const COR_ESCURO As Long = 3877150       'RGB(30, 41, 59)" & vbCrLf
    s = s & "Private Const COR_AZUL As Long = 8950797        'RGB(13, 148, 136)" & vbCrLf
    s = s & "Private Const COR_AZUL_CLARO As Long = 15858636  'RGB(204, 251, 241)" & vbCrLf
    s = s & "Private Const COR_AMARELO As Long = 2408443      'RGB(251, 191, 36)" & vbCrLf
    s = s & "Private Const COR_VERMELHO As Long = 2500316     'RGB(220, 38, 38)" & vbCrLf
    s = s & "Private Const COR_CINZA As Long = 9139300        'RGB(100, 116, 139)" & vbCrLf
    s = s & "'tons neutros de apoio (fundos e bordas)" & vbCrLf
    s = s & "Private Const COR_FUNDO As Long = 16579320       'RGB(248, 250, 252)" & vbCrLf
    s = s & "Private Const COR_LEITURA As Long = 16381425     'RGB(241, 245, 249)" & vbCrLf
    s = s & "Private Const COR_EDICAO As Long = 15465471      'RGB(255, 251, 235)" & vbCrLf
    s = s & "Private Const COR_BORDA As Long = 14800331       'RGB(203, 213, 225)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Fontes da marca. Se n" & ChrW(227) & "o estiverem instaladas no computador, o painel" & vbCrLf
    s = s & "'usa a fonte reserva automaticamente." & vbCrLf
    s = s & "Private Const FONTE_MARCA As String = ""Segoe UI""" & vbCrLf
    s = s & "Private Const FONTE_MARCA_TITULO As String = ""Segoe UI Black""" & vbCrLf
    s = s & "Private Const FONTE_RESERVA As String = ""Segoe UI""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mFonte As String             'fonte de textos e subt" & ChrW(237) & "tulos" & vbCrLf
    s = s & "Private mFonteTitulo As String       'fonte do nome da obra" & vbCrLf
    s = s & "Private mTituloNegrito As Boolean    'True quando o t" & ChrW(237) & "tulo usa a fonte comum" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const LARG As Single = 1092   '880 + 212 do calend" & ChrW(225) & "rio" & vbCrLf
    s = s & "Private Const ALT As Single = 522   '496 + 26 da faixa de avisos" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private tb As ListObject" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mMapa() As Long          'posi" & ChrW(231) & "" & ChrW(227) & "o na lista -> " & ChrW(237) & "ndice ListRow" & vbCrLf
    s = s & "Private mQtd As Long" & vbCrLf
    s = s & "Private mAtual As Long           '" & ChrW(237) & "ndice ListRow da obra aberta" & vbCrLf
    s = s & "Private mObra As clsObra   'objeto da obra aberta" & vbCrLf
    s = s & "Private mCarregando As Boolean" & vbCrLf
    s = s & "Private mEditandoHist As Boolean" & vbCrLf
    s = s & "Private mHistOriginal As String" & vbCrLf
    s = s & "Private mPendenteIr As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Campos da coluna da esquerda (montados em MontarCamposInfo)" & vbCrLf
    s = s & "Private mQtdCampos As Long" & vbCrLf
    s = s & "Private mCampoIdx() As Long          '" & ChrW(237) & "ndice da coluna na tabela" & vbCrLf
    s = s & "Private mCampoRot() As String        'r" & ChrW(243) & "tulo mostrado no painel" & vbCrLf
    s = s & "Private mCampoTipo() As String       '""D"" = data edit" & ChrW(225) & "vel | ""N"" = s" & ChrW(243) & " consulta" & vbCrLf
    s = s & "Private mEditavel() As Boolean       'edit" & ChrW(225) & "vel NESTA obra (sem f" & ChrW(243) & "rmula na c" & ChrW(233) & "lula)" & vbCrLf
    s = s & "Private mOrig() As String            'texto ao carregar a obra" & vbCrLf
    CodPainel2 = s
End Function

Private Function CodPainel3() As String
    Dim s As String
    s = s & "Private mLbl() As MSForms.label" & vbCrLf
    s = s & "Private mBoxEd() As MSForms.TextBox  'caixa de edi" & ChrW(231) & "" & ChrW(227) & "o (datas digitadas)" & vbCrLf
    s = s & "Private mBoxRo() As MSForms.TextBox  'caixa de consulta (vem de f" & ChrW(243) & "rmula)" & vbCrLf
    s = s & "Private mCamposData As Collection    'clsCampoData de cada caixa de edi" & ChrW(231) & "" & ChrW(227) & "o" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private WithEvents txtBusca As MSForms.TextBox" & vbCrLf
    s = s & "Private WithEvents cboObra As MSForms.ComboBox" & vbCrLf
    s = s & "Private WithEvents chkVisiveis As MSForms.CheckBox" & vbCrLf
    s = s & "Private WithEvents btnAnterior As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnProxima As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnEditarHist As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnSalvar As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnAta As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnAtualizar As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnNova As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnIrPlanilha As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnMarcar As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents chkMarcadas As MSForms.CheckBox" & vbCrLf
    s = s & "Private WithEvents lblMcClique As MSForms.Label" & vbCrLf
    s = s & "Private mMarcas As cMarcadores" & vbCrLf
    s = s & "Private mMcIcone As cMarcaIcone" & vbCrLf
    s = s & "Private lblMcFundo As MSForms.Label" & vbCrLf
    s = s & "Private lblMcBarra As MSForms.Label" & vbCrLf
    s = s & "Private lblMcTit As MSForms.Label" & vbCrLf
    s = s & "Private lblMcPrazo As MSForms.Label" & vbCrLf
    s = s & "Private lblMcTexto As MSForms.Label" & vbCrLf
    s = s & "Private lblMcPor As MSForms.Label" & vbCrLf
    s = s & "Private lblMcMedida As MSForms.Label" & vbCrLf
    s = s & "Private mCal As cCalendario" & vbCrLf
    s = s & "Private mFerCampos As Collection" & vbCrLf
    s = s & "Private mFerTag() As MSForms.Label" & vbCrLf
    s = s & "Private mFerQtd As Long" & vbCrLf
    s = s & "Private mFerIdx As Long" & vbCrLf
    s = s & "Private mFerUltimo As Long" & vbCrLf
    s = s & "Private mFerSug1 As Date" & vbCrLf
    s = s & "Private mFerSug2 As Date" & vbCrLf
    s = s & "Private mFerUF As String" & vbCrLf
    s = s & "Private mFerCid As String" & vbCrLf
    s = s & "Private WithEvents mNova As cNovaObraPainel" & vbCrLf
    s = s & "Private fraNova As MSForms.Frame" & vbCrLf
    s = s & "Private WithEvents btnEntrega As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents mEntrega As cEntregaPainel" & vbCrLf
    s = s & "Private fraEntrega As MSForms.Frame" & vbCrLf
    s = s & "Private WithEvents btnFechar As MSForms.CommandButton" & vbCrLf
    s = s & "" & vbCrLf
    CodPainel3 = s
End Function

Private Function CodPainel4() As String
    Dim s As String
    s = s & "Private lblNome As MSForms.label" & vbCrLf
    s = s & "Private lblSub As MSForms.label" & vbCrLf
    s = s & "Private lblContador As MSForms.label" & vbCrLf
    s = s & "Private txtHist As MSForms.TextBox" & vbCrLf
    s = s & "Private txtNova As MSForms.TextBox" & vbCrLf
    s = s & "Private fraInfo As MSForms.Frame" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' INICIALIZA" & ChrW(199) & "" & ChrW(195) & "O" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set tb = PO_TabelaPorNome(TB_OBRAS)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    DefinirFontes" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Me.Caption = ""PAINEL DE OBRAS - DADOS DA OBRA""" & vbCrLf
    s = s & "    Me.BackColor = COR_FUNDO" & vbCrLf
    s = s & "    Me.Width = LARG + 6" & vbCrLf
    s = s & "    Me.Height = ALT + 28" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mCarregando = True" & vbCrLf
    s = s & "    MontarTela" & vbCrLf
    s = s & "    chkVisiveis.value = True          'come" & ChrW(231) & "a pelas obras vis" & ChrW(237) & "veis no filtro da planilha" & vbCrLf
    s = s & "    mCarregando = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not tb Is Nothing Then" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        MontarCamposInfo" & vbCrLf
    s = s & "        FerPrepararCampos" & vbCrLf
    s = s & "        FiltrarLista" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If mQtd = 0 Then" & vbCrLf
    s = s & "            'nenhuma obra vis" & ChrW(237) & "vel no filtro: mostra todas" & vbCrLf
    s = s & "            mCarregando = True" & vbCrLf
    s = s & "            chkVisiveis.value = False" & vbCrLf
    s = s & "            mCarregando = False" & vbCrLf
    s = s & "            FiltrarLista" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If mQtd > 0 Then CarregarObra mMapa(1)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    CodPainel4 = s
End Function

Private Function CodPainel5() As String
    Dim s As String
    s = s & "    AjustarZoom" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Activate()" & vbCrLf
    s = s & "    RolarHistoricoFim" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If Not ConfirmarPendencias() Then Cancel = True" & vbCrLf
    s = s & "    If Cancel = 0 Then FerDesligar" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Chamado pelo m" & ChrW(243) & "dulo: abre o painel numa obra (0 = mant" & ChrW(233) & "m a atual)" & vbCrLf
    s = s & "Public Sub IrParaLinha(ByVal idx As Long)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    If idx < 1 Or idx > tb.ListRows.Count Then Exit Sub" & vbCrLf
    s = s & "    If idx = mAtual Then Exit Sub" & vbCrLf
    s = s & "    If Not ConfirmarPendencias() Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If PosicaoNaLista(idx) = 0 Then" & vbCrLf
    s = s & "        'obra fora do filtro atual: limpa a busca" & vbCrLf
    s = s & "        mCarregando = True" & vbCrLf
    s = s & "        txtBusca.Text = """"" & vbCrLf
    s = s & "        chkVisiveis.value = False" & vbCrLf
    s = s & "        mCarregando = False" & vbCrLf
    s = s & "        FiltrarLista" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    CarregarObra idx" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' CRIA" & ChrW(199) & "" & ChrW(195) & "O DOS CONTROLES" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "Private Function NovoLabel(ByVal pai As Object, ByVal nome As String, ByVal cap As String, _" & vbCrLf
    s = s & "                           ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _" & vbCrLf
    s = s & "                           Optional ByVal tam As Single = 9, Optional ByVal negrito As Boolean = False, _" & vbCrLf
    s = s & "                           Optional ByVal cor As Long = -1) As MSForms.label" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim c As MSForms.label" & vbCrLf
    s = s & "" & vbCrLf
    CodPainel5 = s
End Function

Private Function CodPainel6() As String
    Dim s As String
    s = s & "    Set c = pai.Controls.Add(""Forms.Label.1"", nome, True)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    With c" & vbCrLf
    s = s & "        .Caption = cap" & vbCrLf
    s = s & "        .Left = l: .Top = t: .Width = w: .Height = h" & vbCrLf
    s = s & "        .BackStyle = fmBackStyleTransparent" & vbCrLf
    s = s & "        .Font.Name = mFonte" & vbCrLf
    s = s & "        .Font.Size = tam" & vbCrLf
    s = s & "        .Font.Bold = negrito" & vbCrLf
    s = s & "        If cor <> -1 Then .ForeColor = cor Else .ForeColor = COR_ESCURO" & vbCrLf
    s = s & "        .WordWrap = True" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set NovoLabel = c" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function NovoTexto(ByVal pai As Object, ByVal nome As String, _" & vbCrLf
    s = s & "                           ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _" & vbCrLf
    s = s & "                           Optional ByVal multi As Boolean = False, _" & vbCrLf
    s = s & "                           Optional ByVal bloqueado As Boolean = False) As MSForms.TextBox" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim c As MSForms.TextBox" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set c = pai.Controls.Add(""Forms.TextBox.1"", nome, True)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    With c" & vbCrLf
    s = s & "        .Left = l: .Top = t: .Width = w: .Height = h" & vbCrLf
    s = s & "        .Font.Name = mFonte" & vbCrLf
    s = s & "        .Font.Size = 10" & vbCrLf
    s = s & "        .SpecialEffect = fmSpecialEffectFlat" & vbCrLf
    s = s & "        .BorderStyle = fmBorderStyleSingle" & vbCrLf
    s = s & "        .BorderColor = COR_BORDA" & vbCrLf
    s = s & "        .Multiline = multi" & vbCrLf
    s = s & "        .WordWrap = multi" & vbCrLf
    s = s & "        .EnterKeyBehavior = multi" & vbCrLf
    s = s & "        If multi Then .ScrollBars = fmScrollBarsVertical" & vbCrLf
    s = s & "        .Locked = bloqueado" & vbCrLf
    s = s & "        If bloqueado Then .BackColor = COR_LEITURA Else .BackColor = vbWhite" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set NovoTexto = c" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    CodPainel6 = s
End Function

Private Function CodPainel7() As String
    Dim s As String
    s = s & "Private Function NovoBotao(ByVal nome As String, ByVal cap As String, _" & vbCrLf
    s = s & "                           ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _" & vbCrLf
    s = s & "                           Optional ByVal corFundo As Long = -1, Optional ByVal corTexto As Long = -1) As MSForms.CommandButton" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim c As MSForms.CommandButton" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set c = Me.Controls.Add(""Forms.CommandButton.1"", nome, True)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    With c" & vbCrLf
    s = s & "        .Caption = cap" & vbCrLf
    s = s & "        .Left = l: .Top = t: .Width = w: .Height = h" & vbCrLf
    s = s & "        .Font.Name = mFonte" & vbCrLf
    s = s & "        .Font.Size = 10" & vbCrLf
    s = s & "        .Font.Bold = True" & vbCrLf
    s = s & "        .TakeFocusOnClick = False" & vbCrLf
    s = s & "        If corFundo <> -1 Then .BackColor = corFundo" & vbCrLf
    s = s & "        If corTexto <> -1 Then .ForeColor = corTexto" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set NovoBotao = c" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub MontarTela()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim l As MSForms.label" & vbCrLf
    s = s & "    Dim y As Single" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    '---------------- faixa de t" & ChrW(237) & "tulo" & vbCrLf
    s = s & "    Set l = NovoLabel(Me, ""lblFaixa"", """", 0, 0, LARG + 10, 60)" & vbCrLf
    s = s & "    l.BackStyle = fmBackStyleOpaque" & vbCrLf
    s = s & "    l.BackColor = COR_ESCURO" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set l = NovoLabel(Me, ""lblFaixaAmarela"", """", 0, 60, LARG + 10, 3)" & vbCrLf
    s = s & "    l.BackStyle = fmBackStyleOpaque" & vbCrLf
    s = s & "    l.BackColor = COR_AMARELO" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    NovoLabel Me, ""lblTitulo"", ""FOLLOW UP DE OBRAS  |  DADOS DA OBRA"", 14, 5, 520, 13, 9, True, COR_AZUL_CLARO" & vbCrLf
    s = s & "    Set lblNome = NovoLabel(Me, ""lblNome"", """", 14, 16, LARG - 28, 26, 17, True, vbWhite)" & vbCrLf
    s = s & "    lblNome.WordWrap = False" & vbCrLf
    s = s & "    lblNome.Font.Name = mFonteTitulo" & vbCrLf
    s = s & "    lblNome.Font.Bold = mTituloNegrito" & vbCrLf
    s = s & "    Set lblSub = NovoLabel(Me, ""lblSub"", """", 14, 42, LARG - 28, 16, 10, False, RGB(203, 213, 225))" & vbCrLf
    s = s & "    lblSub.WordWrap = False" & vbCrLf
    s = s & "" & vbCrLf
    CodPainel7 = s
End Function

Private Function CodPainel8() As String
    Dim s As String
    s = s & "    '---------------- navega" & ChrW(231) & "" & ChrW(227) & "o" & vbCrLf
    s = s & "    y = 72" & vbCrLf
    s = s & "    NovoLabel Me, ""lblBusca"", ""Buscar:"", 12, y + 3, 40, 16, 10, True" & vbCrLf
    s = s & "    Set txtBusca = NovoTexto(Me, ""txtBusca"", 52, y, 132, 21)" & vbCrLf
    s = s & "    txtBusca.ControlTipText = ""Filtra por C" & ChrW(211) & "DIGO, nome ou cidade""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    NovoLabel Me, ""lblObra"", ""Obra:"", 192, y + 3, 32, 16, 10, True" & vbCrLf
    s = s & "    Set cboObra = Me.Controls.Add(""Forms.ComboBox.1"", ""cboObra"", True)" & vbCrLf
    s = s & "    With cboObra" & vbCrLf
    s = s & "        .Left = 224: .Top = y: .Width = 292: .Height = 21" & vbCrLf
    s = s & "        .Style = fmStyleDropDownList" & vbCrLf
    s = s & "        .Font.Name = mFonte: .Font.Size = 10" & vbCrLf
    s = s & "        .ListRows = 20" & vbCrLf
    s = s & "        .SpecialEffect = fmSpecialEffectFlat" & vbCrLf
    s = s & "        .BorderStyle = fmBorderStyleSingle" & vbCrLf
    s = s & "        .BorderColor = COR_BORDA" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set chkVisiveis = Me.Controls.Add(""Forms.CheckBox.1"", ""chkVisiveis"", True)" & vbCrLf
    s = s & "    With chkVisiveis" & vbCrLf
    s = s & "        .Left = 524: .Top = y + 2: .Width = 116: .Height = 18" & vbCrLf
    s = s & "        .Caption = ""S" & ChrW(243) & " vis" & ChrW(237) & "veis no filtro""" & vbCrLf
    s = s & "        .Font.Name = mFonte: .Font.Size = 9" & vbCrLf
    s = s & "        .BackStyle = fmBackStyleTransparent" & vbCrLf
    s = s & "        .ControlTipText = ""Respeita os filtros da aba OBRAS""" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set btnAnterior = NovoBotao(""cmdAnterior"", ""< Anterior"", 644, y, 70, 22)" & vbCrLf
    s = s & "    Set btnProxima = NovoBotao(""cmdProxima"", ""Pr" & ChrW(243) & "xima >"", 718, y, 70, 22)" & vbCrLf
    s = s & "    Set lblContador = NovoLabel(Me, ""lblContador"", """", 792, y + 3, LARG - 800, 16, 9.5, False, COR_CINZA)" & vbCrLf
    s = s & "    lblContador.TextAlign = fmTextAlignRight" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    '---------------- coluna esquerda: dados da obra" & vbCrLf
    s = s & "    NovoLabel Me, ""lblSecInfo"", ""DADOS DA OBRA"", 12, 100, 200, 17, 10.5, True, COR_ESCURO" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set fraInfo = Me.Controls.Add(""Forms.Frame.1"", ""fraInfo"", True)" & vbCrLf
    s = s & "    With fraInfo" & vbCrLf
    s = s & "        .Left = 12: .Top = 118: .Width = 330: .Height = 324" & vbCrLf
    s = s & "        .Caption = """"" & vbCrLf
    s = s & "        .SpecialEffect = fmSpecialEffectFlat" & vbCrLf
    s = s & "        .BorderStyle = fmBorderStyleSingle" & vbCrLf
    s = s & "        .BorderColor = COR_BORDA" & vbCrLf
    s = s & "        .BackColor = vbWhite" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "" & vbCrLf
    CodPainel8 = s
End Function

Private Function CodPainel9() As String
    Dim s As String
    s = s & "    '---------------- coluna direita: observa" & ChrW(231) & "" & ChrW(245) & "es" & vbCrLf
    s = s & "    NovoLabel Me, ""lblSecHist"", ""HIST" & ChrW(211) & "RICO DE REGISTROS"", 354, 100, 300, 17, 10.5, True, COR_ESCURO" & vbCrLf
    s = s & "    Set btnEditarHist = NovoBotao(""cmdEditarHist"", ""Editar hist" & ChrW(243) & "rico"", LARG - 132, 97, 120, 20)" & vbCrLf
    s = s & "    btnEditarHist.Font.Size = 9" & vbCrLf
    s = s & "    Set txtHist = NovoTexto(Me, ""txtHist"", 354, 118, LARG - 366, 228, True, True)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    NovoLabel Me, ""lblSecNova"", ""NOVO REGISTRO"", 354, 352, 150, 17, 10.5, True, COR_ESCURO" & vbCrLf
    s = s & "    NovoLabel Me, ""lblDicaNova"", ""(data e hora s" & ChrW(227) & "o inseridas automaticamente ao salvar)"", 504, 354, 340, 16, 9, False, COR_CINZA" & vbCrLf
    s = s & "    Set txtNova = NovoTexto(Me, ""txtNova"", 354, 370, LARG - 366, 72, True, False)" & vbCrLf
    s = s & "    txtNova.BackColor = COR_EDICAO" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    '---------------- bot" & ChrW(245) & "es" & vbCrLf
    s = s & "    y = 454" & vbCrLf
    s = s & "    Set btnSalvar = NovoBotao(""cmdSalvar"", ""SALVAR"", 12, y, 100, 30, COR_AZUL, vbWhite)" & vbCrLf
    s = s & "    btnSalvar.ControlTipText = ""Grava as datas alteradas e a nova observa" & ChrW(231) & "" & ChrW(227) & "o""" & vbCrLf
    s = s & "    Set btnAta = NovoBotao(""cmdAta"", ""GERAR ATA (WORD)"", 120, y, 136, 30, COR_ESCURO, vbWhite)" & vbCrLf
    s = s & "    btnAta.ControlTipText = ""Ata com a " & ChrW(218) & "LTIMO REGISTRO das obras da lista do painel""" & vbCrLf
    s = s & "    Set btnAtualizar = NovoBotao(""cmdAtualizar"", ""ATUALIZAR DADOS"", 264, y, 130, 30, COR_AZUL_CLARO, COR_ESCURO)" & vbCrLf
    s = s & "    btnAtualizar.ControlTipText = ""Atualiza as consultas da planilha e recarrega o painel""" & vbCrLf
    s = s & "    Set btnNova = NovoBotao(""cmdNovaObra"", ""NOVA OBRA"", 402, y, 104, 30, COR_AMARELO, COR_ESCURO)" & vbCrLf
    s = s & "    btnNova.ControlTipText = ""Inclui uma obra no follow up pelo C" & ChrW(211) & "DIGO""" & vbCrLf
    s = s & "    Set btnIrPlanilha = NovoBotao(""cmdIrPlanilha"", ""IR PARA A LINHA"", 514, y, 120, 30)" & vbCrLf
    s = s & "    Set btnFechar = NovoBotao(""cmdFechar"", ""FECHAR"", LARG - 112, y, 100, 30)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    '---------------- calend" & ChrW(225) & "rio, feriados e marcador (instalado por mInstaladorPainelObras)" & vbCrLf
    s = s & "    ExtrasMontar" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    '---------------- OBRA ENTREGUE (instalado por mInstaladorEntregueObras)" & vbCrLf
    s = s & "    Set btnEntrega = NovoBotao(""cmdEntrega"", ""OBRA ENTREGUE"", btnIrPlanilha.Left + btnIrPlanilha.Width + 8, btnIrPlanilha.Top, 140, btnIrPlanilha.Height, RGB(22, 163, 74), vbWhite)" & vbCrLf
    s = s & "    If btnEntrega.Left + btnEntrega.Width > LARG - 120 Then btnEntrega.Width = LARG - 120 - btnEntrega.Left" & vbCrLf
    s = s & "    btnEntrega.ControlTipText = ""Leva esta obra para ENTREGUES (confirma" & ChrW(231) & "" & ChrW(227) & "o dentro do painel)""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub MontarCamposInfo()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim lista As Variant" & vbCrLf
    s = s & "    Dim i As Long, c As Long, n As Long" & vbCrLf
    s = s & "    Dim y As Single" & vbCrLf
    s = s & "    Dim campo As clsCampoData" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    'Coluna da tabela tb_Obras, r" & ChrW(243) & "tulo no painel e tipo:" & vbCrLf
    s = s & "    '   D = data (edit" & ChrW(225) & "vel quando a c" & ChrW(233) & "lula n" & ChrW(227) & "o tem f" & ChrW(243) & "rmula)" & vbCrLf
    s = s & "    '   N = s" & ChrW(243) & " consulta" & vbCrLf
    s = s & "    'Para incluir ou tirar um campo do painel, basta mexer nesta lista." & vbCrLf
    CodPainel9 = s
End Function

Private Function CodPainel10() As String
    Dim s As String
    s = s & "    lista = Array( _" & vbCrLf
    s = s & "        Array(""ALVAR" & ChrW(193) & """, ""ALVAR" & ChrW(193) & """, ""D""), _" & vbCrLf
    s = s & "        Array(""PROJETO EXECUTIVO"", ""PROJETO EXECUTIVO"", ""D""), _" & vbCrLf
    s = s & "        Array(""IN" & ChrW(205) & "CIO DA OBRA"", ""IN" & ChrW(205) & "CIO DA OBRA"", ""D""), _" & vbCrLf
    s = s & "        Array(""PRAZO (DIAS)"", ""PRAZO (DIAS)"", ""N""), _" & vbCrLf
    s = s & "        Array(""FIM DA OBRA"", ""FIM DA OBRA"", ""D""), _" & vbCrLf
    s = s & "        Array(""ENTREGA"", ""ENTREGA"", ""D""))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    n = UBound(lista) - LBound(lista) + 1" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ReDim mCampoIdx(1 To n)" & vbCrLf
    s = s & "    ReDim mCampoRot(1 To n)" & vbCrLf
    s = s & "    ReDim mCampoTipo(1 To n)" & vbCrLf
    s = s & "    ReDim mEditavel(1 To n)" & vbCrLf
    s = s & "    ReDim mOrig(1 To n)" & vbCrLf
    s = s & "    ReDim mLbl(1 To n)" & vbCrLf
    s = s & "    ReDim mBoxEd(1 To n)" & vbCrLf
    s = s & "    ReDim mBoxRo(1 To n)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set mCamposData = New Collection" & vbCrLf
    s = s & "    mQtdCampos = 0" & vbCrLf
    s = s & "    y = 12" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = LBound(lista) To UBound(lista)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        c = PO_ColIdx(tb, CStr(lista(i)(0)))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If c > 0 Then" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            mQtdCampos = mQtdCampos + 1" & vbCrLf
    s = s & "            n = mQtdCampos" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            mCampoIdx(n) = c" & vbCrLf
    s = s & "            mCampoRot(n) = CStr(lista(i)(1))" & vbCrLf
    s = s & "            mCampoTipo(n) = CStr(lista(i)(2))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            Set mLbl(n) = NovoLabel(fraInfo, ""lblCampo"" & n, mCampoRot(n), 6, y + 3, 120, 16, 8, True, COR_ESCURO)" & vbCrLf
    s = s & "            mLbl(n).WordWrap = False" & vbCrLf
    s = s & "            mLbl(n).ControlTipText = tb.ListColumns(c).Name" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            'duas caixas no mesmo lugar: a de consulta (cinza) e a de edi" & ChrW(231) & "" & ChrW(227) & "o;" & vbCrLf
    s = s & "            'o painel mostra uma ou outra conforme a c" & ChrW(233) & "lula da obra" & vbCrLf
    s = s & "            Set mBoxRo(n) = NovoTexto(fraInfo, ""txtRo"" & n, 128, y, 180, 21, False, True)" & vbCrLf
    s = s & "            mBoxRo(n).TabStop = False" & vbCrLf
    s = s & "" & vbCrLf
    CodPainel10 = s
End Function

Private Function CodPainel11() As String
    Dim s As String
    s = s & "            Set mBoxEd(n) = NovoTexto(fraInfo, ""txtEd"" & n, 128, y, 180, 21, False, False)" & vbCrLf
    s = s & "            mBoxEd(n).BorderColor = COR_AMARELO" & vbCrLf
    s = s & "            mBoxEd(n).MaxLength = 10" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            'm" & ChrW(225) & "scara de data, vermelho se inv" & ChrW(225) & "lida, duplo clique = hoje" & vbCrLf
    s = s & "            Set campo = New clsCampoData" & vbCrLf
    s = s & "            Set campo.Caixa = mBoxEd(n)" & vbCrLf
    s = s & "            mCamposData.Add campo" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            y = y + 46" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    NovoLabel fraInfo, ""lblCampoDica"", ""Borda amarela = pode alterar (15112026 ou dd/mm/aaaa; duplo clique = hoje). "" & _" & vbCrLf
    s = s & "              ""Cinza = vem da f" & ChrW(243) & "rmula da tabela."", 6, y + 2, 302, 26, 8, False, COR_CINZA" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub AjustarZoom()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    'Ajusta o painel ao tamanho REAL da tela (" & ChrW(225) & "rea sem a barra de tarefas)" & vbCrLf
    s = s & "    Dim telaW As Single, telaH As Single" & vbCrLf
    s = s & "    Dim f As Single" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Fim" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not AreaTelaPt(telaW, telaH) Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    f = (telaH - 16) / Me.Height" & vbCrLf
    s = s & "    If (telaW - 16) / Me.Width < f Then f = (telaW - 16) / Me.Width" & vbCrLf
    s = s & "    If f > 1.3 Then f = 1.3" & vbCrLf
    s = s & "    If f < 0.75 Then f = 0.75" & vbCrLf
    s = s & "    If Abs(f - 1) < 0.02 Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Me.Zoom = Int(f * 100)" & vbCrLf
    s = s & "    Me.Width = Me.Width * f" & vbCrLf
    s = s & "    Me.Height = (Me.Height - 28) * f + 28" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fim:" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    CodPainel11 = s
End Function

Private Function CodPainel12() As String
    Dim s As String
    s = s & "' LISTA DE OBRAS / NAVEGA" & ChrW(199) & "" & ChrW(195) & "O" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub FiltrarLista()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim f As String" & vbCrLf
    s = s & "    Dim txt As String" & vbCrLf
    s = s & "    Dim lj As New clsObra" & vbCrLf
    s = s & "    Dim ok As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    f = PO_Norm(txtBusca.Text)" & vbCrLf
    s = s & "    mQtd = 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If tb.ListRows.Count = 0 Then Exit Sub" & vbCrLf
    s = s & "    ReDim mMapa(1 To tb.ListRows.Count)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mCarregando = True" & vbCrLf
    s = s & "    cboObra.Clear" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 1 To tb.ListRows.Count" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        lj.Carregar tb, i" & vbCrLf
    s = s & "        txt = lj.rotulo" & vbCrLf
    s = s & "        ok = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If f <> """" Then" & vbCrLf
    s = s & "            ok = (InStr(1, PO_Norm(txt & "" "" & lj.texto(PO_COL_CIDADE)), f) > 0)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If ok And Not chkMarcadas Is Nothing Then" & vbCrLf
    s = s & "            If chkMarcadas.value = True Then ok = MarcasPainel().Marcada(lj.CodObra)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        If ok And chkVisiveis.value = True Then" & vbCrLf
    s = s & "            ok = lj.Visivel" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If ok Then" & vbCrLf
    s = s & "            mQtd = mQtd + 1" & vbCrLf
    s = s & "            mMapa(mQtd) = i" & vbCrLf
    s = s & "            cboObra.AddItem txt" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    CodPainel12 = s
End Function

Private Function CodPainel13() As String
    Dim s As String
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SelecionarNoCombo mAtual" & vbCrLf
    s = s & "    mCarregando = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function PosicaoNaLista(ByVal idx As Long) As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 1 To mQtd" & vbCrLf
    s = s & "        If mMapa(i) = idx Then" & vbCrLf
    s = s & "            PosicaoNaLista = i" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SelecionarNoCombo(ByVal idx As Long)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim pos As Long" & vbCrLf
    s = s & "    Dim antes As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    antes = mCarregando" & vbCrLf
    s = s & "    mCarregando = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    pos = PosicaoNaLista(idx)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If pos > 0 Then" & vbCrLf
    s = s & "        cboObra.ListIndex = pos - 1" & vbCrLf
    s = s & "        lblContador.Caption = pos & "" de "" & mQtd" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        cboObra.ListIndex = -1" & vbCrLf
    s = s & "        lblContador.Caption = ""- de "" & mQtd" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mCarregando = antes" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub txtBusca_Change()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mCarregando Then Exit Sub" & vbCrLf
    CodPainel13 = s
End Function

Private Function CodPainel14() As String
    Dim s As String
    s = s & "    FiltrarLista" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    'Se sobrou s" & ChrW(243) & " uma obra, abre direto" & vbCrLf
    s = s & "    If mQtd = 1 Then" & vbCrLf
    s = s & "        If mMapa(1) <> mAtual Then" & vbCrLf
    s = s & "            If ConfirmarPendencias() Then CarregarObra mMapa(1)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub chkVisiveis_Click()" & vbCrLf
    s = s & "    If mCarregando Then Exit Sub" & vbCrLf
    s = s & "    FiltrarLista" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboObra_Change()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim novo As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mCarregando Then Exit Sub" & vbCrLf
    s = s & "    If cboObra.ListIndex < 0 Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    novo = mMapa(cboObra.ListIndex + 1)" & vbCrLf
    s = s & "    If novo = mAtual Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not ConfirmarPendencias() Then" & vbCrLf
    s = s & "        SelecionarNoCombo mAtual" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    CarregarObra novo" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnAnterior_Click()" & vbCrLf
    s = s & "    Navegar -1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnProxima_Click()" & vbCrLf
    s = s & "    Navegar 1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Navegar(ByVal passo As Long)" & vbCrLf
    s = s & "" & vbCrLf
    CodPainel14 = s
End Function

Private Function CodPainel15() As String
    Dim s As String
    s = s & "    Dim pos As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mQtd = 0 Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    pos = PosicaoNaLista(mAtual)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If pos = 0 Then" & vbCrLf
    s = s & "        pos = 1" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        pos = pos + passo" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If pos < 1 Or pos > mQtd Then Exit Sub" & vbCrLf
    s = s & "    If Not ConfirmarPendencias() Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    CarregarObra mMapa(pos)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' CARREGAR DADOS" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "Private Sub CarregarObra(ByVal idx As Long)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    If idx < 1 Or idx > tb.ListRows.Count Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mCarregando = True" & vbCrLf
    s = s & "    mAtual = idx" & vbCrLf
    s = s & "    Set mObra = New clsObra" & vbCrLf
    s = s & "    mObra.Carregar tb, idx" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    CarregarCampos" & vbCrLf
    s = s & "    CarregarObservacoes" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SelecionarNoCombo mAtual" & vbCrLf
    s = s & "    mCarregando = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    RolarHistoricoFim" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub CarregarCampos()" & vbCrLf
    s = s & "" & vbCrLf
    CodPainel15 = s
End Function

Private Function CodPainel16() As String
    Dim s As String
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim cel As Range" & vbCrLf
    s = s & "    Dim txt As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblNome.Caption = mObra.nome" & vbCrLf
    s = s & "    lblSub.Caption = mObra.Resumo" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mQtdCampos = 0 Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    'tira o cursor das caixas de data antes de mexer nelas" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    If Me.Visible Then txtNova.SetFocus" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 1 To mQtdCampos" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        Set cel = mObra.Linha.Range.Cells(1, mCampoIdx(i))" & vbCrLf
    s = s & "        txt = mObra.TextoColuna(mCampoIdx(i))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        'data digitada = edit" & ChrW(225) & "vel; c" & ChrW(233) & "lula com f" & ChrW(243) & "rmula = s" & ChrW(243) & " consulta" & vbCrLf
    s = s & "        mEditavel(i) = (mCampoTipo(i) = ""D"") And Not CBool(cel.HasFormula)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        mBoxRo(i).Text = txt" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If mEditavel(i) Then" & vbCrLf
    s = s & "            mBoxEd(i).Text = txt" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            mBoxEd(i).Text = """"" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        mOrig(i) = mBoxEd(i).Text" & vbCrLf
    s = s & "        mBoxEd(i).BackColor = vbWhite" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        mBoxEd(i).Visible = mEditavel(i)" & vbCrLf
    s = s & "        mBoxRo(i).Visible = Not mEditavel(i)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If mEditavel(i) Then" & vbCrLf
    s = s & "            mBoxEd(i).ControlTipText = tb.ListColumns(mCampoIdx(i)).Name & ""  |  duplo clique = hoje""" & vbCrLf
    s = s & "        ElseIf cel.HasFormula Then" & vbCrLf
    s = s & "            mBoxRo(i).ControlTipText = ""Preenchido pela f" & ChrW(243) & "rmula da tabela""" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            mBoxRo(i).ControlTipText = """"" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Next i" & vbCrLf
    CodPainel16 = s
End Function

Private Function CodPainel17() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "    ExtrasAtualizar" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'A data i foi alterada pelo usu" & ChrW(225) & "rio?" & vbCrLf
    s = s & "Private Function DataAlterada(ByVal i As Long) As Boolean" & vbCrLf
    s = s & "    If mEditavel(i) Then DataAlterada = (mBoxEd(i).Text <> mOrig(i))" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function DatasAlteradas() As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mObra Is Nothing Then Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 1 To mQtdCampos" & vbCrLf
    s = s & "        If DataAlterada(i) Then" & vbCrLf
    s = s & "            DatasAlteradas = True" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub CarregarObservacoes()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtHist.Text = PO_ParaCaixa(mObra.Historico)" & vbCrLf
    s = s & "    mHistOriginal = txtHist.Text" & vbCrLf
    s = s & "    txtNova.Text = """"" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ModoEdicaoHistorico False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub RolarHistoricoFim()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    If Not Me.Visible Then Exit Sub" & vbCrLf
    s = s & "    txtHist.SetFocus" & vbCrLf
    s = s & "    txtHist.SelStart = Len(txtHist.Text)" & vbCrLf
    s = s & "    txtHist.SelLength = 0" & vbCrLf
    s = s & "    txtNova.SetFocus" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    CodPainel17 = s
End Function

Private Function CodPainel18() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' OBSERVA" & ChrW(199) & "" & ChrW(213) & "ES" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "Private Sub ModoEdicaoHistorico(ByVal ligado As Boolean)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mEditandoHist = ligado" & vbCrLf
    s = s & "    txtHist.Locked = Not ligado" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If ligado Then" & vbCrLf
    s = s & "        txtHist.BackColor = COR_EDICAO" & vbCrLf
    s = s & "        btnEditarHist.Caption = ""Cancelar edi" & ChrW(231) & "" & ChrW(227) & "o""" & vbCrLf
    s = s & "        btnEditarHist.ForeColor = COR_VERMELHO" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        txtHist.BackColor = RGB(248, 249, 251)" & vbCrLf
    s = s & "        btnEditarHist.Caption = ""Editar hist" & ChrW(243) & "rico""" & vbCrLf
    s = s & "        btnEditarHist.ForeColor = COR_ESCURO" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function HistoricoAlterado() As Boolean" & vbCrLf
    s = s & "    HistoricoAlterado = mEditandoHist And (PO_ParaCaixa(txtHist.Text) <> PO_ParaCaixa(mHistOriginal))" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function TemPendencias() As Boolean" & vbCrLf
    s = s & "    If mAtual = 0 Then Exit Function" & vbCrLf
    s = s & "    TemPendencias = (Trim$(Replace(Replace(txtNova.Text, vbCr, """"), vbLf, """")) <> """") _" & vbCrLf
    s = s & "                    Or HistoricoAlterado() Or DatasAlteradas()" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'True = pode seguir | False = usu" & ChrW(225) & "rio cancelou" & vbCrLf
    s = s & "Private Function ConfirmarPendencias() As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim r As VbMsgBoxResult" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not TemPendencias() Then" & vbCrLf
    s = s & "        ConfirmarPendencias = True" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    r = MsgBox(""H" & ChrW(225) & " altera" & ChrW(231) & "" & ChrW(245) & "es n" & ChrW(227) & "o salvas em '"" & lblNome.Caption & ""'."" & vbCrLf & vbCrLf & _" & vbCrLf
    s = s & "               ""Deseja salvar antes de continuar?"", vbYesNoCancel + vbQuestion, ""Painel de Obras"")" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Select Case r" & vbCrLf
    CodPainel18 = s
End Function

Private Function CodPainel19() As String
    Dim s As String
    s = s & "        Case vbYes:  ConfirmarPendencias = SalvarObra(True)" & vbCrLf
    s = s & "        Case vbNo:   ConfirmarPendencias = True" & vbCrLf
    s = s & "        Case Else:   ConfirmarPendencias = False" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnEditarHist_Click()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mAtual = 0 Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mEditandoHist Then" & vbCrLf
    s = s & "        If HistoricoAlterado() Then" & vbCrLf
    s = s & "            If MsgBox(""Descartar as altera" & ChrW(231) & "" & ChrW(245) & "es feitas no hist" & ChrW(243) & "rico?"", vbYesNo + vbQuestion, ""Painel de Obras"") = vbNo Then Exit Sub" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        txtHist.Text = mHistOriginal" & vbCrLf
    s = s & "        ModoEdicaoHistorico False" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        ModoEdicaoHistorico True" & vbCrLf
    s = s & "        On Error Resume Next" & vbCrLf
    s = s & "        txtHist.SetFocus" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnSalvar_Click()" & vbCrLf
    s = s & "    SalvarObra False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function SalvarObra(ByVal silencioso As Boolean) As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim nova As String" & vbCrLf
    s = s & "    Dim histMudou As Boolean" & vbCrLf
    s = s & "    Dim datasMudaram As Boolean" & vbCrLf
    s = s & "    Dim obsMudou As Boolean" & vbCrLf
    s = s & "    Dim ok As Boolean" & vbCrLf
    s = s & "    Dim i As Long, r As Long" & vbCrLf
    s = s & "    Dim qtdDatas As Long, falhas As Long" & vbCrLf
    s = s & "    Dim d As Date, vz As Boolean" & vbCrLf
    s = s & "    Dim erros As String, msg As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mObra Is Nothing Then Exit Function" & vbCrLf
    s = s & "    If Not FerPodeSalvar() Then Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    nova = Trim$(PO_TrimFim(PO_ParaCaixa(txtNova.Text)))" & vbCrLf
    CodPainel19 = s
End Function

Private Function CodPainel20() As String
    Dim s As String
    s = s & "    histMudou = HistoricoAlterado()" & vbCrLf
    s = s & "    datasMudaram = DatasAlteradas()" & vbCrLf
    s = s & "    obsMudou = (nova <> """" Or histMudou)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not obsMudou And Not datasMudaram Then" & vbCrLf
    s = s & "        If Not silencioso Then MsgBox ""N" & ChrW(227) & "o h" & ChrW(225) & " altera" & ChrW(231) & "" & ChrW(245) & "es para salvar."", vbInformation, ""Painel de Obras""" & vbCrLf
    s = s & "        SalvarObra = True" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    'valida todas as datas antes de gravar qualquer coisa" & vbCrLf
    s = s & "    If datasMudaram Then" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        For i = 1 To mQtdCampos" & vbCrLf
    s = s & "            If DataAlterada(i) Then" & vbCrLf
    s = s & "                If Not PO_TextoParaData(mBoxEd(i).Text, d, vz) Then" & vbCrLf
    s = s & "                    erros = erros & vbCrLf & ""  - "" & mCampoRot(i) & "": "" & mBoxEd(i).Text" & vbCrLf
    s = s & "                    mBoxEd(i).BackColor = PO_COR_ERRO" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If erros <> """" Then" & vbCrLf
    s = s & "            MsgBox ""Datas inv" & ChrW(225) & "lidas (use dd/mm/aaaa):"" & erros & vbCrLf & vbCrLf & _" & vbCrLf
    s = s & "                   ""Nada foi gravado. Corrija e salve de novo."", vbExclamation, ""Painel de Obras""" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Application.ScreenUpdating = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If datasMudaram Then" & vbCrLf
    s = s & "        For i = 1 To mQtdCampos" & vbCrLf
    s = s & "            If DataAlterada(i) Then" & vbCrLf
    s = s & "                r = mObra.GravarData(mCampoIdx(i), mBoxEd(i).Text)" & vbCrLf
    s = s & "                If r = 1 Then qtdDatas = qtdDatas + 1" & vbCrLf
    s = s & "                If r < 0 Then falhas = falhas + 1" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        Next i" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ok = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If nova <> """" Then" & vbCrLf
    CodPainel20 = s
End Function

Private Function CodPainel21() As String
    Dim s As String
    s = s & "        If histMudou Then" & vbCrLf
    s = s & "            ok = mObra.AdicionarObservacao(nova, PO_ParaCaixa(txtHist.Text))" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            ok = mObra.AdicionarObservacao(nova)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    ElseIf histMudou Then" & vbCrLf
    s = s & "        ok = mObra.SubstituirHistorico(txtHist.Text)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Application.ScreenUpdating = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    'recarrega o que foi gravado (as datas voltam a ser o ""original"")" & vbCrLf
    s = s & "    mCarregando = True" & vbCrLf
    s = s & "    If datasMudaram Then CarregarCampos" & vbCrLf
    s = s & "    If obsMudou And ok Then CarregarObservacoes" & vbCrLf
    s = s & "    mCarregando = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If obsMudou And ok Then RolarHistoricoFim" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If falhas > 0 Then" & vbCrLf
    s = s & "        MsgBox qtdDatas & "" data(s) gravada(s); "" & falhas & "" n" & ChrW(227) & "o puderam ser gravadas."", _" & vbCrLf
    s = s & "               vbExclamation, ""Painel de Obras""" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not ok Then Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not silencioso Then" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If qtdDatas > 0 Then msg = qtdDatas & "" data(s) atualizada(s) na planilha.""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If obsMudou Then" & vbCrLf
    s = s & "            If msg <> """" Then msg = msg & vbCrLf" & vbCrLf
    s = s & "            msg = msg & ""Observa" & ChrW(231) & "" & ChrW(227) & "o salva com sucesso!""" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If msg = """" Then msg = ""Nenhuma altera" & ChrW(231) & "" & ChrW(227) & "o: os valores j" & ChrW(225) & " eram os mesmos.""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        MsgBox msg, vbInformation, ""Painel de Obras""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SalvarObra = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    CodPainel21 = s
End Function

Private Function CodPainel22() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' BOT" & ChrW(213) & "ES" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "Private Sub btnAta_Click()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not ConfirmarPendencias() Then Exit Sub" & vbCrLf
    s = s & "    PO_GerarAta tb, mMapa, mQtd" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'---------------------------------------------------------- ATUALIZAR DADOS" & vbCrLf
    s = s & "Private Sub btnAtualizar_Click()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim antes As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not ConfirmarPendencias() Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If MsgBox(""Atualizar agora todas as consultas da planilha?"" & vbCrLf & vbCrLf & _" & vbCrLf
    s = s & "              ""Pode levar alguns minutos. O painel " & ChrW(233) & " recarregado no final."", _" & vbCrLf
    s = s & "              vbYesNo + vbQuestion, ""Painel de Obras"") = vbNo Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    antes = mAtual" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    'usa a macro de atualiza" & ChrW(231) & "" & ChrW(227) & "o que j" & ChrW(225) & " existe na planilha; se ela n" & ChrW(227) & "o" & vbCrLf
    s = s & "    'estiver dispon" & ChrW(237) & "vel, atualiza tudo direto" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Application.Run ""'"" & Replace(ThisWorkbook.Name, ""'"", ""''"") & ""'!AtualizarTodasAsQueries""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Err.Number <> 0 Then" & vbCrLf
    s = s & "        Err.Clear" & vbCrLf
    s = s & "        Application.StatusBar = ""Atualizando consultas...""" & vbCrLf
    s = s & "        ThisWorkbook.RefreshAll" & vbCrLf
    s = s & "        Application.CalculateUntilAsyncQueriesDone" & vbCrLf
    s = s & "        Application.StatusBar = False" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    Application.ScreenUpdating = True" & vbCrLf
    s = s & "    Application.EnableEvents = True" & vbCrLf
    s = s & "    Application.Calculation = xlCalculationAutomatic" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Recarregar antes" & vbCrLf
    CodPainel22 = s
End Function

Private Function CodPainel23() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'L" & ChrW(234) & " a tabela de novo e reabre a obra indicada (ou a primeira da lista)" & vbCrLf
    s = s & "Private Sub Recarregar(ByVal idx As Long)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set tb = PO_TabelaPorNome(TB_OBRAS)" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mAtual = 0" & vbCrLf
    s = s & "    FiltrarLista" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If idx >= 1 And idx <= tb.ListRows.Count Then" & vbCrLf
    s = s & "        If PosicaoNaLista(idx) > 0 Then" & vbCrLf
    s = s & "            CarregarObra idx" & vbCrLf
    s = s & "            Exit Sub" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mQtd > 0 Then CarregarObra mMapa(1)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'---------------------------------------------------------- NOVA OBRA" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Abre a obra idx, tirando busca e filtro se ela n" & ChrW(227) & "o estiver na lista" & vbCrLf
    s = s & "Private Sub MostrarObra(ByVal idx As Long)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mAtual = 0" & vbCrLf
    s = s & "    FiltrarLista" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If PosicaoNaLista(idx) = 0 Then" & vbCrLf
    s = s & "        mCarregando = True" & vbCrLf
    s = s & "        txtBusca.Text = """"" & vbCrLf
    s = s & "        chkVisiveis.value = False" & vbCrLf
    s = s & "        mCarregando = False" & vbCrLf
    s = s & "        FiltrarLista" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    CarregarObra idx" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Acrescenta uma linha no fim da tabela e devolve o " & ChrW(237) & "ndice dela (0 = falhou)." & vbCrLf
    s = s & "'Com a tabela filtrada o Excel pode recusar a inclus" & ChrW(227) & "o normal; nesse caso" & vbCrLf
    CodPainel23 = s
End Function

Private Function CodPainel24() As String
    Dim s As String
    s = s & "'a tabela " & ChrW(233) & " ampliada em uma linha." & vbCrLf
    s = s & "Private Function NovaLinhaTabela() As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim antes As Long" & vbCrLf
    s = s & "    Dim rng As Range" & vbCrLf
    s = s & "    Dim evt As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    antes = tb.ListRows.Count" & vbCrLf
    s = s & "    evt = Application.EnableEvents" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Application.EnableEvents = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    tb.ListRows.Add" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If tb.ListRows.Count <= antes Then" & vbCrLf
    s = s & "        Err.Clear" & vbCrLf
    s = s & "        Set rng = tb.Range" & vbCrLf
    s = s & "        tb.Resize rng.Resize(rng.Rows.Count + 1)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Application.EnableEvents = evt" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If tb.ListRows.Count <= antes Then" & vbCrLf
    s = s & "        MsgBox ""O Excel recusou a nova linha na tabela '"" & TB_OBRAS & ""'."" & vbCrLf & _" & vbCrLf
    s = s & "               ""Confira se a planilha n" & ChrW(227) & "o est" & ChrW(225) & " protegida e tente limpar o filtro da tabela."", _" & vbCrLf
    s = s & "               vbExclamation, ""Painel de Obras""" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    NovaLinhaTabela = tb.ListRows.Count" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Colunas que n" & ChrW(227) & "o s" & ChrW(227) & "o ""calculadas"" pela tabela, mas t" & ChrW(234) & "m f" & ChrW(243) & "rmula na linha" & vbCrLf
    s = s & "'de cima: a f" & ChrW(243) & "rmula " & ChrW(233) & " repetida na nova linha, para a obra nova puxar os" & vbCrLf
    s = s & "'dados do mesmo jeito que as demais." & vbCrLf
    s = s & "Private Sub CopiarFormulasDaLinhaAnterior(ByVal nova As Long, ByVal idxCod As Long)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Dim celNova As Range, celAnt As Range" & vbCrLf
    s = s & "    Dim evt As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    CodPainel24 = s
End Function

Private Function CodPainel25() As String
    Dim s As String
    s = s & "    If nova < 2 Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    evt = Application.EnableEvents" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Application.EnableEvents = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For c = 1 To tb.ListColumns.Count" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If c <> idxCod Then" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            Set celNova = tb.ListRows(nova).Range.Cells(1, c)" & vbCrLf
    s = s & "            Set celAnt = tb.ListRows(nova - 1).Range.Cells(1, c)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "            If Not celNova.HasFormula Then" & vbCrLf
    s = s & "                If IsEmpty(celNova.value) And celAnt.HasFormula Then" & vbCrLf
    s = s & "                    celNova.Formula = celAnt.Formula" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Next c" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Application.EnableEvents = evt" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnIrPlanilha_Click()" & vbCrLf
    s = s & "    If Not mObra Is Nothing Then mObra.IrParaPlanilha" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnFechar_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' FONTES DA MARCA (com reserva se n" & ChrW(227) & "o estiverem instaladas)" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "Private Sub DefinirFontes()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If FonteInstalada(FONTE_MARCA) Then" & vbCrLf
    s = s & "        mFonte = FONTE_MARCA" & vbCrLf
    CodPainel25 = s
End Function

Private Function CodPainel26() As String
    Dim s As String
    s = s & "    Else" & vbCrLf
    s = s & "        mFonte = FONTE_RESERVA" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If FonteInstalada(FONTE_MARCA_TITULO) Then" & vbCrLf
    s = s & "        mFonteTitulo = FONTE_MARCA_TITULO" & vbCrLf
    s = s & "        mTituloNegrito = False" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        mFonteTitulo = mFonte" & vbCrLf
    s = s & "        mTituloNegrito = True" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Pede a fonte ao Windows e confere se ele entregou a mesma (se n" & ChrW(227) & "o" & vbCrLf
    s = s & "'existir, o Windows devolve outra no lugar)." & vbCrLf
    s = s & "Private Function FonteInstalada(ByVal nome As String) As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "#If VBA7 Then" & vbCrLf
    s = s & "    Dim hdc As LongPtr" & vbCrLf
    s = s & "    Dim hFonte As LongPtr" & vbCrLf
    s = s & "    Dim hAntes As LongPtr" & vbCrLf
    s = s & "#Else" & vbCrLf
    s = s & "    Dim hdc As Long" & vbCrLf
    s = s & "    Dim hFonte As Long" & vbCrLf
    s = s & "    Dim hAntes As Long" & vbCrLf
    s = s & "#End If" & vbCrLf
    s = s & "    Dim buf As String" & vbCrLf
    s = s & "    Dim n As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Fim" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    hdc = GetDC(0)" & vbCrLf
    s = s & "    If hdc = 0 Then Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    hFonte = CreateFontA(-12, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 0, 0, nome)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If hFonte <> 0 Then" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        hAntes = SelectObject(hdc, hFonte)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        buf = String$(64, vbNullChar)" & vbCrLf
    s = s & "        n = GetTextFaceA(hdc, 64, buf)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If n > 1 Then" & vbCrLf
    CodPainel26 = s
End Function

Private Function CodPainel27() As String
    Dim s As String
    s = s & "            FonteInstalada = (StrComp(Left$(buf, n - 1), nome, vbTextCompare) = 0)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        SelectObject hdc, hAntes" & vbCrLf
    s = s & "        DeleteObject hFonte" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ReleaseDC 0, hdc" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fim:" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' " & ChrW(193) & "REA " & ChrW(218) & "TIL DA TELA EM PONTOS (sem a barra de tarefas)" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "Private Function AreaTelaPt(ByRef largura As Single, ByRef altura As Single) As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim r As RECT_TELA" & vbCrLf
    s = s & "    Dim dpi As Long" & vbCrLf
    s = s & "#If VBA7 Then" & vbCrLf
    s = s & "    Dim hdc As LongPtr" & vbCrLf
    s = s & "#Else" & vbCrLf
    s = s & "    Dim hdc As Long" & vbCrLf
    s = s & "#End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Fim" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If SystemParametersInfoA(48, 0, r, 0) = 0 Then Exit Function    'SPI_GETWORKAREA" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    hdc = GetDC(0)" & vbCrLf
    s = s & "    dpi = GetDeviceCaps(hdc, 90)                                     'LOGPIXELSY" & vbCrLf
    s = s & "    ReleaseDC 0, hdc" & vbCrLf
    s = s & "    If dpi <= 0 Then dpi = 96" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    largura = (r.Direita - r.Esq) * 72 / dpi" & vbCrLf
    s = s & "    altura = (r.Baixo - r.Topo) * 72 / dpi" & vbCrLf
    s = s & "    AreaTelaPt = (largura > 200 And altura > 200)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Fim:" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    CodPainel27 = s
End Function

Private Function CodPainel28() As String
    Dim s As String
    s = s & "' ABREVIA CABE" & ChrW(199) & "ALHOS LONGOS (nome completo fica no ControlTipText)" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "Private Function RotuloCurto(ByVal cab As String) As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Const LIMITE As Long = 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim s As String" & vbCrLf
    s = s & "    Dim abrev As Variant" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim p1 As Long, p2 As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    s = UCase$(Trim$(Replace(Replace(cab, vbCr, "" ""), vbLf, "" "")))" & vbCrLf
    s = s & "    Do While InStr(s, ""  "") > 0" & vbCrLf
    s = s & "        s = Replace(s, ""  "", "" "")" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Len(s) <= LIMITE Then RotuloCurto = s: Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    '0) par" & ChrW(234) & "nteses com texto longo: mant" & ChrW(233) & "m s" & ChrW(243) & " o que vem antes" & vbCrLf
    s = s & "    p1 = InStr(s, ""("")" & vbCrLf
    s = s & "    If p1 > 1 Then" & vbCrLf
    s = s & "        p2 = InStr(p1, s, "")"")" & vbCrLf
    s = s & "        If p2 = 0 Then p2 = Len(s)" & vbCrLf
    s = s & "        If (p2 - p1) > 16 And Len(Trim$(Left$(s, p1 - 1))) <= LIMITE Then" & vbCrLf
    s = s & "            RotuloCurto = Trim$(Left$(s, p1 - 1))" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    '1) abrevia" & ChrW(231) & "" & ChrW(245) & "es comuns, uma por vez, at" & ChrW(233) & " caber" & vbCrLf
    s = s & "    abrev = Array(""PROJETO B" & ChrW(193) & "SICO"", ""PROJ. B" & ChrW(193) & "S."", ""PROJETO EXECUTIVO"", ""PROJ. EXEC."", _" & vbCrLf
    s = s & "                  ""CONTRATO ASSINADO"", ""CONTRATO"", ""CONTRATO"", ""CONTR."", _" & vbCrLf
    s = s & "                  ""TIPO DE OBRA"", ""TIPO"", ""APROVA" & ChrW(199) & "" & ChrW(195) & "O"", ""APROV."", _" & vbCrLf
    s = s & "                  ""PREVISTO"", ""PREV."", ""PREVISTA"", ""PREV."", ""ENTREGA"", ""ENTR."", ""IN" & ChrW(205) & "CIO"", ""IN" & ChrW(205) & "C."")" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = LBound(abrev) To UBound(abrev) - 1 Step 2" & vbCrLf
    s = s & "        s = Replace(s, abrev(i), abrev(i + 1))" & vbCrLf
    s = s & "        If Len(s) <= LIMITE Then RotuloCurto = s: Exit Function" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    '2) remove o texto entre par" & ChrW(234) & "nteses" & vbCrLf
    s = s & "    p1 = InStr(s, ""("")" & vbCrLf
    s = s & "    If p1 > 0 Then" & vbCrLf
    s = s & "        p2 = InStr(p1, s, "")"")" & vbCrLf
    s = s & "        If p2 = 0 Then p2 = Len(s)" & vbCrLf
    CodPainel28 = s
End Function

Private Function CodPainel29() As String
    Dim s As String
    s = s & "        s = Trim$(Left$(s, p1 - 1) & "" "" & Mid$(s, p2 + 1))" & vbCrLf
    s = s & "        If Len(s) <= LIMITE Then RotuloCurto = s: Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    '3) corta" & vbCrLf
    s = s & "    RotuloCurto = Left$(s, LIMITE - 1) & "".""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' CALEND" & ChrW(193) & "RIO, FERIADOS E MARCADOR (instalado por mInstaladorPainelObras)" & vbCrLf
    s = s & "'   - data digitada em fim de semana ou feriado n" & ChrW(227) & "o " & ChrW(233) & " aceita ao salvar" & vbCrLf
    s = s & "'   - datas que v" & ChrW(234) & "m da f" & ChrW(243) & "rmula (cronograma) s" & ChrW(243) & " recebem alerta" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "Private Sub ExtrasMontar()" & vbCrLf
    s = s & "    Const LARG_CAL As Single = 212" & vbCrLf
    s = s & "    Dim nm As Variant" & vbCrLf
    s = s & "    Dim c As Object" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    'abre espa" & ChrW(231) & "o para o calend" & ChrW(225) & "rio: a coluna da direita anda para a direita" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    For Each nm In Array(""lblSecHist"", ""lblSecNova"", ""lblDicaNova"")" & vbCrLf
    s = s & "        Me.Controls(CStr(nm)).Left = Me.Controls(CStr(nm)).Left + LARG_CAL" & vbCrLf
    s = s & "    Next nm" & vbCrLf
    s = s & "    For Each nm In Array(""txtHist"", ""txtNova"")" & vbCrLf
    s = s & "        Me.Controls(CStr(nm)).Left = Me.Controls(CStr(nm)).Left + LARG_CAL" & vbCrLf
    s = s & "        Me.Controls(CStr(nm)).Width = Me.Controls(CStr(nm)).Width - LARG_CAL" & vbCrLf
    s = s & "    Next nm" & vbCrLf
    s = s & "    'bot" & ChrW(245) & "es de baixo descem para a faixa de avisos" & vbCrLf
    s = s & "    For Each c In Me.Controls" & vbCrLf
    s = s & "        If c.Parent.Name = Me.Name Then" & vbCrLf
    s = s & "            If c.Top >= 448 Then c.Top = c.Top + 26" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next c" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    NovoLabel Me, ""lblSecCal"", ""CALEND" & ChrW(193) & "RIO"", 354, 100, 200, 17, 10.5, True, COR_ESCURO" & vbCrLf
    s = s & "    Set mCal = New cCalendario" & vbCrLf
    s = s & "    mCal.Criar Me, 354, 118, 202, 12, 452, LARG - 24" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    '---- marcador" & vbCrLf
    s = s & "    Set btnMarcar = NovoBotao(""cmdMarcar"", ""MARCAR P/ ANALISAR"", LARG - 264, 97, 126, 20)" & vbCrLf
    s = s & "    btnMarcar.Font.Size = 9" & vbCrLf
    s = s & "    btnMarcar.ControlTipText = ""Marca esta obra para analisar depois (marca-p" & ChrW(225) & "gina)""" & vbCrLf
    CodPainel29 = s
End Function

Private Function CodPainel30() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "    chkVisiveis.Caption = ""Vis" & ChrW(237) & "veis""" & vbCrLf
    s = s & "    chkVisiveis.Width = 58" & vbCrLf
    s = s & "    Set chkMarcadas = Me.Controls.Add(""Forms.CheckBox.1"", ""chkMarcadas"", True)" & vbCrLf
    s = s & "    chkMarcadas.Left = chkVisiveis.Left + 60" & vbCrLf
    s = s & "    chkMarcadas.Top = chkVisiveis.Top" & vbCrLf
    s = s & "    chkMarcadas.Width = 58" & vbCrLf
    s = s & "    chkMarcadas.Height = chkVisiveis.Height" & vbCrLf
    s = s & "    chkMarcadas.Caption = ""Marcadas""" & vbCrLf
    s = s & "    chkMarcadas.Font.Name = mFonte" & vbCrLf
    s = s & "    chkMarcadas.Font.Size = 9" & vbCrLf
    s = s & "    chkMarcadas.Font.Bold = True" & vbCrLf
    s = s & "    chkMarcadas.BackStyle = fmBackStyleTransparent" & vbCrLf
    s = s & "    chkMarcadas.ControlTipText = ""Mostra s" & ChrW(243) & " as obras marcadas para analisar""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set lblMcMedida = NovoLabel(Me, ""lblMcMedida"", """", -3000, 16, 10, 26, 17, True, vbWhite)" & vbCrLf
    s = s & "    lblMcMedida.Font.Name = mFonteTitulo" & vbCrLf
    s = s & "    lblMcMedida.Font.Bold = mTituloNegrito" & vbCrLf
    s = s & "    lblMcMedida.WordWrap = False" & vbCrLf
    s = s & "    lblMcMedida.AutoSize = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set lblMcFundo = NovoLabel(Me, ""lblMcFundo"", """", LARG - 404, 5, 396, 51)" & vbCrLf
    s = s & "    lblMcFundo.BackStyle = fmBackStyleOpaque" & vbCrLf
    s = s & "    lblMcFundo.BackColor = MC_NAVY2" & vbCrLf
    s = s & "    Set lblMcBarra = NovoLabel(Me, ""lblMcBarra"", """", LARG - 404, 5, 4, 51)" & vbCrLf
    s = s & "    lblMcBarra.BackStyle = fmBackStyleOpaque" & vbCrLf
    s = s & "    lblMcBarra.BackColor = MC_AMARELO" & vbCrLf
    s = s & "    Set lblMcTit = NovoLabel(Me, ""lblMcTit"", """", LARG - 392, 8, 250, 14, 9, True, MC_AMARELO)" & vbCrLf
    s = s & "    lblMcTit.WordWrap = False" & vbCrLf
    s = s & "    Set lblMcPrazo = NovoLabel(Me, ""lblMcPrazo"", """", LARG - 140, 8, 128, 14, 8, True, COR_ESCURO)" & vbCrLf
    s = s & "    lblMcPrazo.BackStyle = fmBackStyleOpaque" & vbCrLf
    s = s & "    lblMcPrazo.BackColor = MC_AMARELO" & vbCrLf
    s = s & "    lblMcPrazo.TextAlign = fmTextAlignCenter" & vbCrLf
    s = s & "    lblMcPrazo.WordWrap = False" & vbCrLf
    s = s & "    Set lblMcTexto = NovoLabel(Me, ""lblMcTexto"", """", LARG - 392, 23, 380, 15, 9.5, False, vbWhite)" & vbCrLf
    s = s & "    lblMcTexto.WordWrap = False" & vbCrLf
    s = s & "    Set lblMcPor = NovoLabel(Me, ""lblMcPor"", """", LARG - 392, 39, 380, 13, 8, False, MC_TEXTO_CLARO)" & vbCrLf
    s = s & "    lblMcPor.WordWrap = False" & vbCrLf
    s = s & "    Set lblMcClique = NovoLabel(Me, ""lblMcClique"", """", LARG - 404, 5, 396, 51)" & vbCrLf
    s = s & "    lblMcClique.ControlTipText = ""Clique para ver todas as obras marcadas""" & vbCrLf
    s = s & "    Set mMcIcone = New cMarcaIcone" & vbCrLf
    s = s & "    mMcIcone.Criar Me, ""icoMarca"", 0, 20, 1.2, MC_AMARELO" & vbCrLf
    s = s & "    mMcIcone.Mostrar False" & vbCrLf
    s = s & "    MostrarCardMarcador False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodPainel30 = s
End Function

Private Function CodPainel31() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "'Chamado ao fim de CarregarCampos" & vbCrLf
    s = s & "Private Sub ExtrasAtualizar()" & vbCrLf
    s = s & "    FerAtualizarObra" & vbCrLf
    s = s & "    AtualizarMarcadorPainel" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------- marcador" & vbCrLf
    s = s & "Private Function MarcasPainel() As cMarcadores" & vbCrLf
    s = s & "    If mMarcas Is Nothing Then Set mMarcas = New cMarcadores" & vbCrLf
    s = s & "    Set MarcasPainel = mMarcas" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnMarcar_Click()" & vbCrLf
    s = s & "    If mObra Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    If AbrirMarcador(mObra.CodObra, mObra.nome) <> 0 Then" & vbCrLf
    s = s & "        MarcasPainel().Recarregar" & vbCrLf
    s = s & "        If chkMarcadas.value = True Then FiltrarLista" & vbCrLf
    s = s & "        AtualizarMarcadorPainel" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub chkMarcadas_Click()" & vbCrLf
    s = s & "    If mCarregando Then Exit Sub" & vbCrLf
    s = s & "    MarcasPainel().Recarregar" & vbCrLf
    s = s & "    FiltrarLista" & vbCrLf
    s = s & "    If chkMarcadas.value = True And mQtd > 0 Then" & vbCrLf
    s = s & "        If PosicaoNaLista(mAtual) = 0 Then" & vbCrLf
    s = s & "            If ConfirmarPendencias() Then CarregarObra mMapa(1)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Clique no cart" & ChrW(227) & "o do topo: lista das obras marcadas" & vbCrLf
    s = s & "Private Sub lblMcClique_Click()" & vbCrLf
    s = s & "    Dim cod As String" & vbCrLf
    s = s & "    Dim idx As Long" & vbCrLf
    s = s & "    cod = AbrirListaMarcadas()" & vbCrLf
    s = s & "    MarcasPainel().Recarregar" & vbCrLf
    s = s & "    If chkMarcadas.value = True Then FiltrarLista" & vbCrLf
    s = s & "    If Len(cod) > 0 Then" & vbCrLf
    s = s & "        idx = MarcasPainel().LinhaNoFollowUp(cod)" & vbCrLf
    s = s & "        If idx > 0 Then" & vbCrLf
    s = s & "            IrParaLinha idx" & vbCrLf
    s = s & "        Else" & vbCrLf
    CodPainel31 = s
End Function

Private Function CodPainel32() As String
    Dim s As String
    s = s & "            AbrirObraMarcada cod" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    AtualizarMarcadorPainel" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub MostrarCardMarcador(ByVal ligado As Boolean)" & vbCrLf
    s = s & "    lblMcFundo.Visible = ligado" & vbCrLf
    s = s & "    lblMcBarra.Visible = ligado" & vbCrLf
    s = s & "    lblMcTit.Visible = ligado" & vbCrLf
    s = s & "    lblMcPrazo.Visible = ligado" & vbCrLf
    s = s & "    lblMcTexto.Visible = ligado" & vbCrLf
    s = s & "    lblMcPor.Visible = ligado" & vbCrLf
    s = s & "    lblMcClique.Visible = ligado" & vbCrLf
    s = s & "    If ligado Then lblMcClique.ZOrder 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub AtualizarMarcadorPainel()" & vbCrLf
    s = s & "    Dim texto As String" & vbCrLf
    s = s & "    Dim revisar As Variant" & vbCrLf
    s = s & "    Dim soEu As Boolean" & vbCrLf
    s = s & "    Dim por As String" & vbCrLf
    s = s & "    Dim em As Variant" & vbCrLf
    s = s & "    Dim total As Long" & vbCrLf
    s = s & "    Dim hoje As Long" & vbCrLf
    s = s & "    Dim atrasadas As Long" & vbCrLf
    s = s & "    Dim marcada As Boolean" & vbCrLf
    s = s & "    Dim atrasada As Boolean" & vbCrLf
    s = s & "    Dim s As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo Fim" & vbCrLf
    s = s & "    If lblMcFundo Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    If Not mObra Is Nothing Then marcada = MarcasPainel().Ler(mObra.CodObra, texto, revisar, soEu, por, em)" & vbCrLf
    s = s & "    atrasada = (TipoPrazo(revisar) = PZ_ATRASADA)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If marcada Then" & vbCrLf
    s = s & "        lblMcMedida.Font.Size = lblNome.Font.Size" & vbCrLf
    s = s & "        lblMcMedida.Caption = lblNome.Caption" & vbCrLf
    s = s & "        mMcIcone.Mover lblNome.Left + lblMcMedida.Width + 8, lblNome.Top + 5" & vbCrLf
    s = s & "        If atrasada Then" & vbCrLf
    s = s & "            mMcIcone.Cor MC_CORAL" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            mMcIcone.Cor MC_AMARELO" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        mMcIcone.Dica ""Marcada para analisar: "" & Left$(texto, 200)" & vbCrLf
    CodPainel32 = s
End Function

Private Function CodPainel33() As String
    Dim s As String
    s = s & "        mMcIcone.Frente" & vbCrLf
    s = s & "        mMcIcone.Mostrar True" & vbCrLf
    s = s & "        btnMarcar.Caption = ""EDITAR MARCADOR""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        mMcIcone.Mostrar False" & vbCrLf
    s = s & "        btnMarcar.Caption = ""MARCAR P/ ANALISAR""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    MarcasPainel().Resumo total, hoje, atrasadas" & vbCrLf
    s = s & "    If marcada Then" & vbCrLf
    s = s & "        lblMcTit.Caption = ""MARCADA PARA ANALISAR""" & vbCrLf
    s = s & "        lblMcPrazo.Caption = TextoPrazo(revisar)" & vbCrLf
    s = s & "        If atrasada Then" & vbCrLf
    s = s & "            lblMcPrazo.BackColor = MC_CORAL" & vbCrLf
    s = s & "            lblMcPrazo.ForeColor = vbWhite" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            lblMcPrazo.BackColor = MC_AMARELO" & vbCrLf
    s = s & "            lblMcPrazo.ForeColor = COR_ESCURO" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        lblMcTexto.Caption = ChrW(8220) & Replace(Replace(texto, vbCr, "" ""), vbLf, "" "") & ChrW(8221)" & vbCrLf
    s = s & "        lblMcTexto.ControlTipText = Left$(texto, 250)" & vbCrLf
    s = s & "        s = ""marcada por "" & por" & vbCrLf
    s = s & "        If IsDate(em) Then s = s & "" em "" & Format$(em, ""dd\/mm\/yyyy hh:mm"")" & vbCrLf
    s = s & "        If soEu Then s = s & ""  "" & ChrW(183) & ""  s" & ChrW(243) & " voc" & ChrW(234) & " v" & ChrW(234) & """" & vbCrLf
    s = s & "        s = s & ""   "" & ChrW(183) & ""   clique para ver as "" & total & "" marcada(s)""" & vbCrLf
    s = s & "        lblMcPor.Caption = s" & vbCrLf
    s = s & "        MostrarCardMarcador True" & vbCrLf
    s = s & "    ElseIf total > 0 Then" & vbCrLf
    s = s & "        lblMcTit.Caption = ""OBRAS MARCADAS PARA ANALISAR: "" & total" & vbCrLf
    s = s & "        If atrasadas > 0 Then" & vbCrLf
    s = s & "            lblMcPrazo.Caption = atrasadas & "" atrasada(s)""" & vbCrLf
    s = s & "            lblMcPrazo.BackColor = MC_CORAL" & vbCrLf
    s = s & "            lblMcPrazo.ForeColor = vbWhite" & vbCrLf
    s = s & "        ElseIf hoje > 0 Then" & vbCrLf
    s = s & "            lblMcPrazo.Caption = hoje & "" para hoje""" & vbCrLf
    s = s & "            lblMcPrazo.BackColor = MC_AMARELO" & vbCrLf
    s = s & "            lblMcPrazo.ForeColor = COR_ESCURO" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            lblMcPrazo.Caption = ""nenhuma para hoje""" & vbCrLf
    s = s & "            lblMcPrazo.BackColor = MC_CINZA_CLARO" & vbCrLf
    s = s & "            lblMcPrazo.ForeColor = COR_ESCURO" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        lblMcTexto.Caption = ""Esta obra n" & ChrW(227) & "o est" & ChrW(225) & " marcada.""" & vbCrLf
    s = s & "        lblMcTexto.ControlTipText = """"" & vbCrLf
    s = s & "        lblMcPor.Caption = ""Clique aqui para ver a lista e abrir uma obra marcada.""" & vbCrLf
    CodPainel33 = s
End Function

Private Function CodPainel34() As String
    Dim s As String
    s = s & "        MostrarCardMarcador True" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        MostrarCardMarcador False" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "Fim:" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------- feriados" & vbCrLf
    s = s & "Private Sub FerPrepararCampos()" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim ev As cFerCampo" & vbCrLf
    s = s & "    Dim t As MSForms.label" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    FeriadosRecarregar" & vbCrLf
    s = s & "    Set mFerCampos = New Collection" & vbCrLf
    s = s & "    mFerQtd = mQtdCampos" & vbCrLf
    s = s & "    If mFerQtd = 0 Then Exit Sub" & vbCrLf
    s = s & "    ReDim mFerTag(1 To mFerQtd)" & vbCrLf
    s = s & "    For i = 1 To mFerQtd" & vbCrLf
    s = s & "        Set t = fraInfo.Controls.Add(""Forms.Label.1"", ""lblFerTag"" & i, True)" & vbCrLf
    s = s & "        t.Left = mBoxEd(i).Left + mBoxEd(i).Width - 58" & vbCrLf
    s = s & "        t.Top = mBoxEd(i).Top + 4" & vbCrLf
    s = s & "        t.Width = 54" & vbCrLf
    s = s & "        t.Height = 13" & vbCrLf
    s = s & "        t.Font.Name = mFonte" & vbCrLf
    s = s & "        t.Font.Size = 7" & vbCrLf
    s = s & "        t.Font.Bold = True" & vbCrLf
    s = s & "        t.TextAlign = fmTextAlignCenter" & vbCrLf
    s = s & "        t.BackStyle = fmBackStyleOpaque" & vbCrLf
    s = s & "        t.Caption = """"" & vbCrLf
    s = s & "        t.Visible = False" & vbCrLf
    s = s & "        Set mFerTag(i) = t" & vbCrLf
    s = s & "        Set ev = New cFerCampo" & vbCrLf
    s = s & "        ev.Conectar mBoxEd(i), i, Me" & vbCrLf
    s = s & "        mFerCampos.Add ev" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub FerDesligar()" & vbCrLf
    s = s & "    Dim ev As cFerCampo" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    If Not mFerCampos Is Nothing Then" & vbCrLf
    s = s & "        For Each ev In mFerCampos" & vbCrLf
    s = s & "            ev.Desconectar" & vbCrLf
    s = s & "        Next ev" & vbCrLf
    CodPainel34 = s
End Function

Private Function CodPainel35() As String
    Dim s As String
    s = s & "    End If" & vbCrLf
    s = s & "    Set mFerCampos = Nothing" & vbCrLf
    s = s & "    If Not mCal Is Nothing Then mCal.Desligar" & vbCrLf
    s = s & "    Set mCal = Nothing" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Texto mostrado no campo i (caixa de edi" & ChrW(231) & "" & ChrW(227) & "o ou de consulta)" & vbCrLf
    s = s & "Private Function FerTexto(ByVal i As Long) As String" & vbCrLf
    s = s & "    If mEditavel(i) Then" & vbCrLf
    s = s & "        FerTexto = mBoxEd(i).Text" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        FerTexto = mBoxRo(i).Text" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FerData(ByVal i As Long, ByRef d As Date) As Boolean" & vbCrLf
    s = s & "    If i < 1 Or i > mFerQtd Then Exit Function" & vbCrLf
    s = s & "    If mCampoTipo(i) <> ""D"" Then Exit Function" & vbCrLf
    s = s & "    FerData = LerDataFer(FerTexto(i), d)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FerNome(ByVal i As Long) As String" & vbCrLf
    s = s & "    FerNome = mCampoRot(i)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FerUsar(ByVal d As Date) As String" & vbCrLf
    s = s & "    FerUsar = ""Usar "" & Format$(d, ""dd\/mm"") & "" ("" & DiaSemanaCurto(d) & "")""" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FerAlterado(ByVal i As Long) As Boolean" & vbCrLf
    s = s & "    If mEditavel(i) Then FerAlterado = (mBoxEd(i).Text <> mOrig(i))" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FerDatasDaObra() As Collection" & vbCrLf
    s = s & "    Dim col As New Collection" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim d As Date" & vbCrLf
    s = s & "    For i = 1 To mFerQtd" & vbCrLf
    s = s & "        If FerData(i, d) Then col.Add Array(d, FerNome(i))" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    Set FerDatasDaObra = col" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub FerBorda(ByVal i As Long, ByVal cor As Long)" & vbCrLf
    s = s & "    If cor < 0 Then" & vbCrLf
    CodPainel35 = s
End Function

Private Function CodPainel36() As String
    Dim s As String
    s = s & "        mBoxEd(i).BorderColor = COR_AMARELO" & vbCrLf
    s = s & "        mBoxRo(i).BorderColor = COR_BORDA" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        mBoxEd(i).BorderColor = cor" & vbCrLf
    s = s & "        mBoxRo(i).BorderColor = cor" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Etiqueta e borda do campo i" & vbCrLf
    s = s & "Private Sub FerMarcar(ByVal i As Long, ByVal info As Long, ByVal alterado As Boolean)" & vbCrLf
    s = s & "    If info = DIA_UTIL Then" & vbCrLf
    s = s & "        mFerTag(i).Visible = False" & vbCrLf
    s = s & "        FerBorda i, -1" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    mFerTag(i).Caption = EtiquetaDia(info)" & vbCrLf
    s = s & "    If Bloqueia(info) And alterado Then" & vbCrLf
    s = s & "        mFerTag(i).BackColor = FR_CORAL" & vbCrLf
    s = s & "        mFerTag(i).ForeColor = vbWhite" & vbCrLf
    s = s & "        FerBorda i, FR_CORAL" & vbCrLf
    s = s & "        mFerTag(i).ControlTipText = ""Data alterada em dia n" & ChrW(227) & "o " & ChrW(250) & "til: n" & ChrW(227) & "o ser" & ChrW(225) & " aceita ao salvar""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        mFerTag(i).BackColor = FR_AMARELO" & vbCrLf
    s = s & "        mFerTag(i).ForeColor = COR_ESCURO" & vbCrLf
    s = s & "        FerBorda i, -1" & vbCrLf
    s = s & "        If Not mEditavel(i) And Bloqueia(info) Then" & vbCrLf
    s = s & "            mFerTag(i).ControlTipText = ""Vem do Cronograma de Obras: ajuste a data l" & ChrW(225) & """" & vbCrLf
    s = s & "        ElseIf Bloqueia(info) Then" & vbCrLf
    s = s & "            mFerTag(i).ControlTipText = ""Data j" & ChrW(225) & " gravada em dia n" & ChrW(227) & "o " & ChrW(250) & "til: ajuste quando puder""" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            mFerTag(i).ControlTipText = ""Ponto facultativo: a data " & ChrW(233) & " aceita""" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    mFerTag(i).Visible = True" & vbCrLf
    s = s & "    mFerTag(i).ZOrder 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Confere o campo i; digitando = mostra o m" & ChrW(234) & "s no calend" & ChrW(225) & "rio e o aviso" & vbCrLf
    s = s & "Private Sub FerAvaliar(ByVal i As Long, ByVal digitando As Boolean)" & vbCrLf
    s = s & "    Dim d As Date" & vbCrLf
    s = s & "    Dim info As Long" & vbCrLf
    s = s & "    Dim nome As String" & vbCrLf
    s = s & "    Dim abr As String" & vbCrLf
    s = s & "    Dim alterado As Boolean" & vbCrLf
    s = s & "    Dim txt As String" & vbCrLf
    CodPainel36 = s
End Function

Private Function CodPainel37() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "    If i < 1 Or i > mFerQtd Then Exit Sub" & vbCrLf
    s = s & "    If Not FerData(i, d) Then" & vbCrLf
    s = s & "        mFerTag(i).Visible = False" & vbCrLf
    s = s & "        FerBorda i, -1" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    info = InfoDia(d, mFerUF, mFerCid, nome, abr)" & vbCrLf
    s = s & "    alterado = FerAlterado(i)" & vbCrLf
    s = s & "    FerMarcar i, info, alterado" & vbCrLf
    s = s & "    If Not digitando Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mFerIdx = i" & vbCrLf
    s = s & "    mFerUltimo = i" & vbCrLf
    s = s & "    mCal.Datas FerDatasDaObra()" & vbCrLf
    s = s & "    mCal.Mostrar Year(d), Month(d), d" & vbCrLf
    s = s & "    txt = FerNome(i) & "" em "" & Format$(d, ""dd\/mm\/yyyy"") & "" cai em "" & MotivoDia(info, nome, abr) & "".""" & vbCrLf
    s = s & "    If Bloqueia(info) Then" & vbCrLf
    s = s & "        mFerSug1 = DiaUtilAnterior(d, mFerUF, mFerCid)" & vbCrLf
    s = s & "        mFerSug2 = DiaUtilSeguinte(d, mFerUF, mFerCid)" & vbCrLf
    s = s & "        If alterado Then" & vbCrLf
    s = s & "            mCal.Aviso txt & "" A data n" & ChrW(227) & "o pode cair em fim de semana nem feriado."", 1, FerUsar(mFerSug1), FerUsar(mFerSug2)" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            mCal.Aviso txt & "" J" & ChrW(225) & " estava gravada assim; ajuste quando puder."", 2, FerUsar(mFerSug1), FerUsar(mFerSug2)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    ElseIf info = DIA_FACULTATIVO Then" & vbCrLf
    s = s & "        mCal.Aviso txt & "" A data " & ChrW(233) & " aceita; confira se haver" & ChrW(225) & " expediente."", 2" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        mCal.Aviso FerNome(i) & "": "" & Format$(d, ""dd\/mm\/yyyy"") & "" ("" & DiaSemanaCurto(d) & "") " & ChrW(233) & " dia " & ChrW(250) & "til."", 3" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub FerCampoMudou(ByVal i As Long)" & vbCrLf
    s = s & "    If mCarregando Then Exit Sub" & vbCrLf
    s = s & "    If mCal Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    If Not mEditavel(i) Then Exit Sub" & vbCrLf
    s = s & "    FerAvaliar i, True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub FerCampoFoco(ByVal i As Long)" & vbCrLf
    s = s & "    Dim d As Date" & vbCrLf
    s = s & "    If mCal Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    If i = mFerUltimo Then Exit Sub" & vbCrLf
    s = s & "    mFerUltimo = i" & vbCrLf
    s = s & "    If FerData(i, d) Then" & vbCrLf
    CodPainel37 = s
End Function

Private Function CodPainel38() As String
    Dim s As String
    s = s & "        mCal.Datas FerDatasDaObra()" & vbCrLf
    s = s & "        mCal.Mostrar Year(d), Month(d), d" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Campo de data edit" & ChrW(225) & "vel com o cursor (0 = nenhum)" & vbCrLf
    s = s & "Private Function FerCampoAtivo() As Long" & vbCrLf
    s = s & "    Dim c As Object" & vbCrLf
    s = s & "    Dim nm As String" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Set c = Me.ActiveControl" & vbCrLf
    s = s & "    If Not c Is Nothing Then" & vbCrLf
    s = s & "        If TypeName(c) = ""Frame"" Then Set c = c.ActiveControl" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not c Is Nothing Then nm = c.Name" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "    If Len(nm) > 0 Then" & vbCrLf
    s = s & "        For i = 1 To mFerQtd" & vbCrLf
    s = s & "            If mEditavel(i) And mBoxEd(i).Name = nm Then" & vbCrLf
    s = s & "                FerCampoAtivo = i" & vbCrLf
    s = s & "                Exit Function" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        Next i" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If mFerUltimo >= 1 And mFerUltimo <= mFerQtd Then" & vbCrLf
    s = s & "        If mEditavel(mFerUltimo) Then FerCampoAtivo = mFerUltimo" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub FerClique(ByVal chave As String)" & vbCrLf
    s = s & "    Dim d As Date" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim info As Long" & vbCrLf
    s = s & "    Dim nome As String" & vbCrLf
    s = s & "    Dim abr As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mCal Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    Select Case chave" & vbCrLf
    s = s & "        Case ""cal:prev""" & vbCrLf
    s = s & "            mCal.Mover -1" & vbCrLf
    s = s & "        Case ""cal:next""" & vbCrLf
    s = s & "            mCal.Mover 1" & vbCrLf
    s = s & "        Case ""aviso:1"", ""aviso:2""" & vbCrLf
    s = s & "            If mFerIdx < 1 Or mFerIdx > mFerQtd Then Exit Sub" & vbCrLf
    CodPainel38 = s
End Function

Private Function CodPainel39() As String
    Dim s As String
    s = s & "            If Not mEditavel(mFerIdx) Then" & vbCrLf
    s = s & "                mCal.Aviso FerNome(mFerIdx) & "" vem do Cronograma de Obras: ajuste a data l" & ChrW(225) & " e atualize os dados."", 2" & vbCrLf
    s = s & "                Exit Sub" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "            If chave = ""aviso:1"" Then" & vbCrLf
    s = s & "                d = mFerSug1" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                d = mFerSug2" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "            If d = 0 Then Exit Sub" & vbCrLf
    s = s & "            mBoxEd(mFerIdx).Text = Format$(d, ""dd\/mm\/yyyy"")" & vbCrLf
    s = s & "        Case Else" & vbCrLf
    s = s & "            If Left$(chave, 6) <> ""cal:c:"" Then Exit Sub" & vbCrLf
    s = s & "            d = mCal.DataDaCelula(CLng(Mid$(chave, 7)))" & vbCrLf
    s = s & "            If d = 0 Then Exit Sub" & vbCrLf
    s = s & "            info = InfoDia(d, mFerUF, mFerCid, nome, abr)" & vbCrLf
    s = s & "            If Bloqueia(info) Then" & vbCrLf
    s = s & "                mCal.Aviso Format$(d, ""dd\/mm\/yyyy"") & "" " & ChrW(233) & " "" & MotivoDia(info, nome, abr) & "": escolha um dia " & ChrW(250) & "til."", 1" & vbCrLf
    s = s & "                Exit Sub" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "            i = FerCampoAtivo()" & vbCrLf
    s = s & "            If i = 0 Then" & vbCrLf
    s = s & "                mCal.Aviso ""Clique primeiro num campo de data com borda amarela (por exemplo ALVAR" & ChrW(193) & ") e depois no dia do calend" & ChrW(225) & "rio."", 2" & vbCrLf
    s = s & "                Exit Sub" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "            mBoxEd(i).Text = Format$(d, ""dd\/mm\/yyyy"")" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub FerAtualizarObra()" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    If mCal Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    mFerUF = """"" & vbCrLf
    s = s & "    mFerCid = """"" & vbCrLf
    s = s & "    If Not mObra Is Nothing Then" & vbCrLf
    s = s & "        mFerUF = mObra.texto(PO_COL_UF)" & vbCrLf
    s = s & "        mFerCid = mObra.texto(PO_COL_CIDADE)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    mCal.Obra mFerUF, mFerCid" & vbCrLf
    s = s & "    For i = 1 To mFerQtd" & vbCrLf
    s = s & "        FerAvaliar i, False" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    mFerIdx = 0" & vbCrLf
    s = s & "    mFerUltimo = 0" & vbCrLf
    s = s & "    mFerSug1 = 0" & vbCrLf
    CodPainel39 = s
End Function

Private Function CodPainel40() As String
    Dim s As String
    s = s & "    mFerSug2 = 0" & vbCrLf
    s = s & "    mCal.Datas FerDatasDaObra()" & vbCrLf
    s = s & "    mCal.Mostrar Year(Date), Month(Date)" & vbCrLf
    s = s & "    mCal.Dica """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'Antes de salvar: data digitada em fim de semana ou feriado bloqueia" & vbCrLf
    s = s & "Private Function FerPodeSalvar() As Boolean" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim d As Date" & vbCrLf
    s = s & "    Dim info As Long" & vbCrLf
    s = s & "    Dim nome As String" & vbCrLf
    s = s & "    Dim abr As String" & vbCrLf
    s = s & "    Dim lista As String" & vbCrLf
    s = s & "    Dim primeiro As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    FerPodeSalvar = True" & vbCrLf
    s = s & "    If mCal Is Nothing Then Exit Function" & vbCrLf
    s = s & "    For i = 1 To mFerQtd" & vbCrLf
    s = s & "        If FerAlterado(i) Then" & vbCrLf
    s = s & "            If FerData(i, d) Then" & vbCrLf
    s = s & "                info = InfoDia(d, mFerUF, mFerCid, nome, abr)" & vbCrLf
    s = s & "                If Bloqueia(info) Then" & vbCrLf
    s = s & "                    If Len(lista) > 0 Then lista = lista & ""; """ & vbCrLf
    s = s & "                    lista = lista & FerNome(i) & "" "" & Format$(d, ""dd\/mm\/yyyy"") & "" ("" & MotivoDia(info, nome, abr) & "")""" & vbCrLf
    s = s & "                    If primeiro = 0 Then primeiro = i" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    If primeiro = 0 Then Exit Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    FerPodeSalvar = False" & vbCrLf
    s = s & "    FerAvaliar primeiro, True" & vbCrLf
    s = s & "    mCal.Aviso ""N" & ChrW(227) & "o salvei: "" & lista & "". Ajuste para um dia " & ChrW(250) & "til e salve de novo."", 1, FerUsar(mFerSug1), FerUsar(mFerSug2)" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' NOVA OBRA dentro do painel, conferida na base (mInstaladorNovaObraObras)" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "Private Sub btnNova_Click()" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    If Not ConfirmarPendencias() Then Exit Sub" & vbCrLf
    s = s & "    If fraNova Is Nothing Then" & vbCrLf
    s = s & "        Set fraNova = Me.Controls.Add(""Forms.Frame.1"", ""fraNova"", True)" & vbCrLf
    CodPainel40 = s
End Function

Private Function CodPainel41() As String
    Dim s As String
    s = s & "        fraNova.Left = 8" & vbCrLf
    s = s & "        fraNova.Top = 66" & vbCrLf
    s = s & "        fraNova.Width = LARG - 4" & vbCrLf
    s = s & "        fraNova.Height = ALT - 70" & vbCrLf
    s = s & "        fraNova.Caption = """"" & vbCrLf
    s = s & "        fraNova.SpecialEffect = fmSpecialEffectFlat" & vbCrLf
    s = s & "        fraNova.BorderStyle = fmBorderStyleSingle" & vbCrLf
    s = s & "        fraNova.BorderColor = COR_BORDA" & vbCrLf
    s = s & "        fraNova.BackColor = COR_FUNDO" & vbCrLf
    s = s & "        Set mNova = New cNovaObraPainel" & vbCrLf
    s = s & "        mNova.Criar fraNova" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        mNova.Limpar" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    fraNova.Visible = True" & vbCrLf
    s = s & "    fraNova.ZOrder 0" & vbCrLf
    s = s & "    mNova.Focar" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'A " & ChrW(225) & "rea NOVA OBRA terminou: mostra a obra inclu" & ChrW(237) & "da" & vbCrLf
    s = s & "Private Sub mNova_Fechar(ByVal linhaObras As Long)" & vbCrLf
    s = s & "    Dim idx As Long" & vbCrLf
    s = s & "    fraNova.Visible = False" & vbCrLf
    s = s & "    idx = linhaObras" & vbCrLf
    s = s & "    If idx > 0 And idx <= tb.ListRows.Count Then MostrarObra idx" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "' OBRA ENTREGUE dentro do painel (instalado por mInstaladorEntregueObras)" & vbCrLf
    s = s & "'====================================================================" & vbCrLf
    s = s & "Private Sub btnEntrega_Click()" & vbCrLf
    s = s & "    If tb Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    If mObra Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    If Not ConfirmarPendencias() Then Exit Sub" & vbCrLf
    s = s & "    If fraEntrega Is Nothing Then" & vbCrLf
    s = s & "        Set fraEntrega = Me.Controls.Add(""Forms.Frame.1"", ""fraEntrega"", True)" & vbCrLf
    s = s & "        fraEntrega.Left = 8" & vbCrLf
    s = s & "        fraEntrega.Top = 66" & vbCrLf
    s = s & "        fraEntrega.Width = LARG - 4" & vbCrLf
    s = s & "        fraEntrega.Height = ALT - 70" & vbCrLf
    s = s & "        fraEntrega.Caption = """"" & vbCrLf
    s = s & "        fraEntrega.SpecialEffect = fmSpecialEffectFlat" & vbCrLf
    s = s & "        fraEntrega.BorderStyle = fmBorderStyleSingle" & vbCrLf
    s = s & "        fraEntrega.BorderColor = COR_BORDA" & vbCrLf
    s = s & "        fraEntrega.BackColor = COR_FUNDO" & vbCrLf
    CodPainel41 = s
End Function

Private Function CodPainel42() As String
    Dim s As String
    s = s & "        Set mEntrega = New cEntregaPainel" & vbCrLf
    s = s & "        mEntrega.Criar fraEntrega" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    mEntrega.Preparar mObra.CodObra" & vbCrLf
    s = s & "    fraEntrega.Visible = True" & vbCrLf
    s = s & "    fraEntrega.ZOrder 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'A " & ChrW(225) & "rea OBRA ENTREGUE terminou: se a obra saiu do follow up, recarrega o painel" & vbCrLf
    s = s & "Private Sub mEntrega_Fechar(ByVal moveu As Boolean)" & vbCrLf
    s = s & "    Dim idx As Long" & vbCrLf
    s = s & "    fraEntrega.Visible = False" & vbCrLf
    s = s & "    If moveu Then" & vbCrLf
    s = s & "        idx = mAtual" & vbCrLf
    s = s & "        If idx > tb.ListRows.Count Then idx = tb.ListRows.Count" & vbCrLf
    s = s & "        Recarregar idx" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodPainel42 = s
End Function

Private Function CodigoSemente() As String
    Dim s As String
    s = s & CodSemente1()
    CodigoSemente = s
End Function

Private Function CodSemente1() As String
    Dim s As String
    s = s & "09/07;Revolu" & ChrW(231) & "" & ChrW(227) & "o Constitucionalista de 1932;ESTADUAL;SP;;" & vbCrLf
    s = s & "CORPUS CHRISTI;Corpus Christi;MUNICIPAL;SP;BARUERI;" & vbCrLf
    s = s & "24/06;Padroeiro;MUNICIPAL;SP;BARUERI;" & vbCrLf
    s = s & "CORPUS CHRISTI;Corpus Christi;MUNICIPAL;SP;CAMPINAS;" & vbCrLf
    s = s & "12/08;Dia do Evang" & ChrW(233) & "lico;MUNICIPAL;SP;CAMPINAS;" & vbCrLf
    s = s & "08/12;Nossa Senhora da Concei" & ChrW(231) & "" & ChrW(227) & "o (padroeira);MUNICIPAL;SP;CAMPINAS;" & vbCrLf
    s = s & "CORPUS CHRISTI;Corpus Christi;MUNICIPAL;SP;GUARULHOS;" & vbCrLf
    s = s & "08/12;Nossa Senhora da Concei" & ChrW(231) & "" & ChrW(227) & "o / Anivers" & ChrW(225) & "rio;MUNICIPAL;SP;GUARULHOS;" & vbCrLf
    s = s & "CORPUS CHRISTI;Corpus Christi;MUNICIPAL;SP;JUNDIAI;" & vbCrLf
    s = s & "15/08;Nossa Senhora do Desterro (padroeira);MUNICIPAL;SP;JUNDIAI;" & vbCrLf
    s = s & "19/03;S" & ChrW(227) & "o Jos" & ChrW(233) & " (padroeiro);MUNICIPAL;SP;SAO JOSE DOS CAMPOS;" & vbCrLf
    s = s & "CORPUS CHRISTI;Corpus Christi;MUNICIPAL;SP;SAO JOSE DOS CAMPOS;" & vbCrLf
    s = s & "27/07;Anivers" & ChrW(225) & "rio da cidade;MUNICIPAL;SP;SAO JOSE DOS CAMPOS;" & vbCrLf
    s = s & "25/01;Anivers" & ChrW(225) & "rio da cidade;MUNICIPAL;SP;SAO PAULO;" & vbCrLf
    s = s & "CORPUS CHRISTI;Corpus Christi;MUNICIPAL;SP;SAO PAULO;" & vbCrLf
    s = s & "CORPUS CHRISTI;Corpus Christi;MUNICIPAL;SP;SOROCABA;" & vbCrLf
    s = s & "15/08;Nossa Senhora da Ponte (padroeira) / Anivers" & ChrW(225) & "rio;MUNICIPAL;SP;SOROCABA;" & vbCrLf
    s = s & "02/07;Independ" & ChrW(234) & "ncia da Bahia;ESTADUAL;BA;;" & vbCrLf
    s = s & "30/11;Dia do Evang" & ChrW(233) & "lico (Lei distrital 963/1995);ESTADUAL;DF;;" & vbCrLf
    s = s & "PASCOA+8;Nossa Senhora da Penha - Data Magna do ES;ESTADUAL;ES;;" & vbCrLf
    s = s & "28/10;Dia do Servidor P" & ChrW(250) & "blico (servidores estaduais);PONTO FACULTATIVO;GO;;s" & ChrW(243) & " para servidores estaduais; n" & ChrW(227) & "o bloqueia" & vbCrLf
    s = s & "06/03;Data Magna de Pernambuco;ESTADUAL;PE;;" & vbCrLf
    s = s & "CARNAVAL TERCA;Ter" & ChrW(231) & "a-feira de Carnaval (Lei 5.243/2008);ESTADUAL;RJ;;" & vbCrLf
    s = s & "23/04;Dia de S" & ChrW(227) & "o Jorge;ESTADUAL;RJ;;" & vbCrLf
    s = s & "CORPUS CHRISTI;Corpus Christi (Lei 11.002/2025);ESTADUAL;RJ;;" & vbCrLf
    s = s & "CORPUS CHRISTI;Corpus Christi;MUNICIPAL;MG;BELO HORIZONTE;" & vbCrLf
    s = s & "15/08;Assun" & ChrW(231) & "" & ChrW(227) & "o de Nossa Senhora;MUNICIPAL;MG;BELO HORIZONTE;" & vbCrLf
    s = s & "08/12;Imaculada Concei" & ChrW(231) & "" & ChrW(227) & "o;MUNICIPAL;MG;BELO HORIZONTE;" & vbCrLf
    s = s & "24/06;S" & ChrW(227) & "o Jo" & ChrW(227) & "o;MUNICIPAL;PE;RECIFE;" & vbCrLf
    s = s & "16/07;Nossa Senhora do Carmo (padroeira);MUNICIPAL;PE;RECIFE;" & vbCrLf
    s = s & "08/12;Nossa Senhora da Concei" & ChrW(231) & "" & ChrW(227) & "o;MUNICIPAL;PE;RECIFE;" & vbCrLf
    s = s & "CORPUS CHRISTI;Corpus Christi;MUNICIPAL;PR;CURITIBA;" & vbCrLf
    s = s & "08/09;Nossa Senhora da Luz dos Pinhais;MUNICIPAL;PR;CURITIBA;" & vbCrLf
    s = s & "20/01;S" & ChrW(227) & "o Sebasti" & ChrW(227) & "o;MUNICIPAL;RJ;RIO DE JANEIRO;" & vbCrLf
    CodSemente1 = s
End Function

