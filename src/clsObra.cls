Option Explicit

'====================================================================
' PAINEL - CLASSE OBRA
' Representa UMA linha da tb_Obras.
'
'   Dim obra As New clsObra
'   obra.Carregar tb, 7
'   obra.Nome / obra.CodObra / obra.Texto("CIDADE")
'   obra.AdicionarObservacao "texto"      -> grava HISTÓRICO + ÚLTIMO REGISTRO
'   obra.SubstituirHistorico "texto"      -> correção do histórico
'   obra.GravarData idxColuna, "15/11/2026"
'
' Toda gravação passa por PO_Gravar (log na tb_Alteracoes).
'====================================================================

Private mTb As ListObject
Private mIdx As Long
Private mLr As ListRow


'------------------------------------------------------------ carga
Public Sub Carregar(ByVal tabelaOrigem As ListObject, ByVal indiceLinha As Long)

    Set mTb = tabelaOrigem
    mIdx = indiceLinha
    Set mLr = mTb.ListRows(mIdx)

End Sub

Public Property Get Tabela() As ListObject
    Set Tabela = mTb
End Property

Public Property Get Indice() As Long
    Indice = mIdx
End Property

Public Property Get Linha() As ListRow
    Set Linha = mLr
End Property


'------------------------------------------------------------ leitura
Public Function valor(ByVal nomeColuna As String) As Variant
    valor = PO_Valor(mTb, mLr, nomeColuna)
End Function

Public Function texto(ByVal nomeColuna As String) As String
    texto = PO_TextoCampo(mTb, mLr, nomeColuna)
End Function

Public Function TextoColuna(ByVal idxColuna As Long) As String
    TextoColuna = PO_TextoValor(mLr.Range.Cells(1, idxColuna).value)
End Function

Public Property Get nome() As String
    nome = texto(PO_COL_DESCRICAO)
End Property

Public Property Get CodObra() As String
    CodObra = texto(PO_COL_CODIGO)
End Property

Public Property Get rotulo() As String
    rotulo = CodObra & " - " & nome
End Property

Public Property Get Historico() As String
    Historico = texto(PO_COL_OBS)
End Property

Public Property Get UltimaObservacao() As String
    UltimaObservacao = texto(PO_COL_ULTIMA)
End Property

Public Property Get Visivel() As Boolean
    Visivel = Not mLr.Range.EntireRow.Hidden
End Property

'"CÓDIGO 2801 | CLIENTE ATLAS | CAMPINAS/SP | CONSTRUÇÃO | Etapa: ... | Situação: ..."
Public Property Get Resumo() As String

    Dim s As String, v As String

    s = "CÓDIGO " & CodObra
    v = texto("CLIENTE"): If v <> "" Then s = s & "   |   " & v
    v = texto(PO_COL_CIDADE)
    If v <> "" Then
        s = s & "   |   " & v
        If texto(PO_COL_UF) <> "" Then s = s & "/" & texto(PO_COL_UF)
    End If
    v = texto(PO_COL_MODALIDADE): If v <> "" Then s = s & "   |   " & v
    v = texto(PO_COL_FASE):       If v <> "" Then s = s & "   |   Etapa: " & v
    v = texto(PO_COL_STATUS):     If v <> "" Then s = s & "   |   Situação: " & v

    Resumo = s

End Property

'Entradas do histórico: cada item = Array(data|Empty, rótulo, texto)
Public Function Observacoes() As Collection
    Set Observacoes = PO_ParseObs(Historico)
End Function


'------------------------------------------------------------ gravação
'Retorno: 1 alterou | 0 igual | -1 falhou
Public Function GravarColuna(ByVal idxColuna As Long, ByVal novo As Variant) As Long
    GravarColuna = PO_Gravar(mTb, mLr, idxColuna, novo)
End Function

'Texto dd/mm/aaaa (ou vazio para limpar). -2 = data inválida
Public Function GravarData(ByVal idxColuna As Long, ByVal textoData As String) As Long

    Dim d As Date, vz As Boolean

    If Not PO_TextoParaData(textoData, d, vz) Then
        GravarData = -2
    ElseIf vz Then
        GravarData = GravarColuna(idxColuna, Empty)
    Else
        GravarData = GravarColuna(idxColuna, d)
    End If

End Function

'Acrescenta "dd/mm/aaaa hh:mm - texto" ao histórico e à última atualização.
'historicoBase: se informado, é usado no lugar do histórico atual (edição manual).
Public Function AdicionarObservacao(ByVal textoNovo As String, _
                                    Optional ByVal historicoBase As Variant) As Boolean

    Dim idxObs As Long, idxUlt As Long
    Dim base As String, entrada As String

    idxObs = PO_ColIdx(mTb, PO_COL_OBS)
    idxUlt = PO_ColIdx(mTb, PO_COL_ULTIMA)
    If idxObs = 0 Or idxUlt = 0 Then Exit Function

    textoNovo = Trim$(PO_TrimFim(PO_ParaCaixa(textoNovo)))
    If textoNovo = "" Then Exit Function

    If IsMissing(historicoBase) Then
        base = Historico
    Else
        base = CStr(historicoBase)
    End If

    entrada = Format$(Now, "dd\/mm\/yyyy hh:mm") & " - " & textoNovo

    base = PO_TrimFim(base)
    If base <> "" Then base = base & vbCrLf & String(PO_TAM_SEPARADOR, "-") & vbCrLf
    base = base & entrada

    If GravarColuna(idxObs, base) < 0 Then Exit Function
    If GravarColuna(idxUlt, entrada) < 0 Then Exit Function

    AdicionarObservacao = True

End Function

'Substitui o histórico inteiro (correções feitas em "Editar histórico")
Public Function SubstituirHistorico(ByVal NovoTexto As String) As Boolean

    Dim idxObs As Long

    idxObs = PO_ColIdx(mTb, PO_COL_OBS)
    If idxObs = 0 Then Exit Function

    SubstituirHistorico = (GravarColuna(idxObs, PO_TrimFim(PO_ParaCaixa(NovoTexto))) >= 0)

End Function

'Seleciona a linha da obra na planilha
Public Sub IrParaPlanilha()

    On Error Resume Next
    mTb.Parent.Activate
    mLr.Range.Cells(1, 1).Select
    ActiveWindow.ScrollRow = mLr.Range.row

End Sub

