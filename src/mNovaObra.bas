Option Explicit
'===============================================================================
' mNovaObra - apoio da área NOVA OBRA do Painel de Obras (cNovaObraPainel):
' inclusão de obra nova no follow up pelo CÓDIGO, com conferência
' nas bases de onde as fórmulas da tb_Obras trazem os dados:
'   1º tb_Cronograma (Cronograma de Obras)
'   2º tb_Carteira (aba CARTEIRA), quando não está no cronograma
' O CÓDIGO é gravado como texto (igual às bases); o resto vem das fórmulas.
'===============================================================================

Public Const NL_NAVY As Long = 3877150
Public Const NL_AZUL As Long = 8950797
Public Const NL_CORAL As Long = 2500316
Public Const NL_AMARELO As Long = 2408443
Public Const NL_CINZA As Long = 9139300
Public Const NL_FUNDO As Long = 16579320
Public Const NL_BORDA As Long = 15788258
Public Const NL_BRANCO As Long = 16777215
Public Const NL_VERDE As Long = 4891414
Public Const NL_DESTAQUE As Long = 15465471
Public Const NL_VERDE_FUNDO As Long = 15595240
Public Const NL_CORAL_FUNDO As Long = 14869246
Public Const NL_AZUL_FUNDO As Long = 16449008
Public Const NL_EDICAO As Long = 15465471
Public Const NL_DESLIGADO As Long = 12100500

Private mFonteTexto As String
Private mFonteTitulo As String

Public Function TabelaNL(ByVal nome As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        On Error Resume Next
        Set lo = ws.ListObjects(nome)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set TabelaNL = lo
            Exit Function
        End If
    Next ws
End Function

Public Function NormNL(ByVal s As String) As String
    s = Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), Chr$(160), " ")
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    NormNL = UCase$(Trim$(s))
End Function

Public Function ColNL(ByVal tb As ListObject, ByVal nome As String) As Long
    Dim i As Long
    If tb Is Nothing Then Exit Function
    For i = 1 To tb.ListColumns.Count
        If NormNL(tb.ListColumns(i).Name) = NormNL(nome) Then
            ColNL = i
            Exit Function
        End If
    Next i
End Function

Public Function TextoNL(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        TextoNL = Format$(v, "dd\/mm\/yyyy")
    Else
        TextoNL = Trim$(Replace(CStr(v), Chr$(160), " "))
    End If
End Function

' Linha da obra (CÓDIGO comparado como texto) na tabela; 0 = não está.
Public Function LinhaNaTabela(ByVal nome As String, ByVal codObra As String) As Long
    Dim tb As ListObject
    Dim c As Long
    Dim i As Long
    Dim v As Variant
    Set tb = TabelaNL(nome)
    If tb Is Nothing Then Exit Function
    If tb.DataBodyRange Is Nothing Then Exit Function
    c = ColNL(tb, "CÓDIGO")
    If c = 0 Then Exit Function
    codObra = Trim$(codObra)
    If Len(codObra) = 0 Then Exit Function
    v = tb.ListColumns(c).DataBodyRange.Value
    If Not IsArray(v) Then
        If TextoNL(v) = codObra Then LinhaNaTabela = 1
        Exit Function
    End If
    For i = 1 To UBound(v, 1)
        If TextoNL(v(i, 1)) = codObra Then
            LinhaNaTabela = i
            Exit Function
        End If
    Next i
End Function

' Valor da coluna na linha i da tabela (texto; "" se não houver).
Public Function ValorNL(ByVal nome As String, ByVal i As Long, ByVal coluna As String) As String
    Dim tb As ListObject
    Dim c As Long
    If i = 0 Then Exit Function
    Set tb = TabelaNL(nome)
    If tb Is Nothing Then Exit Function
    c = ColNL(tb, coluna)
    If c = 0 Then Exit Function
    ValorNL = TextoNL(tb.DataBodyRange.Cells(i, c).Value)
End Function

'-------------------------------------------------------------------------------
' Fontes da identidade visual (Segoe UI de reserva)
'-------------------------------------------------------------------------------
Public Function FonteNL() As String
    If Len(mFonteTexto) = 0 Then DetectarFontes
    FonteNL = mFonteTexto
End Function

Public Function FonteTituloNL() As String
    If Len(mFonteTitulo) = 0 Then DetectarFontes
    FonteTituloNL = mFonteTitulo
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
