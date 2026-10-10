Option Explicit
'===============================================================================
' mMarcador - marcador "para analisar" das obras (marca-página)
'
' Entradas:
'   AbrirMarcadas          - lista das obras marcadas (abre a obra escolhida
'                            no Painel de Obras)
'   AbrirMarcador cod, obra - janela para marcar / editar / concluir
'   AbrirListaMarcadas     - lista; devolve o CÓDIGO escolhido ("" = nenhum)
'
' Os marcadores ficam na aba oculta MARCADORES (tabela tb_Marcadores).
'===============================================================================

Public Const MC_NAVY As Long = 3877150
Public Const MC_NAVY2 As Long = 5587251
Public Const MC_AZUL As Long = 8950797
Public Const MC_AZULCLARO As Long = 15858636
Public Const MC_CORAL As Long = 2500316
Public Const MC_AMARELO As Long = 2408443
Public Const MC_CINZA As Long = 9139300
Public Const MC_CINZA_CLARO As Long = 12100500
Public Const MC_TEXTO_CLARO As Long = 14800331
Public Const MC_FUNDO As Long = 16579320
Public Const MC_BORDA As Long = 15788258
Public Const MC_BRANCO As Long = 16777215
Public Const MC_DESTAQUE As Long = 15465471
Public Const MC_CORAL_FUNDO As Long = 14869246
Public Const MC_EDICAO As Long = 15465471
Public Const MC_SELECAO As Long = 15858636

' situação do prazo
Public Const PZ_SEM_DATA As Long = 0
Public Const PZ_FUTURA As Long = 1
Public Const PZ_HOJE As Long = 2
Public Const PZ_ATRASADA As Long = 3

Private mFonteTexto As String
Private mFonteTitulo As String

'-------------------------------------------------------------------------------
' Entradas
'-------------------------------------------------------------------------------
' Devolve 0 (nada mudou), 1 (marcador salvo) ou 2 (análise concluída).
Public Function AbrirMarcador(ByVal codObra As String, ByVal obra As String) As Long
    frmMarcador.Preparar codObra, obra
    frmMarcador.Show vbModal
    AbrirMarcador = frmMarcador.Resultado
    Unload frmMarcador
End Function

' Devolve o CÓDIGO da obra escolhida para abrir ("" = nenhuma).
Public Function AbrirListaMarcadas() As String
    frmMarcadas.Preparar
    frmMarcadas.Show vbModal
    AbrirListaMarcadas = frmMarcadas.Escolhida
    Unload frmMarcadas
End Function

' Macro: lista das obras marcadas; a escolhida abre no painel dela.
Public Sub AbrirMarcadas()
    Dim cod As String
    cod = AbrirListaMarcadas()
    If Len(cod) > 0 Then AbrirObraMarcada cod
End Sub

' Abre a obra no Painel de Obras.
Public Sub AbrirObraMarcada(ByVal codObra As String)
    Dim marcas As cMarcadores
    Dim idx As Long
    Dim f As Object
    Dim painel As Object

    Set marcas = New cMarcadores
    idx = marcas.LinhaNoFollowUp(codObra)
    If idx = 0 Then Exit Sub
    For Each f In VBA.UserForms
        If f.Name = "frmPainelObras" Then Set painel = f
    Next f
    If painel Is Nothing Then
        On Error Resume Next
        Set painel = VBA.UserForms.Add("frmPainelObras")
        On Error GoTo 0
    End If
    If painel Is Nothing Then Exit Sub
    painel.IrParaLinha idx
    If Not painel.Visible Then painel.Show vbModeless
End Sub

' Usada pelo instalador: cria a aba oculta MARCADORES (se ainda não existir).
Public Function PrepararMarcadores() As Boolean
    Dim marcas As cMarcadores
    Set marcas = New cMarcadores
    PrepararMarcadores = Not (marcas.Tabela(True) Is Nothing)
End Function

'-------------------------------------------------------------------------------
' Prazos
'-------------------------------------------------------------------------------
Public Function TipoPrazo(ByVal revisar As Variant) As Long
    If Not IsDate(revisar) Then
        TipoPrazo = PZ_SEM_DATA
    ElseIf Int(CDbl(CDate(revisar))) < Int(CDbl(Date)) Then
        TipoPrazo = PZ_ATRASADA
    ElseIf Int(CDbl(CDate(revisar))) = Int(CDbl(Date)) Then
        TipoPrazo = PZ_HOJE
    Else
        TipoPrazo = PZ_FUTURA
    End If
End Function

' Texto curto do prazo: "revisar hoje · 08/10", "atrasada · 06/10" ...
Public Function TextoPrazo(ByVal revisar As Variant) As String
    Select Case TipoPrazo(revisar)
        Case PZ_ATRASADA
            TextoPrazo = "atrasada " & ChrW(183) & " " & Format$(CDate(revisar), "dd/mm")
        Case PZ_HOJE
            TextoPrazo = "revisar hoje " & ChrW(183) & " " & Format$(CDate(revisar), "dd/mm")
        Case PZ_FUTURA
            TextoPrazo = "revisar em " & Format$(CDate(revisar), "dd/mm")
        Case Else
            TextoPrazo = "sem data"
    End Select
End Function

Public Function CorPrazo(ByVal revisar As Variant) As Long
    Select Case TipoPrazo(revisar)
        Case PZ_ATRASADA
            CorPrazo = MC_CORAL
        Case PZ_HOJE
            CorPrazo = MC_AMARELO
        Case Else
            CorPrazo = MC_CINZA_CLARO
    End Select
End Function

Public Function DiaDaSemana(ByVal d As Date) As String
    DiaDaSemana = Choose(Weekday(d, vbSunday), "domingo", "segunda", "terça", "quarta", "quinta", "sexta", "sábado")
End Function

' "15/10/2026" ou "15102026" -> data; devolve False se não for data válida.
Public Function LerData(ByVal texto As String, ByRef d As Date) As Boolean
    Dim s As String
    Dim dd As Long
    Dim mm As Long
    Dim aa As Long
    s = Replace(Replace(Replace(Trim$(texto), "/", ""), "-", ""), ".", "")
    If Len(s) = 6 Then s = Left$(s, 4) & "20" & Right$(s, 2)
    If Len(s) <> 8 Then Exit Function
    If Not s Like "########" Then Exit Function
    dd = CLng(Left$(s, 2))
    mm = CLng(Mid$(s, 3, 2))
    aa = CLng(Right$(s, 4))
    If mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 Or aa < 2000 Or aa > 2100 Then Exit Function
    d = DateSerial(aa, mm, dd)
    If Day(d) <> dd Then Exit Function
    LerData = True
End Function

'-------------------------------------------------------------------------------
' Fontes da identidade visual (Segoe UI de reserva)
'-------------------------------------------------------------------------------
Public Function FonteMarcador() As String
    If Len(mFonteTexto) = 0 Then DetectarFontes
    FonteMarcador = mFonteTexto
End Function

Public Function FonteTituloMarcador() As String
    If Len(mFonteTitulo) = 0 Then DetectarFontes
    FonteTituloMarcador = mFonteTitulo
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
