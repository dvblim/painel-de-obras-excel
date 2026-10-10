Option Explicit
'===============================================================================
' mEntregues - obra entregue: do follow up para ENTREGUES
'
'   - área OBRA ENTREGUE dentro do Painel de Obras (cEntregaPainel)
'   - ArquivarEntregues: botão em lote da aba OBRAS (SITUAÇÃO = ENTREGUE),
'     agora copiando as colunas pelo nome
'===============================================================================

Public Const IA_NAVY As Long = 3877150
Public Const IA_AZUL As Long = 8950797
Public Const IA_CORAL As Long = 2500316
Public Const IA_AMARELO As Long = 2408443
Public Const IA_CINZA As Long = 9139300
Public Const IA_FUNDO As Long = 16579320
Public Const IA_BORDA As Long = 15788258
Public Const IA_BRANCO As Long = 16777215
Public Const IA_VERDE As Long = 4891414
Public Const IA_DESTAQUE As Long = 15465471
Public Const IA_VERDE_FUNDO As Long = 15595240
Public Const IA_CORAL_FUNDO As Long = 14869246

Private mFonteTexto As String
Private mFonteTitulo As String

' Botão em lote da aba OBRAS: todas as obras com SITUAÇÃO = ENTREGUE.
Public Sub ArquivarEntregues()
    Dim arq As cArquivoObras
    Dim lista As Collection
    Dim v As Variant
    Dim ok As Long
    Dim falhas As String

    Set arq = New cArquivoObras
    If Not arq.TemArquivo() Then
        MsgBox "A tabela de entregues (tb_Entregues) não foi encontrada.", vbExclamation, "Arquivar entregues"
        Exit Sub
    End If
    Set lista = arq.CodObraEntregues()
    If lista.Count = 0 Then
        MsgBox "Nenhuma obra do follow up está com SITUAÇÃO ENTREGUE." & vbCrLf & vbCrLf & "Para mover uma obra específica, use o botão OBRA ENTREGUE do Painel de Obras.", vbInformation, "Arquivar entregues"
        Exit Sub
    End If
    If MsgBox(lista.Count & " obra(s) com SITUAÇÃO ENTREGUE vão para ENTREGUES (colunas casadas pelo nome) e sairão do follow up." & vbCrLf & vbCrLf & "Continuar?", vbQuestion + vbYesNo, "Arquivar entregues") <> vbYes Then Exit Sub

    For Each v In lista
        If arq.Entregar(CStr(v)) Then
            ok = ok + 1
        Else
            falhas = falhas & vbCrLf & "CÓDIGO " & v & ": " & arq.UltimoErro
        End If
    Next v

    If Len(falhas) = 0 Then
        MsgBox ok & " obra(s) movida(s) para ENTREGUES." & vbCrLf & arq.AvisoFiltro, vbInformation, "Arquivar entregues"
    Else
        MsgBox ok & " obra(s) movida(s). Não foi possível mover:" & falhas, vbExclamation, "Arquivar entregues"
    End If
End Sub

Public Function FonteIA() As String
    If Len(mFonteTexto) = 0 Then DetectarFontes
    FonteIA = mFonteTexto
End Function

Public Function FonteTituloIA() As String
    If Len(mFonteTitulo) = 0 Then DetectarFontes
    FonteTituloIA = mFonteTitulo
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
