Option Explicit
'===============================================================================
' cMarcaIcone - desenha o marcador (fita de marca-página) com três rótulos:
' o corpo e duas "pernas" que deixam o recorte em V embaixo.
'===============================================================================

Private mCorpo As MSForms.Label
Private mPernaE As MSForms.Label
Private mPernaD As MSForms.Label
Private mK As Single

Public Sub Criar(ByVal pai As Object, ByVal nome As String, ByVal l As Single, ByVal t As Single, _
                 ByVal escala As Single, ByVal corInicial As Long)
    mK = escala
    Set mCorpo = NovaParte(pai, nome & "C")
    Set mPernaE = NovaParte(pai, nome & "E")
    Set mPernaD = NovaParte(pai, nome & "D")
    Mover l, t
    Cor corInicial
End Sub

Private Function NovaParte(ByVal pai As Object, ByVal nome As String) As MSForms.Label
    Dim c As MSForms.Label
    Set c = pai.Controls.Add("Forms.Label.1", nome, True)
    c.Caption = ""
    c.BackStyle = 1
    c.BorderStyle = 0
    Set NovaParte = c
End Function

' 11 x 15 pontos na escala 1
Public Sub Mover(ByVal l As Single, ByVal t As Single)
    If mCorpo Is Nothing Then Exit Sub
    mCorpo.Left = l
    mCorpo.Top = t
    mCorpo.Width = 11 * mK
    mCorpo.Height = 11 * mK
    mPernaE.Left = l
    mPernaE.Top = t + 11 * mK
    mPernaE.Width = 4 * mK
    mPernaE.Height = 4 * mK
    mPernaD.Left = l + 7 * mK
    mPernaD.Top = t + 11 * mK
    mPernaD.Width = 4 * mK
    mPernaD.Height = 4 * mK
End Sub

Public Sub Cor(ByVal valor As Long)
    If mCorpo Is Nothing Then Exit Sub
    mCorpo.BackColor = valor
    mPernaE.BackColor = valor
    mPernaD.BackColor = valor
End Sub

Public Sub Mostrar(ByVal ligado As Boolean)
    If mCorpo Is Nothing Then Exit Sub
    mCorpo.Visible = ligado
    mPernaE.Visible = ligado
    mPernaD.Visible = ligado
End Sub

Public Sub Dica(ByVal texto As String)
    If mCorpo Is Nothing Then Exit Sub
    mCorpo.ControlTipText = texto
    mPernaE.ControlTipText = texto
    mPernaD.ControlTipText = texto
End Sub

Public Sub Frente()
    If mCorpo Is Nothing Then Exit Sub
    mCorpo.ZOrder 0
    mPernaE.ZOrder 0
    mPernaD.ZOrder 0
End Sub

Public Property Get Corpo() As MSForms.Label
    Set Corpo = mCorpo
End Property
