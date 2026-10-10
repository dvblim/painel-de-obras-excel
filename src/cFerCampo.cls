Option Explicit
'===============================================================================
' cFerCampo - acompanha uma caixa de data do cronograma do painel:
' dono.FerCampoMudou indice (texto mudou) e dono.FerCampoFoco indice (clique)
'===============================================================================

Private WithEvents mCaixa As MSForms.TextBox
Private mDono As Object
Private mIndice As Long

Public Sub Conectar(ByVal caixa As MSForms.TextBox, ByVal indice As Long, ByVal dono As Object)
    Set mCaixa = caixa
    mIndice = indice
    Set mDono = dono
End Sub

Public Sub Desconectar()
    Set mCaixa = Nothing
    Set mDono = Nothing
End Sub

Private Sub mCaixa_Change()
    If Not mDono Is Nothing Then mDono.FerCampoMudou mIndice
End Sub

Private Sub mCaixa_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not mDono Is Nothing Then mDono.FerCampoFoco mIndice
End Sub

Private Sub mCaixa_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Not mDono Is Nothing Then mDono.FerCampoFoco mIndice
End Sub
