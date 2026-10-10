Option Explicit
'===============================================================================
' cFerClique - repassa o clique de um rótulo do calendário: dono.FerClique chave
'===============================================================================

Private WithEvents mRotulo As MSForms.Label
Private mDono As Object
Private mChave As String

Public Sub Conectar(ByVal rotulo As MSForms.Label, ByVal chave As String, ByVal dono As Object)
    Set mRotulo = rotulo
    mChave = chave
    Set mDono = dono
End Sub

Public Sub Desconectar()
    Set mRotulo = Nothing
    Set mDono = Nothing
End Sub

Private Sub mRotulo_Click()
    If Not mDono Is Nothing Then mDono.FerClique mChave
End Sub
