Option Explicit
'===============================================================================
' cMarcaClique - repassa clique e duplo clique de um rótulo criado por código:
' dono.TratarClique chave   /   dono.TratarDuploClique chave
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
    If Not mDono Is Nothing Then mDono.TratarClique mChave
End Sub

Private Sub mRotulo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mDono Is Nothing Then mDono.TratarDuploClique mChave
End Sub
