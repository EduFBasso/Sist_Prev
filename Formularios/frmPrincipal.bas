Option Explicit

Private Sub cmdBuscarCliente_Click()
    frmBusca.Show
End Sub

Private Sub cmdCadastro_Click()
    frmCadastro.NovoCliente
    frmCadastro.Show
End Sub

Private Sub cmdDocumentos_Click()
    frmDocumentos.Show
End Sub

Private Sub cmdSimulacoes_Click()
    frmSimulacoes.Show
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub
