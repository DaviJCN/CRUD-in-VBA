VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenuPrincipal 
   Caption         =   "CADASTRO DE MEMBROS"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "frmMenuPrincipal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMenuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnEstagiario_Click()
    Me.Hide
    frmEstagiario.Show
End Sub
 
Private Sub btnServidor_Click()
    Me.Hide
    frmServidor.Show
End Sub
 
Private Sub btnTerceirizado_Click()
    Me.Hide
    frmTerceirizado.Show
End Sub
 
Private Sub btnConselheiro_Click()
    Me.Hide
    frmConselheiro.Show
End Sub
 
Private Sub btnSair_Click()
    Unload Me
End Sub
 
Private Sub btnAjuda_Click()
    Dim resposta As VbMsgBoxResult
    Dim portalLink As String
 
    ' Definir o link para o Portal CRPS
    portalLink = "https://portalcrps.dataprev.gov.br/crps" ' Substitua por seu link real
 
    ' Exibe a mensagem de ajuda
    resposta = MsgBox("Prezado(a) Colega, aqui est�o algumas orienta��es �teis para utilizar o sistema:" & vbCrLf & vbCrLf & _
                       "1. Para cadastrar um novo estagi�rio, clique em 'Cadastrar' ap�s preencher todos os campos obrigat�rios." & vbCrLf & _
                       "2. Para remover um cadastro, digite apenas o CPF do estagi�rio no campo CPF e clique em 'Remover'." & vbCrLf & _
                       "3. O CPF ser� automaticamente formatado no padr�o correto ao sair do campo." & vbCrLf & _
                       "4. Para limpar os campos, clique em 'Cancelar'." & vbCrLf & _
                       "5. Lembre-se de que todos os campos obrigat�rios devem ser preenchidos antes de cadastrar." & vbCrLf & _
                       "Caso tenha d�vidas, entre em contato com a equipe de suporte." & vbCrLf & vbCrLf & _
                       "Deseja abrir um chamado para a equipe de Suporte TI?" & vbCrLf & vbCrLf & _
                       "Agradecemos a sua aten��o!" & vbCrLf & vbCrLf & _
                       "Atenciosamente," & vbCrLf & _
                       "Equipe Suporte TI.", vbYesNo + vbInformation, "Ajuda")
    ' Verifica a resposta do usu�rio
    If resposta = vbYes Then
        ' Se o usu�rio escolher "Sim", redireciona para o link
        ThisWorkbook.FollowHyperlink portalLink
    Else
        ' Se o usu�rio escolher "N�o", apenas fecha a mensagem de ajuda
        MsgBox "Caso precise de mais ajuda, entre em contato com a equipe de suporte.", vbInformation, "Ajuda"
    End If
End Sub
 
 
Private Sub UserForm_Click()
 
End Sub

