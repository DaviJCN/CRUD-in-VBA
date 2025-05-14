VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstagiario 
   Caption         =   "Cadastro De Estagiário"
   ClientHeight    =   10245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9705.001
   OleObjectBlob   =   "frmEstagiario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstagiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variável para controlar se o menu principal deve ser mostrado ao fechar
Dim mostrarMenu As Boolean
Dim linhaSelecionada As Long

' Função para verificar se o CPF já existe
Function CPFJaExiste(cpf As String, Optional ws As Worksheet = Nothing) As Boolean
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets("ESTAGIARIOS")
    End If
    Dim ultimaLinha As Long
    Dim i As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultimaLinha
        If ws.Cells(i, 7).Value = cpf Then
            CPFJaExiste = True
            Exit Function
        End If
    Next i
    CPFJaExiste = False
End Function

' Função para limpar os campos do formulário de estagiário
Sub LimparCamposEstagiario()
    txtEstagiario.Text = ""
    txtLotacaoEst.Text = ""
    txtMatriculaEst.Text = ""
    txtNascimentoEst.Text = ""
    txtIdadeEst.Text = ""
    txtCPFEst.Text = ""
    txtEmailInstitucional.Text = ""
    txtInicio.Text = ""
    txtFim.Text = ""
    txtUF.Text = ""
    txtCodigoUnidade.Text = ""
    optMasculinoEst.Value = False
    optFemininoEst.Value = False
    linhaSelecionada = 0
End Sub

' Função para formatar o CPF
Function FormatCPF(cpf As String) As String
    Dim digitsOnly As String
    Dim c As Integer
    For c = 1 To Len(cpf)
        If Mid(cpf, c, 1) Like "#" Then
            digitsOnly = digitsOnly & Mid(cpf, c, 1)
        End If
    Next c
    If Len(digitsOnly) = 11 Then
        FormatCPF = Format(digitsOnly, "000\.000\.000\-00")
    Else
        FormatCPF = cpf
    End If
    If Len(digitsOnly) <> 11 And Trim(cpf) <> "" Then
        MsgBox "CPF inválido! Digite 11 números.", vbCritical, "Erro de CPF"
        FormatCPF = cpf
        Exit Function
    End If
End Function

' Função para formatar a data de nascimento
Function FormatDataNascimento(Data As String) As String
    Dim dataFormatada As String
    Dim partes() As String
    Dim dia As String
    Dim mes As String
    Dim ano As String

    ' Remove espaços desnecessários e padroniza separadores
    Data = Trim(Replace(Replace(Data, "-", "/"), ".", "/"))

    ' Verifica se a data já está formatada corretamente
    If Data Like "##/##/####" Then
        FormatDataNascimento = Data
        Exit Function
    End If

    partes = Split(Data, "/")

    If UBound(partes) = 2 Then
        dia = partes(0)
        mes = partes(1)
        ano = partes(2)

        ' Garante que dia e mês tenham 2 dígitos e ano 4
        dia = Format(dia, "00")
        mes = Format(mes, "00")
        ano = Format(ano, "0000")

        dataFormatada = dia & "/" & mes & "/" & ano
        FormatDataNascimento = dataFormatada
    Else
        FormatDataNascimento = Data
    End If
End Function


' Função para calcular a idade
Function CalcularIdade(dataNascimento As String) As String
    On Error GoTo erro
    Dim dtNasc As Date
    dtNasc = CDate(dataNascimento)
    Dim hoje As Date
    hoje = Date
    Dim idade As Integer
    idade = DateDiff("yyyy", dtNasc, hoje)
    ' Ajusta se o aniversário ainda não chegou no ano atual
    If Month(hoje) < Month(dtNasc) Or (Month(hoje) = Month(dtNasc) And Day(hoje) < Day(dtNasc)) Then
        idade = idade - 1
    End If

    CalcularIdade = idade
    Exit Function
erro:
    CalcularIdade = ""
End Function

' Função para garantir que as primeiras letras do nome sejam maiúsculas
Function ProperCase(str As String) As String
    Dim i As Integer
    Dim resultado As String
    Dim palavras() As String
    ' Divide o nome em palavras
    palavras = Split(str)
    ' Converte a primeira letra de cada palavra para maiúscula
    For i = LBound(palavras) To UBound(palavras)
        palavras(i) = UCase(Left(palavras(i), 1)) & LCase(Mid(palavras(i), 2))
    Next i
    ' Junta as palavras novamente com um espaço
    resultado = Join(palavras, " ")
    ' Retorna o nome formatado
    ProperCase = resultado
End Function

' Função para procurar a linha pelo CPF
Function ProcurarLinhaPorCPF(cpf As String, Optional ws As Worksheet = Nothing) As Long
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets("ESTAGIARIOS") ' Ajustado para a planilha correta
    End If
    Dim linha As Long
    For linha = 2 To ws.Cells(ws.Rows.Count, 7).End(xlUp).Row
        If Trim(ws.Cells(linha, 7).Value) = Trim(cpf) Then
            ProcurarLinhaPorCPF = linha
            Exit Function
        End If
    Next linha
    ProcurarLinhaPorCPF = 0
End Function

' Função para validar o formato do e-mail
Function IsValidEmail(email As String) As Boolean
    Dim AtPos As Integer
    Dim DotPos As Integer
    Dim i As Integer
    Dim invalidChars As String

    ' Verifica se contém "@" e "."
    AtPos = InStr(1, email, "@")
    DotPos = InStrRev(email, ".")

    If AtPos < 2 Or AtPos > Len(email) - 3 Or DotPos <= AtPos + 1 Or DotPos = Len(email) Then
        IsValidEmail = False
        Exit Function
    End If

    ' Verifica por caracteres inválidos (espaços e alguns símbolos comuns)
    invalidChars = " ""(),:;<>\[]\"
    For i = 1 To Len(invalidChars)
        If InStr(1, email, Mid(invalidChars, i, 1)) > 0 Then
            IsValidEmail = False
            Exit Function
        End If
    Next i

    IsValidEmail = True
End Function

' Função para preencher os campos do formulário de estagiário
Sub PreencherFormularioEstagiario(linha As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ESTAGIARIOS")
    
    linhaSelecionada = linha
    
    txtEstagiario.Text = ws.Cells(linha, 1).Value
    txtLotacaoEst.Text = ws.Cells(linha, 2).Value
    txtMatriculaEst.Text = ws.Cells(linha, 3).Value
    txtNascimentoEst.Text = ws.Cells(linha, 4).Value
    txtIdadeEst.Text = ws.Cells(linha, 5).Value
    
    If ws.Cells(linha, 6).Value = "MASCULINO" Then
        optMasculinoEst.Value = True
        optFemininoEst.Value = False
    Else
        optMasculinoEst.Value = False
        optFemininoEst.Value = True
    End If
    
    txtCPFEst.Text = ws.Cells(linha, 7).Value
    txtEmailInstitucional.Text = ws.Cells(linha, 8).Value
    txtInicio.Text = ws.Cells(linha, 9).Value
    txtFim.Text = ws.Cells(linha, 10).Value
    txtUF.Text = ws.Cells(linha, 11).Value
    txtCodigoUnidade.Text = ws.Cells(linha, 12).Value
    
    botaoSalvarAlteracoes.Caption = "Atualizar"
End Sub


'=======================
'=== BOTÕES DO FORM ====
'=======================

' Evento para cadastrar um novo estagiário
Private Sub botaoCadastrarEstagiario_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ESTAGIARIOS")

    ' Validação de campos obrigatórios
    If Trim(txtEstagiario.Text) = "" Or Trim(txtLotacaoEst.Text) = "" Or Trim(txtMatriculaEst.Text) = "" Or _
       Trim(txtNascimentoEst.Text) = "" Or Trim(txtIdadeEst.Text) = "" Or _
       (Not optMasculinoEst.Value And Not optFemininoEst.Value) Or _
       Trim(txtCPFEst.Text) = "" Or Trim(txtEmailInstitucional.Text) = "" Or Trim(txtInicio.Text) = "" Or _
       Trim(txtFim.Text) = "" Or Trim(txtUF.Text) = "" Or Trim(txtCodigoUnidade.Text) = "" Then
        MsgBox "Preencha todos os campos obrigatórios!", vbCritical, "Campos Incompletos"
        Exit Sub
    End If

    ' Validação do formato do e-mail
    If Not IsValidEmail(Trim(txtEmailInstitucional.Text)) Then
        MsgBox "Formato de e-mail institucional inválido. Verifique se contém '@' e '.' e se não há caracteres inválidos.", vbExclamation, "Email Inválido"
        Exit Sub
    End If

    ' Verifica duplicidade de CPF
    If CPFJaExiste(FormatCPF(txtCPFEst.Text), ws) Then
        MsgBox "CPF já cadastrado!", vbExclamation, "CPF Duplicado"
        Exit Sub
    End If

    ' Determina o sexo
    Dim sexo As String
    If optMasculinoEst.Value Then
        sexo = "MASCULINO"
    ElseIf optFemininoEst.Value Then
        sexo = "FEMININO"
    Else
        MsgBox "Selecione o sexo do estagiário!", vbExclamation, "Sexo Não Selecionado"
        Exit Sub ' Adicionado para tratar o caso em que nenhum sexo é selecionado
    End If

    ' Cadastro de novo estagiário
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1 ' Encontra a próxima linha vazia
    
     ' Se estiver editando, usa a linha selecionada, senão, usa a última linha vazia
    If linhaSelecionada > 0 Then
        ultimaLinha = linhaSelecionada
    End If

    With ws
        .Cells(ultimaLinha, 1).Value = ProperCase(txtEstagiario.Text) ' Formata o nome
        .Cells(ultimaLinha, 2).Value = txtLotacaoEst.Text
        .Cells(ultimaLinha, 3).Value = txtMatriculaEst.Text
        .Cells(ultimaLinha, 4).Value = txtNascimentoEst.Text
        .Cells(ultimaLinha, 5).Value = txtIdadeEst.Text
        .Cells(ultimaLinha, 6).Value = sexo
        .Cells(ultimaLinha, 7).Value = FormatCPF(txtCPFEst.Text)
        .Cells(ultimaLinha, 8).Value = txtEmailInstitucional.Text
        .Cells(ultimaLinha, 9).Value = FormatDataNascimento(txtInicio.Text)
        .Cells(ultimaLinha, 10).Value = FormatDataNascimento(txtFim.Text)
        .Cells(ultimaLinha, 11).Value = txtUF.Text
        .Cells(ultimaLinha, 12).Value = txtCodigoUnidade.Text
    End With

    MsgBox "Estagiário cadastrado com sucesso!", vbInformation, "Sucesso"
    LimparCamposEstagiario
    botaoSalvarAlteracoes.Caption = "Salvar alterações"
End Sub

' Evento para remover um estagiário
Private Sub botaoremover_Click()
    Dim SenhaDigitada As String
    Dim SenhaCorreta As String

    SenhaCorreta = "123" ' Substitua pela senha desejada para remover

    SenhaDigitada = InputBox("Por favor, digite a senha para remover o estagiário:", "Autenticação Necessária")

    If SenhaDigitada = SenhaCorreta Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("ESTAGIARIOS")
        If Trim(txtCPFEst.Text) = "" Then
            MsgBox "Por favor, preencha o campo CPF para remover o cadastro.", vbExclamation, "CPF Necessário"
            Exit Sub
        End If

        Dim linha As Long
        linha = ProcurarLinhaPorCPF(FormatCPF(txtCPFEst.Text), ws) ' Passa a worksheet para a função

        If linha = 0 Then
            MsgBox "CPF não encontrado!", vbExclamation, "CPF Não Encontrado"
            Exit Sub
        End If

        Dim nome As String, cpf As String, email As String 'Declaração das variáveis
        nome = ws.Cells(linha, 1).Value
        cpf = ws.Cells(linha, 7).Value
        email = ws.Cells(linha, 8).Value
        matricula = ws.Cells(linha, 3).Value

        Dim mensagem As String
        mensagem = "Deseja realmente remover este cadastro?" & vbCrLf & vbCrLf & _
                    "Nome: " & nome & vbCrLf & _
                    "CPF: " & cpf & vbCrLf & _
                    "Email: " & email & vbCrLf & _
                    "Matricula : " & matricula

        If MsgBox(mensagem, vbYesNo + vbQuestion, "Confirmar exclusão") = vbYes Then
            ws.Rows(linha).Delete
            MsgBox "Cadastro removido com sucesso.", vbInformation, "Exclusão Concluída"
            LimparCamposEstagiario
        Else
            MsgBox "Exclusão cancelada.", vbInformation, "Exclusão Cancelada"
        End If
        botaoSalvarAlteracoes.Caption = "Salvar alterações"
    Else
        MsgBox "Senha incorreta. O estagiário não foi removido.", vbCritical
    End If
End Sub

' Evento para salvar alterações de um estagiário
Private Sub botaoSalvarAlteracoes_Click()
    Dim SenhaDigitada As String
    Dim SenhaCorreta As String

    SenhaCorreta = "123" ' Substitua pela senha desejada para salvar

    SenhaDigitada = InputBox("Por favor, digite a senha para salvar as alterações do estagiário:", "Autenticação Necessária")

    If SenhaDigitada = SenhaCorreta Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("ESTAGIARIOS")

        If Trim(txtCPFEst.Text) = "" Then
            MsgBox "Informe o CPF para salvar as alterações.", vbExclamation, "CPF Necessário"
            Exit Sub
        End If

        Dim linha As Long
        linha = ProcurarLinhaPorCPF(FormatCPF(txtCPFEst.Text), ws)

        If linha = 0 Then
            MsgBox "CPF não encontrado. Use 'Cadastrar' para novo registro.", vbExclamation, "CPF Não Encontrado"
            Exit Sub
        End If

         ' Verifica se os dados são os mesmos antes de atualizar
        If Trim(txtEstagiario.Text) = ws.Cells(linha, 1).Value And _
           Trim(txtLotacaoEst.Text) = ws.Cells(linha, 2).Value And _
           Trim(txtMatriculaEst.Text) = ws.Cells(linha, 3).Value And _
           Trim(txtNascimentoEst.Text) = ws.Cells(linha, 4).Value And _
           Trim(txtIdadeEst.Text) = ws.Cells(linha, 5).Value And _
           (IIf(optMasculinoEst.Value, "MASCULINO", "FEMININO") = ws.Cells(linha, 6).Value Or (Not optMasculinoEst.Value And Not optFemininoEst.Value)) And _
           FormatCPF(txtCPFEst.Text) = ws.Cells(linha, 7).Value And _
           Trim(txtEmailInstitucional.Text) = ws.Cells(linha, 8).Value And _
           FormatDataNascimento(txtInicio.Text) = ws.Cells(linha, 9).Value And _
           FormatDataNascimento(txtFim.Text) = ws.Cells(linha, 10).Value And _
           Trim(txtUF.Text) = ws.Cells(linha, 11).Value And _
           Trim(txtCodigoUnidade.Text) = ws.Cells(linha, 12).Value Then

            MsgBox "Os dados não foram alterados.", vbInformation
            Exit Sub ' Encerra a sub sem atualizar
        End If

        With ws
            ' Atualiza apenas se o campo de texto não estiver em branco
            If Trim(txtEstagiario.Text) <> "" Then
                .Cells(linha, 1).Value = ProperCase(txtEstagiario.Text)
            End If
            If Trim(txtLotacaoEst.Text) <> "" Then
                .Cells(linha, 2).Value = txtLotacaoEst.Text
            End If
            If Trim(txtMatriculaEst.Text) <> "" Then
                .Cells(linha, 3).Value = txtMatriculaEst.Text
            End If
            If Trim(txtNascimentoEst.Text) <> "" Then
                .Cells(linha, 4).Value = txtNascimentoEst.Text
            End If
            If Trim(txtIdadeEst.Text) <> "" Then
                .Cells(linha, 5).Value = txtIdadeEst.Text
            End If

            ' Atualiza o sexo se algum botão de opção estiver selecionado
            If optMasculinoEst.Value Or optFemininoEst.Value Then
                .Cells(linha, 6).Value = IIf(optMasculinoEst.Value, "MASCULINO", "FEMININO")
            End If

            If Trim(txtCPFEst.Text) <> "" Then
                .Cells(linha, 7).Value = FormatCPF(txtCPFEst.Text)
            End If
            If Trim(txtEmailInstitucional.Text) <> "" Then
                ' Validação do formato do e-mail ao salvar alterações
                If Not IsValidEmail(Trim(txtEmailInstitucional.Text)) Then
                    MsgBox "Formato de e-mail institucional inválido. Verifique se contém '@' e '.' e se não há caracteres inválidos.", vbExclamation, "Email Inválido"
                    Exit Sub
                End If
                .Cells(linha, 8).Value = txtEmailInstitucional.Text
            End If
            If Trim(txtInicio.Text) <> "" Then
                .Cells(linha, 9).Value = FormatDataNascimento(txtInicio.Text)
            End If
            If Trim(txtFim.Text) <> "" Then
                .Cells(linha, 10).Value = FormatDataNascimento(txtFim.Text)
            End If
            If Trim(txtUF.Text) <> "" Then
                .Cells(linha, 11).Value = txtUF.Text
            End If
            If Trim(txtCodigoUnidade.Text) <> "" Then
                .Cells(linha, 12).Value = txtCodigoUnidade.Text
            End If
        End With

        MsgBox "Alterações salvas com sucesso!", vbInformation, "Sucesso"
        LimparCamposEstagiario
        botaoSalvarAlteracoes.Caption = "Salvar alterações"
    Else
        MsgBox "Senha incorreta. As alterações não foram salvas.", vbCritical
    End If
End Sub

' Evento para voltar ao menu principal
Private Sub btnVoltar_Click()
    Unload Me
    frmMenuPrincipal.Show
End Sub

'===========================
'=== FORMATAR CAMPOS =======
'===========================
Private Sub txtCPFEst_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtCPFEst.Text = FormatCPF(txtCPFEst.Text)
    
     ' Se o CPF for encontrado, preenche os outros campos
    Dim linhaEncontradaPorCpf As Long
    linhaEncontradaPorCpf = ProcurarLinhaPorCPF(FormatCPF(txtCPFEst.Text))
    If linhaEncontradaPorCpf > 0 Then
        PreencherFormularioEstagiario linhaEncontradaPorCpf
        MsgBox "CPF preenchido automaticamente!", vbInformation, "Aviso" ' Exibe a mensagem
    ElseIf Trim(txtCPFEst.Text) <> "" Then
        MsgBox "CPF não encontrado!", vbExclamation, "Aviso"
    End If
End Sub

Private Sub txtNascimentoEst_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If IsDate(txtNascimentoEst.Text) Then
        txtNascimentoEst.BackColor = vbWhite
        txtNascimentoEst.Text = FormatDataNascimento(txtNascimentoEst.Text)
        txtIdadeEst.Text = CalcularIdade(txtNascimentoEst.Text)
    Else
        MsgBox "Data de nascimento inválida. Por favor, insira uma data válida no formato dd/mm/aaaa.", vbExclamation, "Data inválida"
        txtNascimentoEst.BackColor = RGB(255, 200, 200) ' Destaque em vermelho claro
        Cancel = True
    End If
End Sub

Private Sub txtInicio_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If IsDate(txtInicio.Text) Then
        txtInicio.Text = FormatDataNascimento(txtInicio.Text)
    Else
        MsgBox "Data de início inválida. Por favor, insira uma data válida no formato dd/mm/aaaa.", vbExclamation, "Data inválida"
        Cancel = True
    End If
End Sub

Private Sub txtFim_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If IsDate(txtFim.Text) Then
        txtFim.Text = FormatDataNascimento(txtFim.Text)
    Else
        MsgBox "Data de término inválida. Por favor, insira uma data válida no formato dd/mm/aaaa.", vbExclamation, "Data inválida"
        Cancel = True
    End If
End Sub

'==============================
'=== EVENTOS DE FECHAMENTO ====
'==============================

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Dim resposta As VbMsgBoxResult
        resposta = MsgBox("Tem certeza que deseja sair?", vbYesNo + vbQuestion, "Confirmar saída")
        If resposta = vbNo Then
            Cancel = True
        Else
            mostrarMenu = True ' sinaliza que deve mostrar o menu depois
        End If
    End If
End Sub

Private Sub UserForm_Terminate()
    If mostrarMenu Then
        frmMenuPrincipal.Show
    End If
End Sub


