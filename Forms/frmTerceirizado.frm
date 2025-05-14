VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTerceirizado 
   Caption         =   "Cadastro De Terceirizado"
   ClientHeight    =   10200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9780.001
   OleObjectBlob   =   "frmTerceirizado.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTerceirizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variável para controlar se o menu principal deve ser mostrado ao fechar
Dim mostrarMenu As Boolean

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

' Função para formatar CPF
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
    If Len(digitsOnly) <> 11 And Trim(cpf) <> "" Then ' Adicionada condição para não mostrar mensagem se o CPF for vazio
        MsgBox "CPF inválido! Digite 11 números.", vbCritical ' Mensagem de erro mais clara
        FormatCPF = cpf
        Exit Function
    End If
End Function

' Função para colocar a primeira letra de cada palavra em maiúsculo
Function ProperCase(texto As String) As String
    ProperCase = Application.WorksheetFunction.Proper(texto)
End Function

' Verifica se o CPF já está cadastrado na aba "TERCEIRIZADOS"
Function CPFJaExiste(cpf As String) As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TERCEIRIZADOS")

    Dim celula As Range
    For Each celula In ws.Range("E2:E" & ws.Cells(ws.Rows.Count, 5).End(xlUp).Row)
        If Replace(Trim(celula.Value), ".", "") = Replace(Trim(cpf), ".", "") Then
            CPFJaExiste = True
            Exit Function
        End If
    Next celula
    CPFJaExiste = False
End Function

' Procura a linha de um CPF na aba "TERCEIRIZADOS"
Function ProcurarLinhaPorCPF(cpf As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TERCEIRIZADOS")

    Dim linha As Long
    For linha = 2 To ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
        If Replace(Trim(ws.Cells(linha, 5).Value), ".", "") = Replace(Trim(cpf), ".", "") Then
            ProcurarLinhaPorCPF = linha
            Exit Function
        End If
    Next linha
    ProcurarLinhaPorCPF = 0
End Function

' Função para formatar a data
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

' Limpa todos os campos do formulário
Sub LimparCamposTerceirizado()
    txtNomeTerceirizado.Text = ""
    txtEmailTerceirizado.Text = ""
    txtMatricula.Text = ""
    txtCPFTerceirizado.Text = ""
    txtSigla.Text = ""
    txtUnidade.Text = ""
    txtDataAdmissao.Text = ""
    txtCargo.Text = ""
    ' Desmarcar as opções de sexo
    optMasculinoTerce.Value = False
    optFemininoTerce.Value = False
    linhaSelecionada = 0
End Sub

' Preenche o formulário com os dados do terceirizado para edição
Sub PreencherFormularioTerceirizado(linha As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TERCEIRIZADOS")

    linhaSelecionada = linha 'armazena a linha

    txtNomeTerceirizado.Text = ws.Cells(linha, 1).Value

    If ws.Cells(linha, 2).Value = "MASCULINO" Then
        optMasculinoTerce.Value = True
        optFemininoTerce.Value = False
    Else
        optMasculinoTerce.Value = False
        optFemininoTerce.Value = True
    End If

    txtDataAdmissao.Text = ws.Cells(linha, 3).Value
    txtUnidade.Text = ws.Cells(linha, 4).Value
    txtCPFTerceirizado.Text = ws.Cells(linha, 5).Value
    txtCargo.Text = ws.Cells(linha, 6).Value
    txtMatricula.Text = ws.Cells(linha, 7).Value
    txtEmailTerceirizado.Text = ws.Cells(linha, 8).Value
    txtSigla.Text = ws.Cells(linha, 9).Value
End Sub

'=======================
'=== BOTÕES DO FORM ====
'=======================

Private Sub botaoCadastrarTerceirizado_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TERCEIRIZADOS")
    ' Validação de campos obrigatórios
    If Trim(txtNomeTerceirizado.Text) = "" Or _
       Trim(txtEmailTerceirizado.Text) = "" Or _
       Trim(txtMatricula.Text) = "" Or _
       Trim(txtCPFTerceirizado.Text) = "" Or _
       Trim(txtSigla.Text) = "" Or _
       Trim(txtUnidade.Text) = "" Or _
       Trim(txtDataAdmissao.Text) = "" Or _
       Trim(txtCargo.Text) = "" Or _
       (Not optMasculinoTerce.Value And Not optFemininoTerce.Value) Then ' Verifica se o sexo foi selecionado
        MsgBox "Preencha todos os campos obrigatórios, incluindo o sexo!", vbCritical
        Exit Sub
    End If
    ' Validação do formato do e-mail
    If Not IsValidEmail(Trim(txtEmailTerceirizado.Text)) Then
        MsgBox "Formato de e-mail inválido. Verifique se contém '@' e '.' e se não há caracteres inválidos.", vbExclamation
        Exit Sub
    End If
    ' Verifica duplicidade de CPF apenas se for um novo cadastro
    If linhaSelecionada = 0 And CPFJaExiste(FormatCPF(txtCPFTerceirizado.Text)) Then
        MsgBox "CPF já cadastrado!", vbExclamation
        Exit Sub
    End If

    ' Cadastro de novo terceirizado
    Dim ultimaLinha As Long
    ultimaLinha = 4 ' Começa a busca a partir da linha 4 (evita cabeçalho)

    Do While ws.Cells(ultimaLinha, 1).Value <> ""
        ultimaLinha = ultimaLinha + 1
    Loop

    ' Verifica o sexo e atribui o valor correspondente
    Dim sexo As String
    If optMasculinoTerce.Value Then
        sexo = "MASCULINO"
    ElseIf optFemininoTerce.Value Then
        sexo = "FEMININO"
    Else
        MsgBox "Selecione o sexo!", vbCritical
        Exit Sub
    End If

     ' Se estiver editando, usa a linha selecionada, senão, usa a última linha vazia
    If linhaSelecionada > 0 Then
        ultimaLinha = linhaSelecionada
    End If

    With ws
        .Cells(ultimaLinha, 1).Value = ProperCase(txtNomeTerceirizado.Text) ' Nome
        .Cells(ultimaLinha, 2).Value = sexo ' Sexo (MASCULINO ou FEMININO)
        .Cells(ultimaLinha, 3).Value = txtDataAdmissao.Text ' Data de Admissão
        .Cells(ultimaLinha, 4).Value = txtUnidade.Text ' Lotação
        .Cells(ultimaLinha, 5).Value = FormatCPF(txtCPFTerceirizado.Text) ' CPF
        .Cells(ultimaLinha, 6).Value = txtCargo.Text ' Cargo
        .Cells(ultimaLinha, 7).Value = txtMatricula.Text ' Matrícula
        .Cells(ultimaLinha, 8).Value = txtEmailTerceirizado.Text ' Email
        .Cells(ultimaLinha, 9).Value = txtSigla.Text ' Sigla
    End With

    If linhaSelecionada > 0 Then
        MsgBox "Terceirizado atualizado com sucesso!", vbInformation
    Else
        MsgBox "Terceirizado cadastrado com sucesso!", vbInformation
    End If
    LimparCamposTerceirizado
End Sub


Private Sub botaoremover_Click()
    Dim SenhaDigitada As String
    Dim SenhaCorreta As String

    SenhaCorreta = "123" ' Substitua pela senha desejada para remover

    SenhaDigitada = InputBox("Por favor, digite a senha para remover o terceirizado:", "Autenticação Necessária")

    If SenhaDigitada = SenhaCorreta Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("TERCEIRIZADOS") ' Corrigido para a planilha certa

        If Trim(txtCPFTerceirizado.Text) = "" Then
            MsgBox "Por favor, preencha o campo CPF para remover o cadastro.", vbExclamation
            Exit Sub
        End If

        Dim linha As Long
        linha = ProcurarLinhaPorCPF(FormatCPF(txtCPFTerceirizado.Text))

        If linha = 0 Then
            MsgBox "CPF não encontrado!", vbExclamation
            Exit Sub
        End If

        Dim nome As String, cpf As String, email As String, lotacao As String, cargo As String
        nome = ws.Cells(linha, 1).Value
        cpf = ws.Cells(linha, 5).Value
        email = ws.Cells(linha, 8).Value
        lotacao = ws.Cells(linha, 4).Value
        cargo = ws.Cells(linha, 6).Value

        Dim mensagem As String
        mensagem = "Deseja realmente remover este cadastro?" & vbCrLf & vbCrLf & _
                    "Nome: " & nome & vbCrLf & _
                    "CPF: " & cpf & vbCrLf & _
                    "Email: " & email & vbCrLf & _
                    "Cargo: " & cargo

        If MsgBox(mensagem, vbYesNo + vbQuestion, "Confirmar exclusão") = vbYes Then
            ws.Rows(linha).Delete
            MsgBox "Cadastro removido com sucesso.", vbInformation
            LimparCamposTerceirizado
        Else
            MsgBox "Exclusão cancelada.", vbInformation
        End If
    Else
        MsgBox "Senha incorreta. O terceirizado não foi removido.", vbCritical
    End If
End Sub

Private Sub botaoSalvarAlteracoes_Click()
    Dim SenhaDigitada As String
    Dim SenhaCorreta As String

    SenhaCorreta = "123" ' Substitua pela senha desejada para salvar

    SenhaDigitada = InputBox("Por favor, digite a senha para salvar as alterações do terceirizado:", "Autenticação Necessária")

    If SenhaDigitada = SenhaCorreta Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("TERCEIRIZADOS")

        If Trim(txtCPFTerceirizado.Text) = "" Then
            MsgBox "Informe o CPF do terceirizado para salvar as alterações.", vbExclamation
            Exit Sub
        End If

        Dim linha As Long
        linha = ProcurarLinhaPorCPF(FormatCPF(txtCPFTerceirizado.Text))

        If linha = 0 Then
            MsgBox "CPF não encontrado. Use 'Cadastrar' para novo registro.", vbExclamation
            Exit Sub
        End If

        ' Verifica se os dados são os mesmos antes de atualizar
        If Trim(txtNomeTerceirizado.Text) = ws.Cells(linha, 1).Value And _
           (IIf(optMasculinoTerce.Value, "MASCULINO", "FEMININO") = ws.Cells(linha, 2).Value Or (Not optMasculinoTerce.Value And Not optFemininoTerce.Value)) And _
           Trim(txtDataAdmissao.Text) = ws.Cells(linha, 3).Value And _
           Trim(txtUnidade.Text) = ws.Cells(linha, 4).Value And _
           FormatCPF(txtCPFTerceirizado.Text) = ws.Cells(linha, 5).Value And _
           Trim(txtCargo.Text) = ws.Cells(linha, 6).Value And _
           Trim(txtMatricula.Text) = ws.Cells(linha, 7).Value And _
           Trim(txtEmailTerceirizado.Text) = ws.Cells(linha, 8).Value And _
           Trim(txtSigla.Text) = ws.Cells(linha, 9).Value Then

            MsgBox "Os dados não foram alterados.", vbInformation
            Exit Sub ' Encerra a sub sem atualizar
        End If

        With ws
            ' Atualiza apenas se o campo de texto não estiver em branco
            If Trim(txtNomeTerceirizado.Text) <> "" Then
                .Cells(linha, 1).Value = ProperCase(txtNomeTerceirizado.Text)
            End If
            If optMasculinoTerce.Value Or optFemininoTerce.Value Then
                .Cells(linha, 2).Value = IIf(optMasculinoTerce.Value, "MASCULINO", "FEMININO")
            End If
            If Trim(txtDataAdmissao.Text) <> "" Then
                .Cells(linha, 3).Value = txtDataAdmissao.Text
            End If
            If Trim(txtUnidade.Text) <> "" Then
                .Cells(linha, 4).Value = txtUnidade.Text
            End If
            If Trim(txtCPFTerceirizado.Text) <> "" Then
                .Cells(linha, 5).Value = FormatCPF(txtCPFTerceirizado.Text)
            End If
            If Trim(txtCargo.Text) <> "" Then
                .Cells(linha, 6).Value = txtCargo.Text
            End If
            If Trim(txtMatricula.Text) <> "" Then
                .Cells(linha, 7).Value = txtMatricula.Text
            End If
            ' Validação do formato do e-mail ao salvar alterações
            If Trim(txtEmailTerceirizado.Text) <> "" Then
                If Not IsValidEmail(Trim(txtEmailTerceirizado.Text)) Then
                    MsgBox "Formato de e-mail inválido. Verifique se contém '@' e '.' e se não há caracteres inválidos.", vbExclamation
                    Exit Sub
                End If
                .Cells(linha, 8).Value = txtEmailTerceirizado.Text
            End If
            If Trim(txtSigla.Text) <> "" Then
                .Cells(linha, 9).Value = txtSigla.Text
            End If
        End With

        MsgBox "Alterações salvas com sucesso!", vbInformation
        LimparCamposTerceirizado
    Else
        MsgBox "Senha incorreta. As alterações não foram salvas.", vbCritical
    End If
End Sub

Private Sub btnVoltar_Click()
    Unload Me
    frmMenuPrincipal.Show
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

'===========================
'=== FORMATAR CAMPOS =======
'===========================

Private Sub txtCPFTerceirizado_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtCPFTerceirizado.Text = FormatCPF(txtCPFTerceirizado.Text)

     ' Se o CPF for encontrado, preenche os outros campos
    Dim linhaEncontradaPorCpf As Long
    linhaEncontradaPorCpf = ProcurarLinhaPorCPF(FormatCPF(txtCPFTerceirizado.Text))
    If linhaEncontradaPorCpf > 0 Then
        PreencherFormularioTerceirizado linhaEncontradaPorCpf
        MsgBox "CPF preenchido automaticamente!", vbInformation, "Aviso" ' Exibe a mensagem
    ElseIf Trim(txtCPFTerceirizado.Text) <> "" Then
        MsgBox "CPF não encontrado!", vbExclamation, "Aviso"
    End If
End Sub

Private Sub txtDataAdmissao_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If IsDate(txtDataAdmissao.Text) Then
        txtDataAdmissao.BackColor = vbWhite
        txtDataAdmissao.Text = FormatDataNascimento(txtDataAdmissao.Text)
    Else
        MsgBox "Data de admissão inválida. Por favor, insira uma data válida no formato dd/mm/aaaa.", vbExclamation, "Data inválida"
        txtDataAdmissao.BackColor = RGB(255, 200, 200) ' Destaque em vermelho claro
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
            mostrarMenu = True
        End If
    End If
End Sub

Private Sub UserForm_Terminate()
    If mostrarMenu Then
        frmMenuPrincipal.Show
    End If
End Sub

Private Sub UserForm_Click()
    ' (Reservado para ações ao clicar no fundo do formulário, se necessário)
End Sub

