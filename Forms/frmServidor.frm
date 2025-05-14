VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmServidor 
   Caption         =   "Cadastro De Servidor P�blico"
   ClientHeight    =   11295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765.001
   OleObjectBlob   =   "frmServidor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================
'=== VARI�VEIS E FUN��ES GERAIS
'===============================
Dim mostrarMenu As Boolean ' Vari�vel para controlar se o menu principal deve ser mostrado ao fechar
Dim linhaSelecionada As Long ' Vari�vel para armazenar a linha do usu�rio selecionado para edi��o

' Fun��o para formatar CPF
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
    If Len(digitsOnly) <> 11 And Trim(cpf) <> "" Then ' Adicionada condi��o para n�o mostrar mensagem se o CPF for vazio
        MsgBox "CPF inv�lido! Digite 11 n�meros.", vbCritical ' Mensagem de erro mais clara
        FormatCPF = cpf
        Exit Function
    End If
End Function

' Fun��o para colocar a primeira letra de cada palavra em mai�sculo
Function ProperCase(texto As String) As String
    ProperCase = Application.WorksheetFunction.Proper(texto)
End Function

' Fun��o para validar o formato do e-mail
Function IsValidEmail(email As String) As Boolean
    Dim AtPos As Integer
    Dim DotPos As Integer
    Dim i As Integer
    Dim invalidChars As String

    ' Verifica se cont�m "@" e "."
    AtPos = InStr(1, email, "@")
    DotPos = InStrRev(email, ".")

    If AtPos < 2 Or AtPos > Len(email) - 3 Or DotPos <= AtPos + 1 Or DotPos = Len(email) Then
        IsValidEmail = False
        Exit Function
    End If

    ' Verifica por caracteres inv�lidos (espa�os e alguns s�mbolos comuns)
    invalidChars = " ""(),:;<>\[]\"
    For i = 1 To Len(invalidChars)
        If InStr(1, email, Mid(invalidChars, i, 1)) > 0 Then
            IsValidEmail = False
            Exit Function
        End If
    Next i

    IsValidEmail = True
End Function

' Verifica se o CPF j� est� cadastrado na aba "SERVIDORES"
Function CPFJaExiste(cpf As String) As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SERVIDORES")

    Dim celula As Range
    For Each celula In ws.Range("E2:E" & ws.Cells(ws.Rows.Count, 5).End(xlUp).Row)
        If Replace(Trim(celula.Value), ".", "") = Replace(Trim(cpf), ".", "") Then
            CPFJaExiste = True
            Exit Function
        End If
    Next celula
    CPFJaExiste = False
End Function

' Procura a linha de um CPF na aba "SERVIDORES"
Function ProcurarLinhaPorCPF(cpf As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SERVIDORES")

    Dim linha As Long
    For linha = 2 To ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
        If Replace(Trim(ws.Cells(linha, 5).Value), ".", "") = Replace(Trim(cpf), ".", "") Then
            ProcurarLinhaPorCPF = linha
            Exit Function
        End If
    Next linha
    ProcurarLinhaPorCPF = 0
End Function

' Procura a linha de um Nome na aba "SERVIDORES"
Function ProcurarLinhaPorNome(nome As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SERVIDORES")
    Dim linha As Long
    Dim nomeSemAcentosEspacosSimbolos As String
    Dim nomeCelulaSemAcentosEspacosSimbolos As String
    
    ' Remove acentos, espa�os e s�mbolos do nome de busca
    nomeSemAcentosEspacosSimbolos = RemoverAcentosEspacosSimbolos(nome)
    
    For linha = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ' Remove acentos, espa�os e s�mbolos do nome da c�lula
        nomeCelulaSemAcentosEspacosSimbolos = RemoverAcentosEspacosSimbolos(ws.Cells(linha, 1).Value)
        
        If StrComp(Trim(nomeCelulaSemAcentosEspacosSimbolos), Trim(nomeSemAcentosEspacosSimbolos), vbTextCompare) = 0 Then
            ProcurarLinhaPorNome = linha
            Exit Function
        End If
    Next linha
    ProcurarLinhaPorNome = 0
End Function

' Fun��o para remover acentos, espa�os e s�mbolos de uma string
Function RemoverAcentosEspacosSimbolos(texto As String) As String
    Dim i As Integer
    Dim resultado As String
    Dim charCode As Integer
    
    For i = 1 To Len(texto)
        charCode = AscW(Mid(texto, i, 1))
        ' Filtra caracteres que est�o fora do intervalo b�sico do alfabeto latino
        If (charCode >= 65 And charCode <= 90) Or (charCode >= 97 And charCode <= 122) Then
            resultado = resultado & Mid(texto, i, 1)
        End If
    Next i
    RemoverAcentosEspacosSimbolos = resultado
End Function

' Limpa todos os campos do formul�rio
Sub LimparCamposServidor()
    txtNomeServidor.Text = ""
    txtEmailServidor.Text = ""
    txtMatricula.Text = ""
    txtCPFServidor.Text = ""
    txtSigla.Text = ""
    txtUnidade.Text = ""
    txtSituacaoFuncional.Text = ""
    txtCargo.Text = ""
    txtRegimeJuridico.Text = ""
    txtModalidade.Text = ""

    ' Limpa os option buttons de sexo
    optMasculinoServidor.Value = False
    optFemininoServidor.Value = False
    
    'Reseta a vari�vel de linha selecionada
    linhaSelecionada = 0
    
    'Muda o texto do bot�o salvar para Cadastrar
    'botaoSalvarAlteracoes.Caption = "Cadastrar" ' Comentei esta linha para n�o alterar o nome do bot�o
End Sub

' Preenche o formul�rio com os dados do servidor para edi��o
Sub PreencherFormularioServidor(linha As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SERVIDORES")
    
    linhaSelecionada = linha 'armazena a linha
    
    txtNomeServidor.Text = ws.Cells(linha, 1).Value
    
    If ws.Cells(linha, 2).Value = "MASCULINO" Then
        optMasculinoServidor.Value = True
        optFemininoServidor.Value = False
    Else
        optMasculinoServidor.Value = False
        optFemininoServidor.Value = True
    End If
    
    txtEmailServidor.Text = ws.Cells(linha, 3).Value
    txtMatricula.Text = ws.Cells(linha, 4).Value
    txtCPFServidor.Text = ws.Cells(linha, 5).Value
    txtSigla.Text = ws.Cells(linha, 6).Value
    txtUnidade.Text = ws.Cells(linha, 7).Value
    txtSituacaoFuncional.Text = ws.Cells(linha, 8).Value
    txtCargo.Text = ws.Cells(linha, 9).Value
    txtRegimeJuridico.Text = ws.Cells(linha, 10).Value
    txtModalidade.Text = ws.Cells(linha, 11).Value
    
    'Muda o texto do bot�o salvar para Atualizar
    'botaoSalvarAlteracoes.Caption = "Atualizar" ' Comentei esta linha para n�o alterar o nome do bot�o
End Sub

'================????
'=== BOT�ES DO FORM ====
'=======================

Private Sub botaoCadastrarServidor_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SERVIDORES")

    ' Valida��o dos campos obrigat�rios
    If Trim(txtNomeServidor.Text) = "" Or Trim(txtEmailServidor.Text) = "" Or _
       Trim(txtMatricula.Text) = "" Or Trim(txtCPFServidor.Text) = "" Or _
       Trim(txtSigla.Text) = "" Or Trim(txtUnidade.Text) = "" Or _
       Trim(txtSituacaoFuncional.Text) = "" Or Trim(txtCargo.Text) = "" Or _
       Trim(txtRegimeJuridico.Text) = "" Or Trim(txtModalidade.Text) = "" Or _
       (Not optMasculinoServidor.Value And Not optFemininoServidor.Value) Then
        MsgBox "Preencha todos os campos obrigat�rios, incluindo o sexo!", vbCritical
        Exit Sub
    End If

    ' Valida��o do formato do e-mail
    If Not IsValidEmail(Trim(txtEmailServidor.Text)) Then
        MsgBox "Formato de e-mail inv�lido. Verifique se cont�m '@' e '.' e se n�o h� caracteres inv�lidos.", vbExclamation
        Exit Sub
    End If

    ' Verifica duplicidade de CPF apenas se for um novo cadastro
    If linhaSelecionada = 0 And CPFJaExiste(FormatCPF(txtCPFServidor.Text)) Then
        MsgBox "CPF j� cadastrado!", vbExclamation
        Exit Sub
    End If

    ' Determina sexo
    Dim sexo As String
    If optMasculinoServidor.Value Then
        sexo = "MASCULINO"
    ElseIf optFemininoServidor.Value Then
        sexo = "FEMININO"
    Else
        MsgBox "Selecione o sexo do servidor!", vbExclamation
        Exit Sub
    End If

    ' Encontra pr�xima linha vazia
    Dim ultimaLinha As Long
    ultimaLinha = 4 ' Come�a a busca a partir da linha 4 (evita cabe�alho)
    Do While ws.Cells(ultimaLinha, 1).Value <> ""
        ultimaLinha = ultimaLinha + 1
    Loop

    ' Se estiver editando, usa a linha selecionada, sen�o, usa a �ltima linha vazia
    If linhaSelecionada > 0 Then
        ultimaLinha = linhaSelecionada
    End If
    
    ' Preenche os dados
    With ws
        .Cells(ultimaLinha, 1).Value = ProperCase(txtNomeServidor.Text)
        .Cells(ultimaLinha, 2).Value = sexo
        .Cells(ultimaLinha, 3).Value = txtEmailServidor.Text
        .Cells(ultimaLinha, 4).Value = txtMatricula.Text
        .Cells(ultimaLinha, 5).Value = FormatCPF(txtCPFServidor.Text)
        .Cells(ultimaLinha, 6).Value = txtSigla.Text
        .Cells(ultimaLinha, 7).Value = txtUnidade.Text
        .Cells(ultimaLinha, 8).Value = txtSituacaoFuncional.Text
        .Cells(ultimaLinha, 9).Value = txtCargo.Text
        .Cells(ultimaLinha, 10).Value = txtRegimeJuridico.Text
        .Cells(ultimaLinha, 11).Value = txtModalidade.Text
    End With

    If linhaSelecionada > 0 Then
        MsgBox "Servidor atualizado com sucesso!", vbInformation
    Else
        MsgBox "Servidor cadastrado com sucesso!", vbInformation
    End If
    
    LimparCamposServidor
End Sub

Private Sub botaoremover_Click()
    Dim SenhaDigitada As String
    Dim SenhaCorreta As String

    SenhaCorreta = "123" ' Substitua pela senha desejada para remover

    SenhaDigitada = InputBox("Por favor, digite a senha para remover o servidor:", "Autentica��o Necess�ria")

    If SenhaDigitada = SenhaCorreta Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("SERVIDORES")

        If Trim(txtCPFServidor.Text) = "" Then
            MsgBox "Por favor, preencha o campo CPF para remover o cadastro.", vbExclamation
            Exit Sub
        End If

        Dim linha As Long
        linha = ProcurarLinhaPorCPF(FormatCPF(txtCPFServidor.Text))

        If linha = 0 Then
            MsgBox "CPF n�o encontrado!", vbExclamation
            Exit Sub
        End If

        Dim nome As String, cpf As String, email As String, unidade As String, cargo As String
        nome = ws.Cells(linha, 1).Value
        cpf = ws.Cells(linha, 5).Value
        email = ws.Cells(linha, 3).Value
        unidade = ws.Cells(linha, 7).Value
        cargo = ws.Cells(linha, 9).Value

        Dim mensagem As String
        mensagem = "Deseja realmente remover este cadastro?" & vbCrLf & vbCrLf & _
                    "Nome: " & nome & vbCrLf & _
                    "CPF: " & cpf & vbCrLf & _
                    "Email: " & email & vbCrLf & _
                    "Unidade: " & unidade & vbCrLf & _
                    "Cargo: " & cargo

        If MsgBox(mensagem, vbYesNo + vbQuestion, "Confirmar exclus�o") = vbYes Then
            ws.Rows(linha).Delete
            MsgBox "Cadastro removido com sucesso.", vbInformation
            LimparCamposServidor
        Else
            MsgBox "Exclus�o cancelada.", vbInformation
        End If
    Else
        MsgBox "Senha incorreta. O servidor n�o foi removido.", vbCritical
    End If
End Sub

Private Sub botaoSalvarAlteracoes_Click()
    Dim SenhaDigitada As String
    Dim SenhaCorreta As String

    SenhaCorreta = "123" ' Substitua pela senha desejada para salvar

    SenhaDigitada = InputBox("Por favor, digite a senha para salvar as altera��es do servidor:", "Autentica��o Necess�ria")

    If SenhaDigitada = SenhaCorreta Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("SERVIDORES")

        If Trim(txtCPFServidor.Text) = "" Then
            MsgBox "Informe o CPF do servidor para salvar as altera��es.", vbExclamation
            Exit Sub
        End If

        Dim linha As Long
        linha = ProcurarLinhaPorCPF(FormatCPF(txtCPFServidor.Text))

        If linha = 0 Then
            MsgBox "CPF n�o encontrado. Use 'Cadastrar' para novo registro.", vbExclamation
            Exit Sub
        End If

        ' Verifica se os dados s�o os mesmos antes de atualizar
        If Trim(txtNomeServidor.Text) = ws.Cells(linha, 1).Value And _
           (IIf(optMasculinoServidor.Value, "MASCULINO", "FEMININO") = ws.Cells(linha, 2).Value Or (Not optMasculinoServidor.Value And Not optFemininoServidor.Value)) And _
           Trim(txtEmailServidor.Text) = ws.Cells(linha, 3).Value And _
           Trim(txtMatricula.Text) = ws.Cells(linha, 4).Value And _
           FormatCPF(txtCPFServidor.Text) = ws.Cells(linha, 5).Value And _
           Trim(txtSigla.Text) = ws.Cells(linha, 6).Value And _
           Trim(txtUnidade.Text) = ws.Cells(linha, 7).Value And _
           Trim(txtSituacaoFuncional.Text) = ws.Cells(linha, 8).Value And _
           Trim(txtCargo.Text) = ws.Cells(linha, 9).Value And _
           Trim(txtRegimeJuridico.Text) = ws.Cells(linha, 10).Value And _
           Trim(txtModalidade.Text) = ws.Cells(linha, 11).Value Then

            MsgBox "Os dados n�o foram alterados.", vbInformation
            Exit Sub ' Encerra a sub sem atualizar
        End If

        With ws
            ' Atualiza apenas se o campo de texto n�o estiver em branco
            If Trim(txtNomeServidor.Text) <> "" Then
                .Cells(linha, 1).Value = ProperCase(txtNomeServidor.Text)
            End If
            If optMasculinoServidor.Value Or optFemininoServidor.Value Then
                .Cells(linha, 2).Value = IIf(optMasculinoServidor.Value, "MASCULINO", "FEMININO")
            End If
            ' Valida��o do formato do e-mail ao salvar altera��es
            If Trim(txtEmailServidor.Text) <> "" Then
                If Not IsValidEmail(Trim(txtEmailServidor.Text)) Then
                    MsgBox "Formato de e-mail inv�lido. Verifique se cont�m '@' e '.' e se n�o h� caracteres inv�lidos.", vbExclamation
                    Exit Sub
                End If
                .Cells(linha, 3).Value = txtEmailServidor.Text
            End If
            If Trim(txtMatricula.Text) <> "" Then
                .Cells(linha, 4).Value = txtMatricula.Text
            End If
            If Trim(txtCPFServidor.Text) <> "" Then
                .Cells(linha, 5).Value = FormatCPF(txtCPFServidor.Text)
            End If
            If Trim(txtSigla.Text) <> "" Then
                .Cells(linha, 6).Value = txtSigla.Text
            End If
            If Trim(txtUnidade.Text) <> "" Then
                .Cells(linha, 7).Value = txtUnidade.Text
            End If
            If Trim(txtSituacaoFuncional.Text) <> "" Then
                .Cells(linha, 8).Value = txtSituacaoFuncional.Text
            End If
            If Trim(txtCargo.Text) <> "" Then
                .Cells(linha, 9).Value = txtCargo.Text
            End If
            If Trim(txtRegimeJuridico.Text) <> "" Then
                .Cells(linha, 10).Value = txtRegimeJuridico.Text
            End If
            If Trim(txtModalidade.Text) <> "" Then
                .Cells(linha, 11).Value = txtModalidade.Text
            End If
        End With

        MsgBox "Altera��es salvas com sucesso!", vbInformation
        LimparCamposServidor
    Else
        MsgBox "Senha incorreta. As altera��es n�o foram salvas.", vbCritical
    End If
End Sub

Private Sub btnVoltar_Click()
    Unload Me
    frmMenuPrincipal.Show
End Sub

'===========================
'=== FORMATAR CAMPOS =======
'===========================

Private Sub txtCPFServidor_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtCPFServidor.Text = FormatCPF(txtCPFServidor.Text)
    
    ' Se o CPF for encontrado, preenche os outros campos
    Dim linhaEncontradaPorCpf As Long
    linhaEncontradaPorCpf = ProcurarLinhaPorCPF(FormatCPF(txtCPFServidor.Text))
    If linhaEncontradaPorCpf > 0 Then
        PreencherFormularioServidor linhaEncontradaPorCpf
        MsgBox "CPF preenchido automaticamente!", vbInformation, "Aviso" ' Exibe a mensagem
    ElseIf Trim(txtCPFServidor.Text) <> "" Then
        MsgBox "CPF n�o encontrado!", vbExclamation, "Aviso"
    End If
End Sub

Private Sub txtModalidade_Change()
    ' (Pode adicionar alguma l�gica aqui se necess�rio ao alterar a modalidade)
End Sub

'==============================
'=== EVENTOS DE FECHAMENTO ====
'==============================

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Dim resposta As VbMsgBoxResult
        resposta = MsgBox("Tem certeza que deseja sair?", vbYesNo + vbQuestion, "Confirmar sa�da")

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
    ' (Reservado para a��es futuras ao clicar no fundo do formul�rio)
End Sub

