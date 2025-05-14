VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConselheiro 
   Caption         =   "Cadastro De Conselheiro"
   ClientHeight    =   11640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9720.001
   OleObjectBlob   =   "frmConselheiro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConselheiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Vari�vel para controlar se o menu principal deve ser mostrado ao fechar
Dim mostrarMenu As Boolean
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
Function CPFJaExiste(cpf As String, nomeAba As String) As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nomeAba) 'Permite usar a fun��o com diferentes abas

    Dim celula As Range
    ' Ajusta a coluna de busca para 'E'
    For Each celula In ws.Range("E2:E" & ws.Cells(ws.Rows.Count, 5).End(xlUp).Row)
        If Replace(Trim(celula.Value), ".", "") = Replace(Trim(cpf), ".", "") Then
            CPFJaExiste = True
            Exit Function
        End If
    Next celula
    CPFJaExiste = False
End Function

' Procura a linha de um CPF na aba "SERVIDORES"
Function ProcurarLinhaPorCPF(cpf As String, nomeAba As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nomeAba)  'Permite usar a fun��o com diferentes abas

    Dim linha As Long
    ' Ajusta a coluna de busca para 'E'
    For linha = 2 To ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
        If Replace(Trim(ws.Cells(linha, 5).Value), ".", "") = Replace(Trim(cpf), ".", "") Then
            ProcurarLinhaPorCPF = linha
            Exit Function
        End If
    Next linha
    ProcurarLinhaPorCPF = 0
End Function

' Procura a linha de um Nome na aba "SERVIDORES"
Function ProcurarLinhaPorNome(nome As String, nomeAba As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nomeAba) 'Permite usar a fun��o com diferentes abas
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
Sub LimparCamposConselheiro()
    txtNomeConselheiro.Text = ""
    txtEmailConselheiro.Text = ""
    txtUnidadeConselheiro.Text = ""
    txtRepresentacao.Text = ""
    txtCPFConselheiro.Text = ""
    txtMandato.Text = ""
    txtFormacao.Text = ""
    txtOcorrencias.Text = ""
    txtVinculo.Text = ""
    txtFim.Text = ""

    ' Desmarcar as op��es de Titular ou Suplente
    FrameTipoConselheiro.Controls("optTitular").Value = False
    FrameTipoConselheiro.Controls("optSuplente").Value = False

    ' Desmarcar as op��es de sexo
    FrameSexoConselheiro.Controls("optMasculinoConselheiro").Value = False
    FrameSexoConselheiro.Controls("optFemininoConselheiro").Value = False
    
    'Reseta a vari�vel de linha selecionada
    linhaSelecionada = 0
End Sub

' Preenche o formul�rio com os dados do Conselheiro para edi��o
Sub PreencherFormularioConselheiro(linha As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("CONSELHEIROS")
    
    linhaSelecionada = linha 'armazena a linha
    
    txtNomeConselheiro.Text = ws.Cells(linha, 1).Value
    
    If ws.Cells(linha, 2).Value = "MASCULINO" Then
        FrameSexoConselheiro.Controls("optMasculinoConselheiro").Value = True
        FrameSexoConselheiro.Controls("optFemininoConselheiro").Value = False
    Else
        FrameSexoConselheiro.Controls("optMasculinoConselheiro").Value = False
        FrameSexoConselheiro.Controls("optFemininoConselheiro").Value = True
    End If
    
    txtUnidadeConselheiro.Text = ws.Cells(linha, 3).Value
    txtRepresentacao.Text = ws.Cells(linha, 4).Value
    txtCPFConselheiro.Text = ws.Cells(linha, 5).Value
    txtEmailConselheiro.Text = ws.Cells(linha, 6).Value
    
     If ws.Cells(linha, 7).Value = "TITULAR" Then
        FrameTipoConselheiro.Controls("optTitular").Value = True
        FrameTipoConselheiro.Controls("optSuplente").Value = False
    Else
        FrameTipoConselheiro.Controls("optTitular").Value = False
        FrameTipoConselheiro.Controls("optSuplente").Value = True
    End If
    
    txtFim.Text = ws.Cells(linha, 8).Value
    txtMandato.Text = ws.Cells(linha, 9).Value
    txtFormacao.Text = ws.Cells(linha, 10).Value
    txtOcorrencias.Text = ws.Cells(linha, 11).Value
    txtVinculo.Text = ws.Cells(linha, 12).Value
    
    'Muda o texto do bot�o salvar para Atualizar
    'botaoSalvarAlteracoes.Caption = "Atualizar"
End Sub


'=======================
'=== BOT�ES DO FORM ====
'=======================

Private Sub botaoCadastrarConselheiro_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("CONSELHEIROS")
    ' Valida��o de campos obrigat�rios
    If Trim(txtNomeConselheiro.Text) = "" Or _
       Trim(txtEmailConselheiro.Text) = "" Or _
       Trim(txtCPFConselheiro.Text) = "" Or _
       Trim(txtUnidadeConselheiro.Text) = "" Or _
       Trim(txtRepresentacao.Text) = "" Or _
       Trim(txtMandato.Text) = "" Or _
       Trim(txtFormacao.Text) = "" Or _
       (Not FrameTipoConselheiro.Controls("optTitular").Value And Not FrameTipoConselheiro.Controls("optSuplente").Value) Or _
       (Not FrameSexoConselheiro.Controls("optMasculinoConselheiro").Value And Not FrameSexoConselheiro.Controls("optFemininoConselheiro").Value) Then
        MsgBox "Preencha todos os campos obrigat�rios, incluindo sexo e tipo de conselheiro!", vbCritical
        Exit Sub
    End If

    ' Valida��o do formato do e-mail
    If Not IsValidEmail(Trim(txtEmailConselheiro.Text)) Then
        MsgBox "Formato de e-mail inv�lido. Verifique se cont�m '@' e '.' e se n�o h� caracteres inv�lidos.", vbExclamation
        Exit Sub
    End If

    ' Verifica duplicidade de CPF
    If CPFJaExiste(FormatCPF(txtCPFConselheiro.Text), "CONSELHEIROS") Then
        MsgBox "CPF j� cadastrado!", vbExclamation
        Exit Sub
    End If

    ' Determina o sexo
    Dim sexo As String
    If FrameSexoConselheiro.Controls("optMasculinoConselheiro").Value Then
        sexo = "MASCULINO"
    ElseIf FrameSexoConselheiro.Controls("optFemininoConselheiro").Value Then
        sexo = "FEMININO"
    Else
        MsgBox "Selecione o sexo do conselheiro!", vbExclamation
        Exit Sub
    End If

    ' Cadastro de novo conselheiro
    Dim ultimaLinha As Long
    ultimaLinha = 4 ' Come�a a busca a partir da linha 4 (evita cabe�alho)

    Do While ws.Cells(ultimaLinha, 1).Value <> ""
        ultimaLinha = ultimaLinha + 1
    Loop
    
     ' Se estiver editando, usa a linha selecionada, sen�o, usa a �ltima linha vazia
    If linhaSelecionada > 0 Then
        ultimaLinha = linhaSelecionada
    End If

    ' Verifica se � titular ou suplente e atribui o valor correspondente
    Dim tipoConselheiro As String
    If FrameTipoConselheiro.Controls("optTitular").Value Then
        tipoConselheiro = "TITULAR"
    ElseIf FrameTipoConselheiro.Controls("optSuplente").Value Then
        tipoConselheiro = "SUPLENTE"
    Else
        MsgBox "Selecione o tipo de conselheiro!", vbCritical
        Exit Sub
    End If

    With ws
        .Cells(ultimaLinha, 1).Value = ProperCase(txtNomeConselheiro.Text) ' Nome
        .Cells(ultimaLinha, 2).Value = sexo ' Sexo
        .Cells(ultimaLinha, 3).Value = txtUnidadeConselheiro.Text ' Unidade
        .Cells(ultimaLinha, 4).Value = txtRepresentacao.Text ' Representa��o
        .Cells(ultimaLinha, 5).Value = FormatCPF(txtCPFConselheiro.Text) ' CPF
        .Cells(ultimaLinha, 6).Value = txtEmailConselheiro.Text ' Email
        .Cells(ultimaLinha, 7).Value = tipoConselheiro ' Titular ou Suplente
        .Cells(ultimaLinha, 8).Value = txtFim.Text ' Fim
        .Cells(ultimaLinha, 9).Value = txtMandato.Text ' Mandato
        .Cells(ultimaLinha, 10).Value = txtFormacao.Text ' Forma��o
        .Cells(ultimaLinha, 11).Value = txtOcorrencias.Text ' Ocorr�ncias
        .Cells(ultimaLinha, 12).Value = txtVinculo.Text ' V�nculo
    End With

    MsgBox "Conselheiro cadastrado com sucesso!", vbInformation
    LimparCamposConselheiro
End Sub

Private Sub botaoremover_Click()
    Dim SenhaDigitada As String
    Dim SenhaCorreta As String

    SenhaCorreta = "123" ' Voc� pode usar uma senha diferente para remover

    SenhaDigitada = InputBox("Por favor, digite a senha para remover o usu�rio:", "Autentica��o Necess�ria")

    If SenhaDigitada = SenhaCorreta Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("CONSELHEIROS")

        If Trim(txtCPFConselheiro.Text) = "" Then
            MsgBox "Por favor, preencha o campo CPF para remover o cadastro.", vbExclamation
            Exit Sub
        End If

        Dim linha As Long
        linha = ProcurarLinhaPorCPF(FormatCPF(txtCPFConselheiro.Text), "CONSELHEIROS")

        If linha = 0 Then
            MsgBox "CPF n�o encontrado!", vbExclamation
            Exit Sub
        End If

        Dim nome As String, cpf As String, email As String, tipo As String
        nome = ws.Cells(linha, 1).Value
        cpf = ws.Cells(linha, 5).Value
        email = ws.Cells(linha, 6).Value
        tipo = ws.Cells(linha, 7).Value

        Dim mensagem As String
        mensagem = "Deseja realmente remover este cadastro?" & vbCrLf & vbCrLf & _
                    "Nome: " & nome & vbCrLf & _
                    "CPF: " & cpf & vbCrLf & _
                    "Email: " & email & vbCrLf & _
                    "Titular ou Suplente: " & tipo

        If MsgBox(mensagem, vbYesNo + vbQuestion, "Confirmar exclus�o") = vbYes Then
            ws.Rows(linha).Delete
            MsgBox "Cadastro removido com sucesso.", vbInformation
            LimparCamposConselheiro
        Else
            MsgBox "Exclus�o cancelada.", vbInformation
        End If
    Else
        MsgBox "Senha incorreta. O usu�rio n�o foi removido.", vbCritical
    End If
End Sub

Private Sub botaoSalvarAlteracoes_Click()
    Dim SenhaDigitada As String
    Dim SenhaCorreta As String

    SenhaCorreta = "123" ' Substitua pela senha que voc� deseja usar

    SenhaDigitada = InputBox("Por favor, digite a senha para salvar as altera��es:", "Autentica��o Necess�ria")

    If SenhaDigitada = SenhaCorreta Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("CONSELHEIROS")

        If Trim(txtCPFConselheiro.Text) = "" Then
            MsgBox "Informe o CPF para salvar as altera��es.", vbExclamation
            Exit Sub
        End If

        Dim linha As Long
        linha = ProcurarLinhaPorCPF(FormatCPF(txtCPFConselheiro.Text), "CONSELHEIROS")

        If linha = 0 Then
            MsgBox "CPF n�o encontrado. Use 'Cadastrar' para novo registro.", vbExclamation
            Exit Sub
        End If

         ' Verifica se os dados s�o os mesmos antes de atualizar
        If Trim(txtNomeConselheiro.Text) = ws.Cells(linha, 1).Value And _
           (IIf(FrameSexoConselheiro.Controls("optMasculinoConselheiro").Value, "MASCULINO", "FEMININO") = ws.Cells(linha, 2).Value Or (Not FrameSexoConselheiro.Controls("optMasculinoConselheiro").Value And Not FrameSexoConselheiro.Controls("optFemininoConselheiro").Value)) And _
           Trim(txtUnidadeConselheiro.Text) = ws.Cells(linha, 3).Value And _
           Trim(txtRepresentacao.Text) = ws.Cells(linha, 4).Value And _
           FormatCPF(txtCPFConselheiro.Text) = ws.Cells(linha, 5).Value And _
           Trim(txtEmailConselheiro.Text) = ws.Cells(linha, 6).Value And _
           (IIf(FrameTipoConselheiro.Controls("optTitular").Value, "TITULAR", "SUPLENTE") = ws.Cells(linha, 7).Value Or (Not FrameTipoConselheiro.Controls("optTitular").Value And Not FrameTipoConselheiro.Controls("optSuplente").Value)) And _
           Trim(txtFim.Text) = ws.Cells(linha, 8).Value And _
           Trim(txtMandato.Text) = ws.Cells(linha, 9).Value And _
           Trim(txtFormacao.Text) = ws.Cells(linha, 10).Value And _
           Trim(txtOcorrencias.Text) = ws.Cells(linha, 11).Value And _
           Trim(txtVinculo.Text) = ws.Cells(linha, 12).Value Then

            MsgBox "Os dados n�o foram alterados.", vbInformation
            Exit Sub ' Encerra a sub sem atualizar
        End If

        With ws
            ' Atualiza apenas se o campo de texto n�o estiver em branco
            If Trim(txtNomeConselheiro.Text) <> "" Then
                .Cells(linha, 1).Value = ProperCase(txtNomeConselheiro.Text)
            End If

            ' Para o sexo, atualiza se algum bot�o de op��o estiver selecionado
            If FrameSexoConselheiro.Controls("optMasculinoConselheiro").Value Or FrameSexoConselheiro.Controls("optFemininoConselheiro").Value Then
                .Cells(linha, 2).Value = IIf(FrameSexoConselheiro.Controls("optMasculinoConselheiro").Value, "MASCULINO", "FEMININO")
            End If

            If Trim(txtUnidadeConselheiro.Text) <> "" Then
                .Cells(linha, 3).Value = txtUnidadeConselheiro.Text
            End If

            If Trim(txtRepresentacao.Text) <> "" Then
                .Cells(linha, 4).Value = txtRepresentacao.Text
            End If

            If Trim(txtCPFConselheiro.Text) <> "" Then
                .Cells(linha, 5).Value = FormatCPF(txtCPFConselheiro.Text)
            End If

            ' Valida��o do formato do e-mail ao salvar altera��es
            If Trim(txtEmailConselheiro.Text) <> "" Then
                If Not IsValidEmail(Trim(txtEmailConselheiro.Text)) Then
                    MsgBox "Formato de e-mail inv�lido. Verifique se cont�m '@' e '.' e se n�o h� caracteres inv�lidos.", vbExclamation
                    Exit Sub
                End If
                .Cells(linha, 6).Value = txtEmailConselheiro.Text
            End If

            ' Para o tipo, atualiza se algum bot�o de op��o estiver selecionado
            If FrameTipoConselheiro.Controls("optTitular").Value Or FrameTipoConselheiro.Controls("optSuplente").Value Then
                .Cells(linha, 7).Value = IIf(FrameTipoConselheiro.Controls("optTitular").Value, "TITULAR", "SUPLENTE")
            End If

            If Trim(txtFim.Text) <> "" Then
                .Cells(linha, 8).Value = txtFim.Text
            End If

            If Trim(txtMandato.Text) <> "" Then
                .Cells(linha, 9).Value = txtMandato.Text
            End If

            If Trim(txtFormacao.Text) <> "" Then
                .Cells(linha, 10).Value = txtFormacao.Text
            End If

            If Trim(txtOcorrencias.Text) <> "" Then
                .Cells(linha, 11).Value = txtOcorrencias.Text
            End If

            If Trim(txtVinculo.Text) <> "" Then
                .Cells(linha, 12).Value = txtVinculo.Text
            End If
        End With

        MsgBox "Altera��es salvas com sucesso!", vbInformation
        LimparCamposConselheiro
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

Private Sub txtCPFConselheiro_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtCPFConselheiro.Text = FormatCPF(txtCPFConselheiro.Text)
    
     ' Se o CPF for encontrado, preenche os outros campos
    Dim linhaEncontradaPorCpf As Long
    linhaEncontradaPorCpf = ProcurarLinhaPorCPF(FormatCPF(txtCPFConselheiro.Text), "CONSELHEIROS")
    If linhaEncontradaPorCpf > 0 Then
        PreencherFormularioConselheiro linhaEncontradaPorCpf
        MsgBox "CPF preenchido automaticamente!", vbInformation, "Aviso" ' Exibe a mensagem
    ElseIf Trim(txtCPFConselheiro.Text) <> "" Then
        MsgBox "CPF n�o encontrado!", vbExclamation, "Aviso"
    End If
End Sub

Private Sub OptionButton1_Click()
    ' N�o utilizado
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
    ' (Reservado para a��es ao clicar no fundo do formul�rio, se necess�rio)
End Sub



