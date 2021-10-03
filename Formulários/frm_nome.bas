VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_nome 
   Caption         =   "                                                                                                 Início -  Ciência da computação ENADE 2014"
   ClientHeight    =   5790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12825
   OleObjectBlob   =   "frm_nome.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_nome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ProxQD1_Click()
    
    'O nome será registrado na guia de respostas
    ThisWorkbook.Worksheets("Respostas").Cells(linha, 2).Value = nome
    'O código do usuário também será registrado
    ThisWorkbook.Worksheets("Respostas").Cells(linha, 1).Value = codigo
    
    'Todas as questões da prova começam sem resposta
    For i = 1 To 35
        Q(i) = "NDA"
        'if para pular o registro nas colunas 13, 14 e 15, pois elas são dissertativas
        If coluna = 13 Then
            coluna = coluna + 3
        End If
        
        'Registro de cada questão vazia nas celulas
        ThisWorkbook.Worksheets("Respostas").Cells(linha, coluna).Value = Q(i)
        'implementação da coluna
        coluna = coluna + 1
    Next
    
    'fechar este formulário
    Unload Me
    'Abrir o próximo
    frm_QD1.Show
    
End Sub

Private Sub txt_nome_Change()
    'Variavel nome assume o o texto da caixa
    nome = txt_nome.Text
    'Quando o usuario escrever na caixa de texto, o botão será habilitado
    cmd_ProxQD1.Enabled = True
    'O botão assumirá a cor azul, quando estiver habilitado
    cmd_ProxQD1.BackColor = &H8000000D
    
    'Caso a caixa de texto volte a ficar vazia...
    If txt_nome = "" Then
        'O botão será desabilitado
        cmd_ProxQD1.Enabled = False
        'A cor do botão será cinza
        cmd_ProxQD1.BackColor = &H8000000A
    End If
End Sub

Private Sub UserForm_Activate()
    'O botão iniciará desabilitado...
    If cmd_ProxQD1.Enabled = False Then
        '...Então a cor dele será cinza
        cmd_ProxQD1.BackColor = &H8000000A
    End If
    
    'Iniciando o valor da variável linha
    linha = 5
    
    'Enquantoa célula não estiver vazia...
    Do While Not (ThisWorkbook.Worksheets("Respostas").Cells(linha, 2)) = ""
        'A variável linha será incrementada em um valor
        linha = linha + 1
    Loop
    
    'Cálculo para registro do código de usuário
    codigo = linha - 4
    
    'Zerando as varíaveis de acúmulo
    acmAcertos = 0
    acmErros = 0
    acmBrancos = 0
    acmRespondidas = 0
    acmDissertBrancos = 0
    Dvazio = 0
    
    'Iniciando valor da coluna para registro das questões vazias
    coluna = 5

    
End Sub
