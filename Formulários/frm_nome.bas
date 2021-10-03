VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_nome 
   Caption         =   "                                                                                                 In�cio -  Ci�ncia da computa��o ENADE 2014"
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
    
    'O nome ser� registrado na guia de respostas
    ThisWorkbook.Worksheets("Respostas").Cells(linha, 2).Value = nome
    'O c�digo do usu�rio tamb�m ser� registrado
    ThisWorkbook.Worksheets("Respostas").Cells(linha, 1).Value = codigo
    
    'Todas as quest�es da prova come�am sem resposta
    For i = 1 To 35
        Q(i) = "NDA"
        'if para pular o registro nas colunas 13, 14 e 15, pois elas s�o dissertativas
        If coluna = 13 Then
            coluna = coluna + 3
        End If
        
        'Registro de cada quest�o vazia nas celulas
        ThisWorkbook.Worksheets("Respostas").Cells(linha, coluna).Value = Q(i)
        'implementa��o da coluna
        coluna = coluna + 1
    Next
    
    'fechar este formul�rio
    Unload Me
    'Abrir o pr�ximo
    frm_QD1.Show
    
End Sub

Private Sub txt_nome_Change()
    'Variavel nome assume o o texto da caixa
    nome = txt_nome.Text
    'Quando o usuario escrever na caixa de texto, o bot�o ser� habilitado
    cmd_ProxQD1.Enabled = True
    'O bot�o assumir� a cor azul, quando estiver habilitado
    cmd_ProxQD1.BackColor = &H8000000D
    
    'Caso a caixa de texto volte a ficar vazia...
    If txt_nome = "" Then
        'O bot�o ser� desabilitado
        cmd_ProxQD1.Enabled = False
        'A cor do bot�o ser� cinza
        cmd_ProxQD1.BackColor = &H8000000A
    End If
End Sub

Private Sub UserForm_Activate()
    'O bot�o iniciar� desabilitado...
    If cmd_ProxQD1.Enabled = False Then
        '...Ent�o a cor dele ser� cinza
        cmd_ProxQD1.BackColor = &H8000000A
    End If
    
    'Iniciando o valor da vari�vel linha
    linha = 5
    
    'Enquantoa c�lula n�o estiver vazia...
    Do While Not (ThisWorkbook.Worksheets("Respostas").Cells(linha, 2)) = ""
        'A vari�vel linha ser� incrementada em um valor
        linha = linha + 1
    Loop
    
    'C�lculo para registro do c�digo de usu�rio
    codigo = linha - 4
    
    'Zerando as var�aveis de ac�mulo
    acmAcertos = 0
    acmErros = 0
    acmBrancos = 0
    acmRespondidas = 0
    acmDissertBrancos = 0
    Dvazio = 0
    
    'Iniciando valor da coluna para registro das quest�es vazias
    coluna = 5

    
End Sub
