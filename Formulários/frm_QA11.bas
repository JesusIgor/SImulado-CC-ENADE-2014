VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA11 
   Caption         =   "               Questão alternativa 11 - Ciência da computação ENADE 2014"
   ClientHeight    =   10515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6555
   OleObjectBlob   =   "frm_QA11.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_fecharQA11_Click()
    'Descarrega o atual formulário da memória
    Unload Me
    'Caso o usuário tenha pressionado o botão "próximo"...
    If verifi = 1 Then
        'Será exibido a próxima questão
        frm_QA12.Show
    End If
    
    'Caso o usuário tenha pressionado o botão "finalizar"...
    If verifi = 2 Then
        'O formulário final será exibido
        frm_final.Show
    End If
End Sub

Private Sub cmd_finalizarQA11_Click()
    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA11
    'Verifica se o usuário pressionou o botão de próximo
    verifi = 2
    
End Sub

Private Sub cmd_proxQA12_Click()
    
    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA11
    'Verifica se o usuário pressionou o botão de próximo
    verifi = 1
    
End Sub

Private Sub opt_altAQA11_Click()
    'A opção A armazenará a letra na variárial
    Q(11) = "A"
End Sub

Private Sub opt_altBQA11_Click()
    'A opção B armazenará a letra na variárial
    Q(11) = "B"
End Sub

Private Sub opt_altCQA11_Click()
    'A opção C armazenará a letra na variárial
    Q(11) = "C"
End Sub
Private Sub opt_altDQA11_Click()
    'A opção D armazenará a letra na variárial
    Q(11) = "D"
End Sub
Private Sub opt_altEQA11_Click()
    'A opção E armazenará a letra na variárial
    Q(11) = "E"
End Sub

Sub acumuloQA11()
i = 11
        'Mostra a resposta correta ao usuário
        resp_QA11.Visible = True
        
        'Caso a resposta seja "B"...
        If Q(i) = "B" Then
            'O acúmulo de acertos será incrementado em um valor
            acmAcertos = acmAcertos + 1
            'Informa o acerto ao usuário
            lbl_acerto.Visible = True
        'Caso contrário...
        Else
            'Caso não haja resposta
            If Q(i) = "NDA" Then
                'A questão se manterá "NDA"
                Q(i) = "NDA"
            'Caso haja, mas não correta...
            Else
                'O acúmulo de erros será incrementado em um valor
                acmErros = acmErros + 1
            End If
            'Informa o erro ao usuário
            lbl_erro.Visible = True
        End If
        
        'Inativando todos os botões após o usuário registrar a resposta
        opt_altAQA11.Enabled = False
        opt_altBQA11.Enabled = False
        opt_altCQA11.Enabled = False
        opt_altDQA11.Enabled = False
        opt_altEQA11.Enabled = False
        cmd_proxQA12.Enabled = False
        cmd_finalizarQA11.Enabled = False
        'Mostra dentro da planilha "respostas" qual foi a alternativa escolhida pelo usuario
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 18).Value = Q(i)

End Sub

Private Sub UserForm_Activate()
'Laço para alterações no formulário ativo
With Me
    'Definindo a altura interna da barra de rolagem
    .ScrollHeight = .InsideHeight * 1.13
End With
End Sub
