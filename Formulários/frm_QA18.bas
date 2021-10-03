VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA18 
   Caption         =   "                                                                                           Questão alternativa 18 - Ciência da computação ENADE 2014"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13830
   OleObjectBlob   =   "frm_QA18.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_fecharQA18_Click()
    'Descarrega o atual formulário da memória
    Unload Me
    'Caso o usuário tenha pressionado o botão "próximo"...
    If verifi = 1 Then
        'Será exibido a próxima questão
        frm_QA19.Show
    End If
    
    'Caso o usuário tenha pressionado o botão "finalizar"...
    If verifi = 2 Then
        'O formulário final será exibido
        frm_final.Show
    End If
End Sub

Private Sub cmd_finalizarQA18_Click()
    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA18
    'Verifica se o usuário pressionou o botão de próximo
    verifi = 2
    
End Sub

Private Sub cmd_proxQA19_Click()
    
    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA18
    'Verifica se o usuário pressionou o botão de próximo
    verifi = 1
    
End Sub

Private Sub opt_altAQA18_Click()
    'A opção A armazenará a letra na variárial
    Q(18) = "A"
End Sub

Private Sub opt_altBQA18_Click()
    'A opção B armazenará a letra na variárial
    Q(18) = "B"
End Sub

Private Sub opt_altCQA18_Click()
    'A opção C armazenará a letra na variárial
    Q(18) = "C"
End Sub
Private Sub opt_altDQA18_Click()
    'A opção D armazenará a letra na variárial
    Q(18) = "D"
End Sub
Private Sub opt_altEQA18_Click()
    'A opção E armazenará a letra na variárial
    Q(18) = "E"
End Sub

Sub acumuloQA18()
i = 18
        'Mostra a resposta correta ao usuário
        resp_QA18.Visible = True
        
        'Caso a resposta seja "E"...
        If Q(i) = "E" Then
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
        opt_altAQA18.Enabled = False
        opt_altBQA18.Enabled = False
        opt_altCQA18.Enabled = False
        opt_altDQA18.Enabled = False
        opt_altEQA18.Enabled = False
        cmd_proxQA19.Enabled = False
        cmd_finalizarQA18.Enabled = False
        'Mostra dentro da planilha "respostas" qual foi a alternativa escolhida pelo usuario
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 25).Value = Q(i)

End Sub

