VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA17 
   Caption         =   "                                                                              Questão alternativa 17 - Ciência da computação ENADE 2014"
   ClientHeight    =   6885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12705
   OleObjectBlob   =   "frm_QA17.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_fecharQA17_Click()
    'Descarrega o atual formulário da memória
    Unload Me
    'Caso o usuário tenha pressionado o botão "próximo"...
    If verifi = 1 Then
        'Será exibido a próxima questão
        frm_QA18.Show
    End If
    
    'Caso o usuário tenha pressionado o botão "finalizar"...
    If verifi = 2 Then
        'O formulário final será exibido
        frm_final.Show
    End If
End Sub

Private Sub cmd_finalizarQA17_Click()
    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA17
    'Verifica se o usuário pressionou o botão de próximo
    verifi = 2
    
End Sub

Private Sub cmd_proxQA18_Click()
    
    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA17
    'Verifica se o usuário pressionou o botão de próximo
    verifi = 1
    
End Sub

Private Sub opt_altAQA17_Click()
    'A opção A armazenará a letra na variárial
    Q(17) = "A"
End Sub

Private Sub opt_altBQA17_Click()
    'A opção B armazenará a letra na variárial
    Q(17) = "B"
End Sub

Private Sub opt_altCQA17_Click()
    'A opção C armazenará a letra na variárial
    Q(17) = "C"
End Sub
Private Sub opt_altDQA17_Click()
    'A opção D armazenará a letra na variárial
    Q(17) = "D"
End Sub
Private Sub opt_altEQA17_Click()
    'A opção E armazenará a letra na variárial
    Q(17) = "E"
End Sub

Sub acumuloQA17()
i = 17
        'Mostra a resposta correta ao usuário
        resp_QA17.Visible = True
        
        'Caso a resposta seja "A"...
        If Q(i) = "A" Then
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
        opt_altAQA17.Enabled = False
        opt_altBQA17.Enabled = False
        opt_altCQA17.Enabled = False
        opt_altDQA17.Enabled = False
        opt_altEQA17.Enabled = False
        cmd_proxQA18.Enabled = False
        cmd_finalizarQA17.Enabled = False
        'Mostra dentro da planilha "respostas" qual foi a alternativa escolhida pelo usuario
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 24).Value = Q(i)

End Sub
