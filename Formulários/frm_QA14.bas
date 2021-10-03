VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA14 
   Caption         =   "                                 Questão alternativa 14 - Ciência da computação ENADE 2014"
   ClientHeight    =   8970.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8550.001
   OleObjectBlob   =   "frm_QA14.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_fecharQA14_Click()
    'Descarrega o atual formulário da memória
    Unload Me
    'Caso o usuário tenha pressionado o botão "próximo"...
    If verifi = 1 Then
        'Será exibido a próxima questão
        frm_QA15.Show
    End If
    
    'Caso o usuário tenha pressionado o botão "finalizar"...
    If verifi = 2 Then
        'O formulário final será exibido
        frm_final.Show
    End If
End Sub

Private Sub cmd_finalizarQA14_Click()
    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA14
    'Verifica se o usuário pressionou o botão de próximo
    verifi = 2
    
End Sub

Private Sub cmd_proxQA15_Click()
    
    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA14
    'Verifica se o usuário pressionou o botão de próximo
    verifi = 1
    
End Sub

Private Sub opt_altAQA14_Click()
    'A opção A armazenará a letra na variárial
    Q(14) = "A"
End Sub

Private Sub opt_altBQA14_Click()
    'A opção B armazenará a letra na variárial
    Q(14) = "B"
End Sub

Private Sub opt_altCQA14_Click()
    'A opção C armazenará a letra na variárial
    Q(14) = "C"
End Sub
Private Sub opt_altDQA14_Click()
    'A opção D armazenará a letra na variárial
    Q(14) = "D"
End Sub
Private Sub opt_altEQA14_Click()
    'A opção E armazenará a letra na variárial
    Q(14) = "E"
End Sub

Sub acumuloQA14()
i = 14
        'Mostra a resposta correta ao usuário
        resp_QA14.Visible = True
        
        'Inicio do encadeamento if
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
        opt_altAQA14.Enabled = False
        opt_altBQA14.Enabled = False
        opt_altCQA14.Enabled = False
        opt_altDQA14.Enabled = False
        opt_altEQA14.Enabled = False
        cmd_proxQA15.Enabled = False
        cmd_finalizarQA14.Enabled = False
        'Mostra dentro da planilha "respostas" qual foi a alternativa escolhida pelo usuario
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 21).Value = Q(i)

End Sub

