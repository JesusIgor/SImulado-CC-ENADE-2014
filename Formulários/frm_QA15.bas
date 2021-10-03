VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA15 
   Caption         =   "                        Questão alternativa 15 - Ciência da computação ENADE 2014"
   ClientHeight    =   8850.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7800
   OleObjectBlob   =   "frm_QA15.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_fecharQA15_Click()
    'Descarrega o atual formulário da memória
    Unload Me
    'Caso o usuário tenha pressionado o botão "próximo"...
    If verifi = 1 Then
        'Será exibido a próxima questão
        frm_QA16.Show
    End If
    
    'Caso o usuário tenha pressionado o botão "finalizar"...
    If verifi = 2 Then
        'O formulário final será exibido
        frm_final.Show
    End If
End Sub

Private Sub cmd_finalizarQA15_Click()
    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA15
    'Verifica se o usuário pressionou o botão de próximo
    verifi = 2
    
End Sub

Private Sub cmd_proxQA16_Click()
    
    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA15
    'Verifica se o usuário pressionou o botão de próximo
    verifi = 1
    
End Sub

Private Sub opt_altAQA15_Click()
    'A opção A armazenará a letra na variárial
    Q(15) = "A"
End Sub

Private Sub opt_altBQA15_Click()
    'A opção B armazenará a letra na variárial
    Q(15) = "B"
End Sub

Private Sub opt_altCQA15_Click()
    'A opção C armazenará a letra na variárial
    Q(15) = "C"
End Sub
Private Sub opt_altDQA15_Click()
    'A opção D armazenará a letra na variárial
    Q(15) = "D"
End Sub
Private Sub opt_altEQA15_Click()
    'A opção E armazenará a letra na variárial
    Q(15) = "E"
End Sub

Sub acumuloQA15()
i = 15
        'Mostra a resposta correta ao usuário
        resp_QA15.Visible = True
        
        'Inicio do encadeamento if
        If Q(i) = "C" Then
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
        opt_altAQA15.Enabled = False
        opt_altBQA15.Enabled = False
        opt_altCQA15.Enabled = False
        opt_altDQA15.Enabled = False
        opt_altEQA15.Enabled = False
        cmd_proxQA16.Enabled = False
        cmd_finalizarQA15.Enabled = False
        'Mostra dentro da planilha "respostas" qual foi a alternativa escolhida pelo usuario
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 22).Value = Q(i)

End Sub


