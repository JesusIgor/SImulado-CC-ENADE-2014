VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA10 
   Caption         =   "          Questão alternativa 10 - Ciência da computação ENADE 2014"
   ClientHeight    =   10245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6375
   OleObjectBlob   =   "frm_QA10.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_fecharQA10_Click()
    'Descarrega o atual formulário da memória
    Unload Me
    'Caso o usuário tenha pressionado o botão "próximo"...
    If verifi = 1 Then
        'Será exibido a próxima questão
        frm_QA11.Show
    End If
    
    'Caso o usuário tenha pressionado o botão "finalizar"...
    If verifi = 2 Then
        'O formulário final será exibido
        frm_final.Show
    End If
End Sub



Private Sub cmd_finalizarQA10_Click()
    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA10
    'Verifica se o usuário pressionou o botão de finalizar
    verifi = 2
    
End Sub

Private Sub cmd_proxQA11_Click()

    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA10
    'Verifica se o usuário pressionou o botão de próximo
    verifi = 1
   
End Sub


Private Sub opt_altAQA10_Click()
    'A opção A armazenará a letra na variárial
    Q(10) = "A"
End Sub

Private Sub opt_altBQA10_Click()
    'A opção B armazenará a letra na variárial
    Q(10) = "B"
End Sub

Private Sub opt_altCQA10_Click()
    'A opção B armazenará a letra na variárial
    Q(10) = "C"
End Sub

Private Sub opt_altDQA10_Click()
    'A opção D armazenará a letra na variárial
    Q(10) = "D"
End Sub

Private Sub opt_altEQA10_Click()
    'A opção E armazenará a letra na variárial
    Q(10) = "E"
End Sub
Sub acumuloQA10()

i = 10
        'Mostra a resposta correta ao usuário
        resp_QA10.Visible = True
        
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
        opt_altAQA10.Enabled = False
        opt_altBQA10.Enabled = False
        opt_altCQA10.Enabled = False
        opt_altDQA10.Enabled = False
        opt_altEQA10.Enabled = False
        cmd_proxQA11.Enabled = False
        cmd_finalizarQA10.Enabled = False
        
        'A resposta será registrada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 17).Value = Q(i)
        
End Sub



