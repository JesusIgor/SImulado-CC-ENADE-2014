VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA16 
   Caption         =   "                              Questão alternativa 16 - Ciência da computação ENADE 2014"
   ClientHeight    =   9885.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8130
   OleObjectBlob   =   "frm_QA16.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_fecharQA16_Click()
    'Descarrega o atual formulário da memória
    Unload Me
    'Caso o usuário tenha pressionado o botão "próximo"...
    If verifi = 1 Then
        'Será exibido a próxima questão
        frm_QA17.Show
    End If
    
    'Caso o usuário tenha pressionado o botão "finalizar"...
    If verifi = 2 Then
        'O formulário final será exibido
        frm_final.Show
    End If
End Sub



Private Sub cmd_finalizarQA16_Click()
    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA16
    'Verifica se o usuário pressionou o botão de finalizar
    verifi = 2
    
End Sub

Private Sub cmd_proxQA17_Click()

    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA16
    'Verifica se o usuário pressionou o botão de próximo
    verifi = 1
   
End Sub


Private Sub opt_altAQA16_Click()
    'A opção A armazenará a letra na variárial
    Q(16) = "A"
End Sub

Private Sub opt_altBQA16_Click()
    'A opção B armazenará a letra na variárial
    Q(16) = "B"
End Sub

Private Sub opt_altCQA16_Click()
    'A opção B armazenará a letra na variárial
    Q(16) = "C"
End Sub

Private Sub opt_altDQA16_Click()
    'A opção D armazenará a letra na variárial
    Q(16) = "D"
End Sub

Private Sub opt_altEQA16_Click()
    'A opção E armazenará a letra na variárial
    Q(16) = "E"
End Sub
Sub acumuloQA16()

i = 16
        'Mostra a resposta correta ao usuário
        resp_QA16.Visible = True
        
        'Caso a resposta seja "C"...
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
        opt_altAQA16.Enabled = False
        opt_altBQA16.Enabled = False
        opt_altCQA16.Enabled = False
        opt_altDQA16.Enabled = False
        opt_altEQA16.Enabled = False
        cmd_proxQA17.Enabled = False
        cmd_finalizarQA16.Enabled = False
        
        'A resposta será registrada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 23).Value = Q(i)
        
End Sub

