VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA5 
   Caption         =   "                         Questão alternativa 5 - Ciência da computação ENADE 2014"
   ClientHeight    =   10620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7500
   OleObjectBlob   =   "frm_QA5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_fecharQA5_Click()
    'Descarrega o atual formulário da memória
    Unload Me
    'Caso o usuário tenha pressionado o botão "próximo"...
    If verifi = 1 Then
        'Será exibido a próxima questão
        frm_QA6.Show
    End If
    
    'Caso o usuário tenha pressionado o botão "finalizar"...
    If verifi = 2 Then
        'O formulário final será exibido
        frm_final.Show
    End If
End Sub



Private Sub cmd_finalizarQA5_Click()
    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA5
    'Verifica se o usuário pressionou o botão de finalizar
    verifi = 2
    
End Sub

Private Sub cmd_proxQA6_Click()

    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA5
    'Verifica se o usuário pressionou o botão de próximo
    verifi = 1
   
End Sub


Private Sub opt_altAQA5_Click()
    'A opção A armazenará a letra na variárial
    Q(5) = "A"
End Sub

Private Sub opt_altBQA5_Click()
    'A opção B armazenará a letra na variárial
    Q(5) = "B"
End Sub

Private Sub opt_altCQA5_Click()
    'A opção B armazenará a letra na variárial
    Q(5) = "C"
End Sub

Private Sub opt_altDQA5_Click()
    'A opção D armazenará a letra na variárial
    Q(5) = "D"
End Sub

Private Sub opt_altEQA5_Click()
    'A opção E armazenará a letra na variárial
    Q(5) = "E"
End Sub
Sub acumuloQA5()

i = 5
        'Mostra a resposta correta ao usuário
        resp_QA5.Visible = True
        
        'Caso a resposta seja "D"...
        If Q(i) = "D" Then
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
        
        'Desativando todos os botões após o usuário registrar a resposta
        opt_altAQA5.Enabled = False
        opt_altBQA5.Enabled = False
        opt_altCQA5.Enabled = False
        opt_altDQA5.Enabled = False
        opt_altEQA5.Enabled = False
        cmd_proxQA6.Enabled = False
        cmd_finalizarQA5.Enabled = False
        
        'A resposta será registrada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 9).Value = Q(i)
        
End Sub

Private Sub UserForm_Activate()

'Laço para alterações no formulário ativo
With Me
    'Definindo a altura interna da barra de rolagem
    .ScrollHeight = .InsideHeight * 1.7
End With

End Sub
