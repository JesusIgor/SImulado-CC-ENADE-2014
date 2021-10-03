VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA4 
   Caption         =   "                Questão alternativa 4 - Ciência da computação ENADE 2014"
   ClientHeight    =   10530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6555
   OleObjectBlob   =   "frm_QA4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_fecharQA4_Click()
    'Descarrega o atual formulário da memória
    Unload Me
    'Caso o usuário tenha pressionado o botão "próximo"...
    If verifi = 1 Then
        'Será exibido a próxima questão
        frm_QA5.Show
    End If
    
    'Caso o usuário tenha pressionado o botão "finalizar"...
    If verifi = 2 Then
        'O formulário final será exibido
        frm_final.Show
    End If
End Sub



Private Sub cmd_finalizarQA4_Click()
    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA4
    'Verifica se o usuário pressionou o botão de finalizar
    verifi = 2
    
End Sub

Private Sub cmd_proxQA5_Click()

    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA4
    'Verifica se o usuário pressionou o botão de próximo
    verifi = 1
   
End Sub


Private Sub opt_altAQA4_Click()
    'A opção A armazenará a letra na variárial
    Q(4) = "A"
End Sub

Private Sub opt_altBQA4_Click()
    'A opção B armazenará a letra na variárial
    Q(4) = "B"
End Sub

Private Sub opt_altCQA4_Click()
    'A opção B armazenará a letra na variárial
    Q(4) = "C"
End Sub

Private Sub opt_altDQA4_Click()
    'A opção D armazenará a letra na variárial
    Q(4) = "D"
End Sub

Private Sub opt_altEQA4_Click()
    'A opção E armazenará a letra na variárial
    Q(4) = "E"
End Sub
Sub acumuloQA4()

i = 4
        'Mostra a resposta correta ao usuário
        resp_QA4.Visible = True
        
        'Caso a resposta seja "A"...
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
        opt_altAQA4.Enabled = False
        opt_altBQA4.Enabled = False
        opt_altCQA4.Enabled = False
        opt_altDQA4.Enabled = False
        opt_altEQA4.Enabled = False
        cmd_proxQA5.Enabled = False
        cmd_finalizarQA4.Enabled = False
        
        'A resposta será registrada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 8).Value = Q(i)
        
End Sub

Private Sub UserForm_Activate()

'Laço para alterações no formulário ativo
With Me
    'Definindo a altura interna da barra de rolagem
    .ScrollHeight = .InsideHeight * 1.409
End With

End Sub


