VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA32 
   Caption         =   "            Questão alternativa 32 - Ciência da computação ENADE 2014"
   ClientHeight    =   10500
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6360
   OleObjectBlob   =   "frm_QA32.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_fecharQA32_Click()
    'Descarrega o atual formulário da memória
    Unload Me
    'Caso o usuário tenha pressionado o botão "próximo"...
    If verifi = 1 Then
        'Será exibido a próxima questão
        frm_QA33.Show
    End If
    
    'Caso o usuário tenha pressionado o botão "finalizar"...
    If verifi = 2 Then
        'O formulário final será exibido
        frm_final.Show
    End If
End Sub



Private Sub cmd_finalizarQA32_Click()
    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA32
    'Verifica se o usuário pressionou o botão de finalizar
    verifi = 2
    
End Sub

Private Sub cmd_proxQA33_Click()

    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA32
    'Verifica se o usuário pressionou o botão de próximo
    verifi = 1
   
End Sub


Private Sub opt_altAQA32_Click()
    'A opção A armazenará a letra na variárial
    Q(32) = "A"
End Sub

Private Sub opt_altBQA32_Click()
    'A opção B armazenará a letra na variárial
    Q(32) = "B"
End Sub

Private Sub opt_altCQA32_Click()
    'A opção B armazenará a letra na variárial
    Q(32) = "C"
End Sub

Private Sub opt_altDQA32_Click()
    'A opção D armazenará a letra na variárial
    Q(32) = "D"
End Sub

Private Sub opt_altEQA32_Click()
    'A opção E armazenará a letra na variárial
    Q(32) = "E"
End Sub
Sub acumuloQA32()

i = 32
        'Mostra a resposta correta ao usuário
        resp_QA32.Visible = True
        
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
        opt_altAQA32.Enabled = False
        opt_altBQA32.Enabled = False
        opt_altCQA32.Enabled = False
        opt_altDQA32.Enabled = False
        opt_altEQA32.Enabled = False
        cmd_proxQA33.Enabled = False
        cmd_finalizarQA32.Enabled = False
        
        'A resposta será registrada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 39).Value = Q(i)
        
End Sub

Private Sub UserForm_Activate()

'Laço para alterações no formulário ativo
With Me
    'Definindo a altura interna da barra de rolagem
    .ScrollHeight = .InsideHeight * 1.41
End With

End Sub
