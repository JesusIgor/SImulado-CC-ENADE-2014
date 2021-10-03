VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA8 
   Caption         =   "                                  Questão alternativa 8 - Ciência da computação ENADE 2014"
   ClientHeight    =   10425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8250.001
   OleObjectBlob   =   "frm_QA8.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_fecharQA8_Click()
    'Descarrega o atual formulário da memória
    Unload Me
    'Caso o usuário tenha pressionado o botão "próximo"...
    If verifi = 1 Then
        'Será exibido a próxima questão
        frm_QD3.Show
    End If
    
    'Caso o usuário tenha pressionado o botão "finalizar"...
    If verifi = 2 Then
        'O formulário final será exibido
        frm_final.Show
    End If
End Sub



Private Sub cmd_finalizarQA8_Click()
    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA8
    'Verifica se o usuário pressionou o botão de finalizar
    verifi = 2
    
End Sub

Private Sub cmd_proxQD3_Click()

    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA8
    'Verifica se o usuário pressionou o botão de próximo
    verifi = 1
   
End Sub


Private Sub opt_altAQA8_Click()
    'A opção A armazenará a letra na variárial
    Q(8) = "A"
End Sub

Private Sub opt_altBQA8_Click()
    'A opção B armazenará a letra na variárial
    Q(8) = "B"
End Sub

Private Sub opt_altCQA8_Click()
    'A opção B armazenará a letra na variárial
    Q(8) = "C"
End Sub

Private Sub opt_altDQA8_Click()
    'A opção D armazenará a letra na variárial
    Q(8) = "D"
End Sub

Private Sub opt_altEQA8_Click()
    'A opção E armazenará a letra na variárial
    Q(8) = "E"
End Sub
Sub acumuloQA8()

i = 8
        'Mostra a resposta correta ao usuário
        resp_QA8.Visible = True
        
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
        
        'Inativando todos os botões após o usuário registrar a resposta
        opt_altAQA8.Enabled = False
        opt_altBQA8.Enabled = False
        opt_altCQA8.Enabled = False
        opt_altDQA8.Enabled = False
        opt_altEQA8.Enabled = False
        cmd_proxQD3.Enabled = False
        cmd_finalizarQA8.Enabled = False
        
        'A resposta será registrada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 12).Value = Q(i)
        
End Sub


Private Sub UserForm_Activate()

'Laço para alterações no formulário ativo
With Me
    'Definindo a altura interna da barra de rolagem
    .ScrollHeight = .InsideHeight * 1.75
End With

End Sub
