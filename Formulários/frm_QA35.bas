VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA35 
   Caption         =   "                                                                                                  Questão alternativa 35 - Ciência da computação ENADE 2014"
   ClientHeight    =   10545
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   13800
   OleObjectBlob   =   "frm_QA35.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA35"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_fecharQA35_Click()
    'Descarrega o atual formulário da memória
    Unload Me

     'O formulário final será exibido
     frm_final.Show
End Sub



Private Sub cmd_finalizarQA35_Click()
    'chamará a sub que registrará o que foi respondido no formulário
    Call acumuloQA35
    
End Sub


Private Sub opt_altAQA35_Click()
    'A opção A armazenará a letra na variárial
    Q(35) = "A"
End Sub

Private Sub opt_altBQA35_Click()
    'A opção B armazenará a letra na variárial
    Q(35) = "B"
End Sub

Private Sub opt_altCQA35_Click()
    'A opção B armazenará a letra na variárial
    Q(35) = "C"
End Sub

Private Sub opt_altDQA35_Click()
    'A opção D armazenará a letra na variárial
    Q(35) = "D"
End Sub

Private Sub opt_altEQA35_Click()
    'A opção E armazenará a letra na variárial
    Q(35) = "E"
End Sub
Sub acumuloQA35()

i = 35
        'Mostra a resposta correta ao usuário
        resp_QA35.Visible = True
        
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
        opt_altAQA35.Enabled = False
        opt_altBQA35.Enabled = False
        opt_altCQA35.Enabled = False
        opt_altDQA35.Enabled = False
        opt_altEQA35.Enabled = False
        cmd_finalizarQA35.Enabled = False
        
        'A resposta será registrada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 42).Value = Q(i)
        
End Sub

Private Sub UserForm_Activate()

'Laço para alterações no formulário ativo
With Me
    'Definindo a altura interna da barra de rolagem
    .ScrollHeight = .InsideHeight * 1.07
End With

End Sub

