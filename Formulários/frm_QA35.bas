VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA35 
   Caption         =   "                                                                                                  Quest�o alternativa 35 - Ci�ncia da computa��o ENADE 2014"
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
    'Descarrega o atual formul�rio da mem�ria
    Unload Me

     'O formul�rio final ser� exibido
     frm_final.Show
End Sub



Private Sub cmd_finalizarQA35_Click()
    'chamar� a sub que registrar� o que foi respondido no formul�rio
    Call acumuloQA35
    
End Sub


Private Sub opt_altAQA35_Click()
    'A op��o A armazenar� a letra na vari�rial
    Q(35) = "A"
End Sub

Private Sub opt_altBQA35_Click()
    'A op��o B armazenar� a letra na vari�rial
    Q(35) = "B"
End Sub

Private Sub opt_altCQA35_Click()
    'A op��o B armazenar� a letra na vari�rial
    Q(35) = "C"
End Sub

Private Sub opt_altDQA35_Click()
    'A op��o D armazenar� a letra na vari�rial
    Q(35) = "D"
End Sub

Private Sub opt_altEQA35_Click()
    'A op��o E armazenar� a letra na vari�rial
    Q(35) = "E"
End Sub
Sub acumuloQA35()

i = 35
        'Mostra a resposta correta ao usu�rio
        resp_QA35.Visible = True
        
        'Caso a resposta seja "E"...
        If Q(i) = "E" Then
            'O ac�mulo de acertos ser� incrementado em um valor
            acmAcertos = acmAcertos + 1
            'Informa o acerto ao usu�rio
            lbl_acerto.Visible = True
        'Caso contr�rio...
        Else
            'Caso n�o haja resposta
            If Q(i) = "NDA" Then
                'A quest�o se manter� "NDA"
                Q(i) = "NDA"
            'Caso haja, mas n�o correta...
            Else
                'O ac�mulo de erros ser� incrementado em um valor
                acmErros = acmErros + 1
            End If
            'Informa o erro ao usu�rio
            lbl_erro.Visible = True
        End If
        
        'Inativando todos os bot�es ap�s o usu�rio registrar a resposta
        opt_altAQA35.Enabled = False
        opt_altBQA35.Enabled = False
        opt_altCQA35.Enabled = False
        opt_altDQA35.Enabled = False
        opt_altEQA35.Enabled = False
        cmd_finalizarQA35.Enabled = False
        
        'A resposta ser� registrada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 42).Value = Q(i)
        
End Sub

Private Sub UserForm_Activate()

'La�o para altera��es no formul�rio ativo
With Me
    'Definindo a altura interna da barra de rolagem
    .ScrollHeight = .InsideHeight * 1.07
End With

End Sub

