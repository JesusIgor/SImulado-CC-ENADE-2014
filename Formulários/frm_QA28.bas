VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA28 
   Caption         =   "                                                                                                     Quest�o alternativa 28 - Ci�ncia da computa��o ENADE 2014"
   ClientHeight    =   6195
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   15015
   OleObjectBlob   =   "frm_QA28.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_fecharQA28_Click()
    'Descarrega o atual formul�rio da mem�ria
    Unload Me
    'Caso o usu�rio tenha pressionado o bot�o "pr�ximo"...
    If verifi = 1 Then
        'Ser� exibido a pr�xima quest�o
        frm_QA29.Show
    End If
    
    'Caso o usu�rio tenha pressionado o bot�o "finalizar"...
    If verifi = 2 Then
        'O formul�rio final ser� exibido
        frm_final.Show
    End If
End Sub



Private Sub cmd_finalizarQA28_Click()
    'chamar� a sub que registrar� o que foi respondido no formul�rio
    Call acumuloQA28
    'Verifica se o usu�rio pressionou o bot�o de finalizar
    verifi = 2
    
End Sub

Private Sub cmd_proxQA29_Click()

    'chamar� a sub que registrar� o que foi respondido no formul�rio
    Call acumuloQA28
    'Verifica se o usu�rio pressionou o bot�o de pr�ximo
    verifi = 1
   
End Sub


Private Sub opt_altAQA28_Click()
    'A op��o A armazenar� a letra na vari�rial
    Q(28) = "A"
End Sub

Private Sub opt_altBQA28_Click()
    'A op��o B armazenar� a letra na vari�rial
    Q(28) = "B"
End Sub

Private Sub opt_altCQA28_Click()
    'A op��o B armazenar� a letra na vari�rial
    Q(28) = "C"
End Sub

Private Sub opt_altDQA28_Click()
    'A op��o D armazenar� a letra na vari�rial
    Q(28) = "D"
End Sub

Private Sub opt_altEQA28_Click()
    'A op��o E armazenar� a letra na vari�rial
    Q(28) = "E"
End Sub
Sub acumuloQA28()

i = 28
        'Mostra a resposta correta ao usu�rio
        resp_QA28.Visible = True
        
        'Caso a resposta seja "A"...
        If Q(i) = "A" Then
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
        opt_altAQA28.Enabled = False
        opt_altBQA28.Enabled = False
        opt_altCQA28.Enabled = False
        opt_altDQA28.Enabled = False
        opt_altEQA28.Enabled = False
        cmd_proxQA29.Enabled = False
        cmd_finalizarQA28.Enabled = False
        
        'A resposta ser� registrada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 35).Value = Q(i)
        
End Sub


