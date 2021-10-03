VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA17 
   Caption         =   "                                                                              Quest�o alternativa 17 - Ci�ncia da computa��o ENADE 2014"
   ClientHeight    =   6885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12705
   OleObjectBlob   =   "frm_QA17.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_fecharQA17_Click()
    'Descarrega o atual formul�rio da mem�ria
    Unload Me
    'Caso o usu�rio tenha pressionado o bot�o "pr�ximo"...
    If verifi = 1 Then
        'Ser� exibido a pr�xima quest�o
        frm_QA18.Show
    End If
    
    'Caso o usu�rio tenha pressionado o bot�o "finalizar"...
    If verifi = 2 Then
        'O formul�rio final ser� exibido
        frm_final.Show
    End If
End Sub

Private Sub cmd_finalizarQA17_Click()
    'chamar� a sub que registrar� o que foi respondido no formul�rio
    Call acumuloQA17
    'Verifica se o usu�rio pressionou o bot�o de pr�ximo
    verifi = 2
    
End Sub

Private Sub cmd_proxQA18_Click()
    
    'chamar� a sub que registrar� o que foi respondido no formul�rio
    Call acumuloQA17
    'Verifica se o usu�rio pressionou o bot�o de pr�ximo
    verifi = 1
    
End Sub

Private Sub opt_altAQA17_Click()
    'A op��o A armazenar� a letra na vari�rial
    Q(17) = "A"
End Sub

Private Sub opt_altBQA17_Click()
    'A op��o B armazenar� a letra na vari�rial
    Q(17) = "B"
End Sub

Private Sub opt_altCQA17_Click()
    'A op��o C armazenar� a letra na vari�rial
    Q(17) = "C"
End Sub
Private Sub opt_altDQA17_Click()
    'A op��o D armazenar� a letra na vari�rial
    Q(17) = "D"
End Sub
Private Sub opt_altEQA17_Click()
    'A op��o E armazenar� a letra na vari�rial
    Q(17) = "E"
End Sub

Sub acumuloQA17()
i = 17
        'Mostra a resposta correta ao usu�rio
        resp_QA17.Visible = True
        
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
        opt_altAQA17.Enabled = False
        opt_altBQA17.Enabled = False
        opt_altCQA17.Enabled = False
        opt_altDQA17.Enabled = False
        opt_altEQA17.Enabled = False
        cmd_proxQA18.Enabled = False
        cmd_finalizarQA17.Enabled = False
        'Mostra dentro da planilha "respostas" qual foi a alternativa escolhida pelo usuario
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 24).Value = Q(i)

End Sub
