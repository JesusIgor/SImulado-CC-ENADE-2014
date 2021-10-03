VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA1 
   Caption         =   "                                                                                           Quest�o alternativa 1 - Ci�ncia da computa��o ENADE 2014"
   ClientHeight    =   9315.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13365
   OleObjectBlob   =   "frm_QA1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_fecharQA1_Click()
    'Descarrega o atual formul�rio da mem�ria
    Unload Me
    'Caso o usu�rio tenha pressionado o bot�o "pr�ximo"...
    If verifi = 1 Then
        'Ser� exibido a pr�xima quest�o
        frm_QA2.Show
    End If
    
    'Caso o usu�rio tenha pressionado o bot�o "finalizar"...
    If verifi = 2 Then
        'O formul�rio final ser� exibido
        frm_final.Show
    End If
End Sub



Private Sub cmd_finalizarQA1_Click()
    'chamar� a sub que registrar� o que foi respondido no formul�rio
    Call acumuloQA1
    'Verifica se o usu�rio pressionou o bot�o de finalizar
    verifi = 2
    
End Sub

Private Sub cmd_proxQA2_Click()

    'chamar� a sub que registrar� o que foi respondido no formul�rio
    Call acumuloQA1
    'Verifica se o usu�rio pressionou o bot�o de pr�ximo
    verifi = 1
   
End Sub


Private Sub opt_altAQA1_Click()
    'A op��o A armazenar� a letra na vari�rial
    Q(1) = "A"
End Sub

Private Sub opt_altBQA1_Click()
    'A op��o B armazenar� a letra na vari�rial
    Q(1) = "B"
End Sub

Private Sub opt_altCQA1_Click()
    'A op��o B armazenar� a letra na vari�rial
    Q(1) = "C"
End Sub

Private Sub opt_altDQA1_Click()
    'A op��o D armazenar� a letra na vari�rial
    Q(1) = "D"
End Sub

Private Sub opt_altEQA1_Click()
    'A op��o E armazenar� a letra na vari�rial
    Q(1) = "E"
End Sub
Sub acumuloQA1()

i = 1
        'Mostra a resposta correta ao usu�rio
        resp_QA1.Visible = True
        
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
        opt_altAQA1.Enabled = False
        opt_altBQA1.Enabled = False
        opt_altCQA1.Enabled = False
        opt_altDQA1.Enabled = False
        opt_altEQA1.Enabled = False
        cmd_ProxQA2.Enabled = False
        cmd_finalizarQA1.Enabled = False
        
        'A resposta ser� registrada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 5).Value = Q(i)
        
End Sub
