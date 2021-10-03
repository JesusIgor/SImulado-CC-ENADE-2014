VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA7 
   Caption         =   "           Quest�o alternativa 7 - Ci�ncia da computa��o ENADE 2014"
   ClientHeight    =   10560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6090
   OleObjectBlob   =   "frm_QA7.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_fecharQA7_Click()
    'Descarrega o atual formul�rio da mem�ria
    Unload Me
    'Caso o usu�rio tenha pressionado o bot�o "pr�ximo"...
    If verifi = 1 Then
        'Ser� exibido a pr�xima quest�o
        frm_QA8.Show
    End If
    
    'Caso o usu�rio tenha pressionado o bot�o "finalizar"...
    If verifi = 2 Then
        'O formul�rio final ser� exibido
        frm_final.Show
    End If
End Sub



Private Sub cmd_finalizarQA7_Click()
    'chamar� a sub que registrar� o que foi respondido no formul�rio
    Call acumuloQ1
    'Verifica se o usu�rio pressionou o bot�o de finalizar
    verifi = 2
    
End Sub

Private Sub cmd_proxQA8_Click()

    'chamar� a sub que registrar� o que foi respondido no formul�rio
    Call acumuloQ1
    'Verifica se o usu�rio pressionou o bot�o de pr�ximo
    verifi = 1
   
End Sub


Private Sub opt_altAQA7_Click()
    'A op��o A armazenar� a letra na vari�rial
    Q(7) = "A"
End Sub

Private Sub opt_altBQA7_Click()
    'A op��o B armazenar� a letra na vari�rial
    Q(7) = "B"
End Sub

Private Sub opt_altCQA7_Click()
    'A op��o B armazenar� a letra na vari�rial
    Q(7) = "C"
End Sub

Private Sub opt_altDQA7_Click()
    'A op��o D armazenar� a letra na vari�rial
    Q(7) = "D"
End Sub

Private Sub opt_altEQA7_Click()
    'A op��o E armazenar� a letra na vari�rial
    Q(7) = "E"
End Sub
Sub acumuloQ1()

i = 7
        'Mostra a resposta correta ao usu�rio
        resp_QA7.Visible = True
        
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
        opt_altAQA7.Enabled = False
        opt_altBQA7.Enabled = False
        opt_altCQA7.Enabled = False
        opt_altDQA7.Enabled = False
        opt_altEQA7.Enabled = False
        cmd_proxQA8.Enabled = False
        cmd_finalizarQA7.Enabled = False
        
        'A resposta ser� registrada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 11).Value = Q(i)
        
End Sub

Private Sub UserForm_Activate()

'La�o para altera��es no formul�rio ativo
With Me
    'Definindo a altura interna da barra de rolagem
    .ScrollHeight = .InsideHeight * 1.36
End With

End Sub
