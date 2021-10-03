VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA5 
   Caption         =   "                         Quest�o alternativa 5 - Ci�ncia da computa��o ENADE 2014"
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
    'Descarrega o atual formul�rio da mem�ria
    Unload Me
    'Caso o usu�rio tenha pressionado o bot�o "pr�ximo"...
    If verifi = 1 Then
        'Ser� exibido a pr�xima quest�o
        frm_QA6.Show
    End If
    
    'Caso o usu�rio tenha pressionado o bot�o "finalizar"...
    If verifi = 2 Then
        'O formul�rio final ser� exibido
        frm_final.Show
    End If
End Sub



Private Sub cmd_finalizarQA5_Click()
    'chamar� a sub que registrar� o que foi respondido no formul�rio
    Call acumuloQA5
    'Verifica se o usu�rio pressionou o bot�o de finalizar
    verifi = 2
    
End Sub

Private Sub cmd_proxQA6_Click()

    'chamar� a sub que registrar� o que foi respondido no formul�rio
    Call acumuloQA5
    'Verifica se o usu�rio pressionou o bot�o de pr�ximo
    verifi = 1
   
End Sub


Private Sub opt_altAQA5_Click()
    'A op��o A armazenar� a letra na vari�rial
    Q(5) = "A"
End Sub

Private Sub opt_altBQA5_Click()
    'A op��o B armazenar� a letra na vari�rial
    Q(5) = "B"
End Sub

Private Sub opt_altCQA5_Click()
    'A op��o B armazenar� a letra na vari�rial
    Q(5) = "C"
End Sub

Private Sub opt_altDQA5_Click()
    'A op��o D armazenar� a letra na vari�rial
    Q(5) = "D"
End Sub

Private Sub opt_altEQA5_Click()
    'A op��o E armazenar� a letra na vari�rial
    Q(5) = "E"
End Sub
Sub acumuloQA5()

i = 5
        'Mostra a resposta correta ao usu�rio
        resp_QA5.Visible = True
        
        'Caso a resposta seja "D"...
        If Q(i) = "D" Then
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
        
        'Desativando todos os bot�es ap�s o usu�rio registrar a resposta
        opt_altAQA5.Enabled = False
        opt_altBQA5.Enabled = False
        opt_altCQA5.Enabled = False
        opt_altDQA5.Enabled = False
        opt_altEQA5.Enabled = False
        cmd_proxQA6.Enabled = False
        cmd_finalizarQA5.Enabled = False
        
        'A resposta ser� registrada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 9).Value = Q(i)
        
End Sub

Private Sub UserForm_Activate()

'La�o para altera��es no formul�rio ativo
With Me
    'Definindo a altura interna da barra de rolagem
    .ScrollHeight = .InsideHeight * 1.7
End With

End Sub
