VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA31 
   Caption         =   "                                                   Quest�o alternativa 31 - Ci�ncia da computa��o ENADE 2014"
   ClientHeight    =   10530
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   10125
   OleObjectBlob   =   "frm_QA31.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_fecharQA31_Click()
    'Descarrega o atual formul�rio da mem�ria
    Unload Me
    'Caso o usu�rio tenha pressionado o bot�o "pr�ximo"...
    If verifi = 1 Then
        'Ser� exibido a pr�xima quest�o
        frm_QA32.Show
    End If
    
    'Caso o usu�rio tenha pressionado o bot�o "finalizar"...
    If verifi = 2 Then
        'O formul�rio final ser� exibido
        frm_final.Show
    End If
End Sub



Private Sub cmd_finalizarQA31_Click()
    'chamar� a sub que registrar� o que foi respondido no formul�rio
    Call acumuloQA31
    'Verifica se o usu�rio pressionou o bot�o de finalizar
    verifi = 2
    
End Sub

Private Sub cmd_proxQA32_Click()

    'chamar� a sub que registrar� o que foi respondido no formul�rio
    Call acumuloQA31
    'Verifica se o usu�rio pressionou o bot�o de pr�ximo
    verifi = 1
   
End Sub


Private Sub opt_altAQA31_Click()
    'A op��o A armazenar� a letra na vari�rial
    Q(31) = "A"
End Sub

Private Sub opt_altBQA31_Click()
    'A op��o B armazenar� a letra na vari�rial
    Q(31) = "B"
End Sub

Private Sub opt_altCQA31_Click()
    'A op��o B armazenar� a letra na vari�rial
    Q(31) = "C"
End Sub

Private Sub opt_altDQA31_Click()
    'A op��o D armazenar� a letra na vari�rial
    Q(31) = "D"
End Sub

Private Sub opt_altEQA31_Click()
    'A op��o E armazenar� a letra na vari�rial
    Q(31) = "E"
End Sub
Sub acumuloQA31()

i = 31
        'Mostra a resposta correta ao usu�rio
        resp_QA31.Visible = True
        
        'Caso a resposta seja "C"...
        If Q(i) = "C" Then
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
        opt_altAQA31.Enabled = False
        opt_altBQA31.Enabled = False
        opt_altCQA31.Enabled = False
        opt_altDQA31.Enabled = False
        opt_altEQA31.Enabled = False
        cmd_proxQA32.Enabled = False
        cmd_finalizarQA31.Enabled = False
        
        'A resposta ser� registrada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 38).Value = Q(i)
        
End Sub

Private Sub UserForm_Activate()

'La�o para altera��es no formul�rio ativo
With Me
    'Definindo a altura interna da barra de rolagem
    .ScrollHeight = .InsideHeight * 1.2
End With

End Sub

