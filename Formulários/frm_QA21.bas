VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QA21 
   Caption         =   "                                  Quest�o alternativa 21 - Ci�ncia da computa��o ENADE 2014"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8550.001
   OleObjectBlob   =   "frm_QA21.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QA21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_fecharQA21_Click()
    'Descarrega o atual formul�rio da mem�ria
    Unload Me
    'Caso o usu�rio tenha pressionado o bot�o "pr�ximo"...
    If verifi = 1 Then
        'Ser� exibido a pr�xima quest�o
        frm_QA22.Show
    End If
    
    'Caso o usu�rio tenha pressionado o bot�o "finalizar"...
    If verifi = 2 Then
        'O formul�rio final ser� exibido
        frm_final.Show
    End If
End Sub



Private Sub cmd_finalizarQA21_Click()
    'chamar� a sub que registrar� o que foi respondido no formul�rio
    Call acumuloQA21
    'Verifica se o usu�rio pressionou o bot�o de finalizar
    verifi = 2
    
End Sub

Private Sub cmd_proxQA22_Click()

    'chamar� a sub que registrar� o que foi respondido no formul�rio
    Call acumuloQA21
    'Verifica se o usu�rio pressionou o bot�o de pr�ximo
    verifi = 1
   
End Sub


Private Sub opt_altAQA21_Click()
    'A op��o A armazenar� a letra na vari�rial
    Q(21) = "A"
End Sub

Private Sub opt_altBQA21_Click()
    'A op��o B armazenar� a letra na vari�rial
    Q(21) = "B"
End Sub

Private Sub opt_altCQA21_Click()
    'A op��o B armazenar� a letra na vari�rial
    Q(21) = "C"
End Sub

Private Sub opt_altDQA21_Click()
    'A op��o D armazenar� a letra na vari�rial
    Q(21) = "D"
End Sub

Private Sub opt_altEQA21_Click()
    'A op��o E armazenar� a letra na vari�rial
    Q(21) = "E"
End Sub
Sub acumuloQA21()

i = 21
        'Mostra a resposta correta ao usu�rio
        resp_QA21.Visible = True
        
        'Caso a resposta seja "B"...
        If Q(i) = "B" Then
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
        opt_altAQA21.Enabled = False
        opt_altBQA21.Enabled = False
        opt_altCQA21.Enabled = False
        opt_altDQA21.Enabled = False
        opt_altEQA21.Enabled = False
        cmd_proxQA22.Enabled = False
        cmd_finalizarQA21.Enabled = False
        
        'A resposta ser� registrada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 28).Value = Q(i)
        
End Sub

Private Sub UserForm_Activate()
'La�o para altera��es no formul�rio ativo
With Me
    'Definindo a altura interna da barra de rolagem
    .ScrollHeight = .InsideHeight * 1#
End With
End Sub

