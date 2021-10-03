VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QD2 
   Caption         =   "                                                                                          Quest�o dissertativa 2 - Ci�ncia da computa��o ENADE 2014"
   ClientHeight    =   8145
   ClientLeft      =   120
   ClientTop       =   2460
   ClientWidth     =   13440
   OleObjectBlob   =   "frm_QD2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QD2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Quando o bot�o finalizar for clicado...
Private Sub cmd_finalizarQD2_Click()
    'Vari�vel para verificar qual bot�o o usu�rio pressionou
    verifi = 2
    'Chamar a fun��o de caso para verificar se a resposta est� vazia
    Call VazioQD2
    
End Sub

'Quando o bot�o for clicado...
Private Sub cmd_naoQD2_Click()
    'O frame exibido ficar� invis�vel
    frameQD2.Visible = False
    
End Sub

'Quando o bot�o pr�ximo for clicado...
Private Sub cmd_ProxQA1_Click()
    'Vari�vel para verificar qual bot�o o usu�rio pressionou
    verifi = 1
    'Chamar a fun��o de caso para verificar se a resposta est� vazia
    Call VazioQD2
    
'Fim da sub
End Sub

'Caso o usu�rio opte por deixar a caixa de texto em branco...
Private Sub cmd_simQD2_Click()
        
        
        'A resposta "Em branco!" ser� armazenada na planilha do excel
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 3).Value = "Em branco!"
        
        'Mensagem de iforma��o para o usu�rio
        MsgBox ("As quest�es dissertativas ser�o corrigiadas posteriormente!")
        
        'O forul�rio ser� descarregado da mem�ria
        Unload frm_QD2
        
        'Caso o usu�rio tenha pressionado o bot�o "pr�ximo"...
        If verifi = 1 Then
        'A pr�xima quest�o aparecer�
        frm_QA1.Show
        End If
        
        'Caso o usu�rio tenha pressionado o bot�o "finalizar"...
        If verifi = 2 Then
        'O formul�rio final aparecer�
        frm_final.Show
        End If
End Sub


'Quando o usu�rio pressionar uma tecla...
Private Sub txt_QD2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Caso seja a tecla 13 (Enter)...
    If KeyCode = 13 Then
        'Ser� executada uma quebra de linha
        txt_QD2.Text = txt_QD2.Text & Chr(13)
    End If
'Fim da sub
End Sub

Sub VazioQD2()

    'Caso a caixa de texto esteja vazia...
    If txt_QD2 = "" Then
        'O frame de verifica��o ficar� vis�vel
        frameQD2.Visible = True
    'Caso contr�rio...
    Else
        'A resposta ser� armazenada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 4).Value = txt_QD2.Text
        
        'Ac�mulo de dissertativas respondidas
        Dvazio = Dvazio + 1
        
        'Mensagem de iforma��o para o usu�rio
        MsgBox ("As quest�es dissertativas ser�o corrigiadas posteriormente!")
        
        'O formul�rio ser� descarregado da mem�ria
        Unload Me
        
        'Se o usu�rio apertar o bot�o "pr�ximo"...
        If verifi = 1 Then
        'A pr�xima quest�o aparecer�
        frm_QA1.Show
        End If
        
        'Se o usu�rio apertar o bot�o "finalizar"...
        If verifi = 2 Then
        'O formul�rio final aparecer�
        frm_final.Show
        End If
        
    End If
End Sub





