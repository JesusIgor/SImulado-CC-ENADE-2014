VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QD5 
   Caption         =   "                                                                                                 Quest�o dissertativa 5 - Ci�ncia da computa��o ENADE 2014"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13425
   OleObjectBlob   =   "frm_QD5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QD5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Quando o bot�o finalizar for clicado...
Private Sub cmd_finalizarQD5_Click()
    'Vari�vel para verificar qual bot�o o usu�rio pressionou
    verifi = 2
    'Chamar a fun��o de caso para verificar se a resposta est� vazia
    Call VazioQD5
    
End Sub

'Quando o bot�o for clicado...
Private Sub cmd_naoQD5_Click()
    'O frame exibido ficar� invis�vel
    frameQD5.Visible = False
    
End Sub

'Quando o bot�o pr�ximo for clicado...
Private Sub cmd_ProxQA9_Click()
    'Vari�vel para verificar qual bot�o o usu�rio pressionou
    verifi = 1
    'Chamar a fun��o de caso para verificar se a resposta est� vazia
    Call VazioQD5
    
'Fim da sub
End Sub

'Caso o usu�rio opte por deixar a caixa de texto em branco...
Private Sub cmd_simQD5_Click()
        
        
        'A resposta "Em branco!" ser� armazenada na planilha do excel
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 15).Value = "Em branco!"
        
        'Mensagem de iforma��o para o usu�rio
        MsgBox ("As quest�es dissertativas ser�o corrigiadas posteriormente!")
        
        'O forul�rio ser� descarregado da mem�ria
        Unload frm_QD5
        
        'Caso o usu�rio tiver apertado o bot�o pr�ximo
        If verifi = 1 Then
            'A quest�o 25 aparecer�
            frm_QA9.Show
        End If
        
        'Caso o usu�rio tiver apertado o bot�o finalizar
        If verifi = 2 Then
            'O formul�rio final aparecer�
            frm_final.Show
        End If

End Sub

'Quando o usu�rio pressionar uma tecla...
Private Sub txt_QD5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Caso seja a tecla 13 (Enter)...
    If KeyCode = 13 Then
        'Ser� executada uma quebra de linha
        txt_QD5.Text = txt_QD5.Text & Chr(13)
    End If
'Fim da sub
End Sub

Sub VazioQD5()

    'Caso a caixa de texto esteja vazia...
    If txt_QD5 = "" Then
        'O frame de verifica��o ficar� vis�vel
        frameQD5.Visible = True
    'Caso contr�rio...
    Else
        'A resposta ser� armazenada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 15).Value = txt_QD5.Text
        
        'Ac�mulo de dissertativas respondidas
        Dvazio = Dvazio + 1
        
        'Mensagem de iforma��o para o usu�rio
        MsgBox ("As quest�es dissertativas ser�o corrigiadas posteriormente!")
        
        'O formul�rio ser� descarregado da mem�ria
        Unload Me
        
        'Se o usu�rio apertar o bot�o "pr�ximo"...
        If verifi = 1 Then
        'A pr�xima quest�o aparecer�
        frm_QA9.Show
        End If
        
        'Se o usu�rio apertar o bot�o "finalizar"...
        If verifi = 2 Then
        'O formul�rio final aparecer�
        frm_final.Show
        End If
                
                
    End If
End Sub
