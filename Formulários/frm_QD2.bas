VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QD2 
   Caption         =   "                                                                                          Questão dissertativa 2 - Ciência da computação ENADE 2014"
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
'Quando o botão finalizar for clicado...
Private Sub cmd_finalizarQD2_Click()
    'Variável para verificar qual botão o usuário pressionou
    verifi = 2
    'Chamar a função de caso para verificar se a resposta está vazia
    Call VazioQD2
    
End Sub

'Quando o botão for clicado...
Private Sub cmd_naoQD2_Click()
    'O frame exibido ficará invisível
    frameQD2.Visible = False
    
End Sub

'Quando o botão próximo for clicado...
Private Sub cmd_ProxQA1_Click()
    'Variável para verificar qual botão o usuário pressionou
    verifi = 1
    'Chamar a função de caso para verificar se a resposta está vazia
    Call VazioQD2
    
'Fim da sub
End Sub

'Caso o usuário opte por deixar a caixa de texto em branco...
Private Sub cmd_simQD2_Click()
        
        
        'A resposta "Em branco!" será armazenada na planilha do excel
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 3).Value = "Em branco!"
        
        'Mensagem de iformação para o usuário
        MsgBox ("As questões dissertativas serão corrigiadas posteriormente!")
        
        'O forulário será descarregado da memória
        Unload frm_QD2
        
        'Caso o usuário tenha pressionado o botão "próximo"...
        If verifi = 1 Then
        'A próxima questão aparecerá
        frm_QA1.Show
        End If
        
        'Caso o usuário tenha pressionado o botão "finalizar"...
        If verifi = 2 Then
        'O formulário final aparecerá
        frm_final.Show
        End If
End Sub


'Quando o usuário pressionar uma tecla...
Private Sub txt_QD2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Caso seja a tecla 13 (Enter)...
    If KeyCode = 13 Then
        'Será executada uma quebra de linha
        txt_QD2.Text = txt_QD2.Text & Chr(13)
    End If
'Fim da sub
End Sub

Sub VazioQD2()

    'Caso a caixa de texto esteja vazia...
    If txt_QD2 = "" Then
        'O frame de verificação ficará visível
        frameQD2.Visible = True
    'Caso contrário...
    Else
        'A resposta será armazenada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 4).Value = txt_QD2.Text
        
        'Acúmulo de dissertativas respondidas
        Dvazio = Dvazio + 1
        
        'Mensagem de iformação para o usuário
        MsgBox ("As questões dissertativas serão corrigiadas posteriormente!")
        
        'O formulário será descarregado da memória
        Unload Me
        
        'Se o usuário apertar o botão "próximo"...
        If verifi = 1 Then
        'A próxima questão aparecerá
        frm_QA1.Show
        End If
        
        'Se o usuário apertar o botão "finalizar"...
        If verifi = 2 Then
        'O formulário final aparecerá
        frm_final.Show
        End If
        
    End If
End Sub





