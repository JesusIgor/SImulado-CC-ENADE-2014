VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_QD3 
   Caption         =   "                                                                                         Questão dissertativa 3 - Ciência da computação ENADE 2014"
   ClientHeight    =   9990.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13440
   OleObjectBlob   =   "frm_QD3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_QD3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Quando o botão finalizar for clicado...
Private Sub cmd_finalizarQD3_Click()
    'Variável para verificar qual botão o usuário pressionou
    verifi = 2
    'Chamar a função de caso para verificar se a resposta está vazia
    Call VazioQD3
    
End Sub

'Quando o botão for clicado...
Private Sub cmd_naoQD3_Click()
    'O frame exibido ficará invisível
    frameQD3.Visible = False
    
End Sub

'Quando o botão próximo for clicado...
Private Sub cmd_ProxQD4_Click()
    'Variável para verificar qual botão o usuário pressionou
    verifi = 1
    'Chamar a função de caso para verificar se a resposta está vazia
    Call VazioQD3
    
'Fim da sub
End Sub

'Caso o usuário opte por deixar a caixa de texto em branco...
Private Sub cmd_simQD3_Click()
        
        
        'A resposta "Em branco!" será armazenada na planilha do excel
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 13).Value = "Em branco!"
        
        'Mensagem de iformação para o usuário
        MsgBox ("As questões dissertativas serão corrigiadas posteriormente!")
        
        'O forulário será descarregado da memória
        Unload frm_QD3
        
        'Caso o usuário tenha pressionado o botão "próximo"...
        If verifi = 1 Then
            'A questão 25 aparecerá
            frm_QD4.Show
        End If
        
        'Caso o usuário tenha pressionado o botão "finalizar"...
        If verifi = 2 Then
            'O formulário final aparecerá
            frm_final.Show
        End If
        
End Sub


'Quando o usuário pressionar uma tecla...
Private Sub txt_QD3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Caso seja a tecla 13 (Enter)...
    If KeyCode = 13 Then
        'Será executada uma quebra de linha
        txt_QD3.Text = txt_QD3.Text & Chr(13)
    End If
'Fim da sub
End Sub

Sub VazioQD3()

    'Caso a caixa de texto esteja vazia...
    If txt_QD3 = "" Then
        'O frame de verificação ficará visível
        frameQD3.Visible = True
    'Caso contrário...
    Else
        'A resposta será armazenada na planilha
        ThisWorkbook.Worksheets("Respostas").Cells(linha, 13).Value = txt_QD3.Text
        
        'Acúmulo de dissertativas respondidas
        Dvazio = Dvazio + 1
        
        'Mensagem de iformação para o usuário
        MsgBox ("As questões dissertativas serão corrigiadas posteriormente!")
        
        'O formulário será descarregado da memória
        Unload Me
        
        'Se o usuário apertar o botão "próximo"...
        If verifi = 1 Then
        'A próxima questão aparecerá
        frm_QD4.Show
        End If
        
        'Se o usuário apertar o botão "finalizar"...
        If verifi = 2 Then
        'O formulário final aparecerá
        frm_final.Show
        End If
                
                
    End If
End Sub

