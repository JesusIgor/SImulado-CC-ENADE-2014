VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_inicio 
   Caption         =   "                                                                                                                            Início -  Ciência da computação ENADE 2014"
   ClientHeight    =   9105.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15345
   OleObjectBlob   =   "frm_inicio.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Caso o usuário clique no botão "Instruções"
Private Sub cmd_instrucoes_Click()
    'O formulário de instruções aparecerá
    frm_instrucoes.Show
    
End Sub

'Caso o usuário clique no botão "Próximo"
Private Sub cmd_proxInicio_Click()
    'O formulário inicial irá descarregar da memória
    Unload Me
    
    'O formulário de nome irá aparecer
    frm_nome.Show
End Sub

'Caso o usuário clique no botão "Sobre"
Private Sub cmd_sobre_Click()
    
    'O formulário "Sobre" irá abrir
    frm_sobre.Show
    
End Sub




