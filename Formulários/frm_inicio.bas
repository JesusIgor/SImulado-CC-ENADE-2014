VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_inicio 
   Caption         =   "                                                                                                                            In�cio -  Ci�ncia da computa��o ENADE 2014"
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
'Caso o usu�rio clique no bot�o "Instru��es"
Private Sub cmd_instrucoes_Click()
    'O formul�rio de instru��es aparecer�
    frm_instrucoes.Show
    
End Sub

'Caso o usu�rio clique no bot�o "Pr�ximo"
Private Sub cmd_proxInicio_Click()
    'O formul�rio inicial ir� descarregar da mem�ria
    Unload Me
    
    'O formul�rio de nome ir� aparecer
    frm_nome.Show
End Sub

'Caso o usu�rio clique no bot�o "Sobre"
Private Sub cmd_sobre_Click()
    
    'O formul�rio "Sobre" ir� abrir
    frm_sobre.Show
    
End Sub




