VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_final 
   Caption         =   "                                                                                                                                     Final -  Ci�ncia da computa��o ENADE 2014"
   ClientHeight    =   9810.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15435
   OleObjectBlob   =   "frm_final.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_final"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_fechar_Click()
    'Encerra as a��es de todoss os formul�rios
    End
End Sub

Private Sub cmd_gabarito_Click()
    'Mostra o formul�rio de gabarito para o usu�rio
    frm_gabarito.Show
    
End Sub

Private Sub cmd_resultado_Click()
    'Mostra o formul�rio de gabarito para o usu�rio
    frm_resultado.Show
End Sub

Private Sub cmd_sair_Click()
    'Sai da applica��o do Excel
    Application.Quit
End Sub

Private Sub cmd_voltarInicio_Click()
    'Descarrega este formul�rio da mem�ria
    Unload Me
    
    'Retorna o usu�rio para o formul�rio inicial
    frm_inicio.Show
    
End Sub

Private Sub UserForm_Activate()

acmBrancos = 35 - (acmAcertos + acmErros)
acmDissertBrancos = 5 - Dvazio
acmBrancos = acmBrancos + acmDissertBrancos

'Registra o n�mero de quest�es respondidas na guia de respostas
ThisWorkbook.Worksheets("Respostas").Cells(linha, 43).Value = acmAcertos + acmErros + Dvazio

'Registra o n�mero de acertos na guia de respostas
ThisWorkbook.Worksheets("Respostas").Cells(linha, 44).Value = acmAcertos

'Registra o n�mero de erros na guia de respostas
ThisWorkbook.Worksheets("Respostas").Cells(linha, 45).Value = acmErros

'Registra o n�mero de quest�es em branco na guia de respostas
ThisWorkbook.Worksheets("Respostas").Cells(linha, 46).Value = acmBrancos

'Registra o n�mero de quest�es anuladas na guia de respostas
ThisWorkbook.Worksheets("Respostas").Cells(linha, 47).Value = 0

'Registra o n�mero de dissertativas na guia de respostas
ThisWorkbook.Worksheets("Respostas").Cells(linha, 48).Value = 5

'Calcula o desempenho do usu�rio
desempenho = (acmAcertos * 100) / 35
End Sub
