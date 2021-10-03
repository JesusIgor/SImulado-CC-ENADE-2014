VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_final 
   Caption         =   "                                                                                                                                     Final -  Ciência da computação ENADE 2014"
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
    'Encerra as ações de todoss os formulários
    End
End Sub

Private Sub cmd_gabarito_Click()
    'Mostra o formulário de gabarito para o usuário
    frm_gabarito.Show
    
End Sub

Private Sub cmd_resultado_Click()
    'Mostra o formulário de gabarito para o usuário
    frm_resultado.Show
End Sub

Private Sub cmd_sair_Click()
    'Sai da applicação do Excel
    Application.Quit
End Sub

Private Sub cmd_voltarInicio_Click()
    'Descarrega este formulário da memória
    Unload Me
    
    'Retorna o usuário para o formulário inicial
    frm_inicio.Show
    
End Sub

Private Sub UserForm_Activate()

acmBrancos = 35 - (acmAcertos + acmErros)
acmDissertBrancos = 5 - Dvazio
acmBrancos = acmBrancos + acmDissertBrancos

'Registra o número de questões respondidas na guia de respostas
ThisWorkbook.Worksheets("Respostas").Cells(linha, 43).Value = acmAcertos + acmErros + Dvazio

'Registra o número de acertos na guia de respostas
ThisWorkbook.Worksheets("Respostas").Cells(linha, 44).Value = acmAcertos

'Registra o número de erros na guia de respostas
ThisWorkbook.Worksheets("Respostas").Cells(linha, 45).Value = acmErros

'Registra o número de questões em branco na guia de respostas
ThisWorkbook.Worksheets("Respostas").Cells(linha, 46).Value = acmBrancos

'Registra o número de questões anuladas na guia de respostas
ThisWorkbook.Worksheets("Respostas").Cells(linha, 47).Value = 0

'Registra o número de dissertativas na guia de respostas
ThisWorkbook.Worksheets("Respostas").Cells(linha, 48).Value = 5

'Calcula o desempenho do usuário
desempenho = (acmAcertos * 100) / 35
End Sub
