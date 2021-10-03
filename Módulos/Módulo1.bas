Attribute VB_Name = "Módulo1"
Sub Instrucoes()
Attribute Instrucoes.VB_Description = "Esta macro direciona a planilha ´para a guia de intruções"
Attribute Instrucoes.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Instrucoes Macro
' Esta macro direciona a planilha ´para a guia de intruções
'

'
    Sheets("Instruções").Select
End Sub
Sub Dissertativas()
Attribute Dissertativas.VB_Description = "Esta macro direciona a planilha para a guia de dissertativas"
Attribute Dissertativas.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Dissertativas Macro
' Esta macro direciona a planilha para a guia de dissertativas
'

'
    Sheets("Q1D").Select
End Sub
Sub Alternativas()
Attribute Alternativas.VB_Description = "Esta macro direciona a planilha para as questões alternativas"
Attribute Alternativas.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Alternativas Macro
' Esta macro direciona a planilha para as questões alternativas
'

'
    Sheets("Q1").Select
End Sub
