Attribute VB_Name = "Módulo2"
Sub Respostas()
Attribute Respostas.VB_Description = "Esta macro direciona a planilha para a guia de respostas"
Attribute Respostas.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Respostas Macro
' Esta macro direciona a planilha para a guia de respostas
'

'
    Sheets("Respostas").Select
End Sub
Sub Gabarito()
Attribute Gabarito.VB_Description = "Esta Macro direciona a planilha para a guia de gabarito"
Attribute Gabarito.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Gabarito Macro
' Esta Macro direciona a planilha para a guia de gabarito
'

'
    Sheets("Gabarito").Select
    Range("A2:U2").Select
End Sub
