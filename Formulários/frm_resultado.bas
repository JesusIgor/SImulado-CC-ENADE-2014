VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_resultado 
   Caption         =   "                                                                                    Resultados -  Ci�ncia da computa��o ENADE 2014"
   ClientHeight    =   8985.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12045
   OleObjectBlob   =   "frm_resultado.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_resultado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Quando o bot�o fechar for clicado...
Private Sub cmd_fecharResultado_Click()
    'O formul�rio ser� descarregado da mem�ria
    Unload Me
End Sub



'Quando o formul�rio for ativado...
Private Sub UserForm_Activate()

'(La�o para modifica��o do label)
With lbl_resAcertos
    'Os acertos ser�o armazenados dentro do label
    lbl_resAcertos.Caption = acmAcertos
End With

'(La�o para modifica��o do label)
With lbl_resErros
    'Os erros ser�o armazenados dentro do label
    lbl_resErros.Caption = acmErros
End With

'(La�o para modifica��o do label)
With lbl_resBrancos
    'As respostas em branco ser�o armazenadas dentro do label
    lbl_resBrancos.Caption = acmBrancos
End With

'(La�o para modifica��o do label)
With lbl_total
    'O total de 40 quest�es alternativas ser� armazenado dentro do label
    lbl_total.Caption = 40
End With

'(La�o para modifica��o do label)
With lbl_respondidas
    'A soma de erros e acertos ser� armazenada no label de total respondido
    lbl_respondidas.Caption = acmAcertos + acmErros + Dvazio
End With

'(La�o para modifica��o do label)
With lbl_nomeResultado
    'O nome informado no in�cio do simulado ir� aparecer no topo do formul�rio
    lbl_nomeResultado.Caption = nome
End With

'(La�o para modifica��o do label)
With lbl_porcentagem
    'O desempenho do candidato aparecer� no canto esquerdo inferior do formul�rio
    lbl_porcentagem.Caption = desempenho
End With

'(La�o para modifica��o do label)
With lbl_dissertBrancos
    'O n�mero de quest�es dissertativas em branco ir� aparecer no label
    lbl_dissertBrancos.Caption = acmDissertBrancos
End With

'(La�o para modifica��o do label)
With lbl_dissertResp
    'O n�mero de quest�es dissertativas em branco ir� aparecer no label
    lbl_dissertResp.Caption = Dvazio
End With
'Fim da sub
End Sub
