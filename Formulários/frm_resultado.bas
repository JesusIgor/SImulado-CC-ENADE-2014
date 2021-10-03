VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_resultado 
   Caption         =   "                                                                                    Resultados -  Ciência da computação ENADE 2014"
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
'Quando o botão fechar for clicado...
Private Sub cmd_fecharResultado_Click()
    'O formulário será descarregado da memória
    Unload Me
End Sub



'Quando o formulário for ativado...
Private Sub UserForm_Activate()

'(Laço para modificação do label)
With lbl_resAcertos
    'Os acertos serão armazenados dentro do label
    lbl_resAcertos.Caption = acmAcertos
End With

'(Laço para modificação do label)
With lbl_resErros
    'Os erros serão armazenados dentro do label
    lbl_resErros.Caption = acmErros
End With

'(Laço para modificação do label)
With lbl_resBrancos
    'As respostas em branco serão armazenadas dentro do label
    lbl_resBrancos.Caption = acmBrancos
End With

'(Laço para modificação do label)
With lbl_total
    'O total de 40 questões alternativas será armazenado dentro do label
    lbl_total.Caption = 40
End With

'(Laço para modificação do label)
With lbl_respondidas
    'A soma de erros e acertos será armazenada no label de total respondido
    lbl_respondidas.Caption = acmAcertos + acmErros + Dvazio
End With

'(Laço para modificação do label)
With lbl_nomeResultado
    'O nome informado no início do simulado irá aparecer no topo do formulário
    lbl_nomeResultado.Caption = nome
End With

'(Laço para modificação do label)
With lbl_porcentagem
    'O desempenho do candidato aparecerá no canto esquerdo inferior do formulário
    lbl_porcentagem.Caption = desempenho
End With

'(Laço para modificação do label)
With lbl_dissertBrancos
    'O número de questões dissertativas em branco irá aparecer no label
    lbl_dissertBrancos.Caption = acmDissertBrancos
End With

'(Laço para modificação do label)
With lbl_dissertResp
    'O número de questões dissertativas em branco irá aparecer no label
    lbl_dissertResp.Caption = Dvazio
End With
'Fim da sub
End Sub
