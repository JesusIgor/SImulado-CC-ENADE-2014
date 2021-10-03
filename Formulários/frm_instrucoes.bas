VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_instrucoes 
   Caption         =   "                                                    Instruções - Ciência da computação ENADE 2014"
   ClientHeight    =   10170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9210.001
   OleObjectBlob   =   "frm_instrucoes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_instrucoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_fecharInstrucoes_Click()
    Unload Me
    
End Sub

Private Sub UserForm_Activate()

With Me
    .ScrollHeight = .InsideHeight * 1.3
End With

End Sub


