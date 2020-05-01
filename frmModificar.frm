VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmModificar 
   Caption         =   "EXCELeINFO - Modificar registros"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8265
   OleObjectBlob   =   "frmModificar.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmModificar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Actualizar el registro
Private Sub CommandButton1_Click()
For i = 1 To 8
    ActiveCell.Offset(0, i - 1).Value = Me.Controls("TextBox" & i).Value
Next i
Unload Me
End Sub
'
'Cerrar formulario
Private Sub CommandButton2_Click()
Unload Me
End Sub
'
'Llenar los cuadro de texto con los datos del registro elegido
Private Sub UserForm_Initialize()
For i = 1 To 8
    Me.Controls("TextBox" & i).Value = ActiveCell.Offset(0, i - 1).Value
Next i
End Sub

