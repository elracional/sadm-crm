VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Salir 
   Caption         =   "UserForm1"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5835
   OleObjectBlob   =   "Salir.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Salir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
End
End Sub

Private Sub CommandButton2_Click()
Salir.Hide
frmPrincipal.Show
End Sub
