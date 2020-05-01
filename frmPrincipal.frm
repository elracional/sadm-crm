VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrincipal 
   Caption         =   "EXCELeINFO"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10620
   OleObjectBlob   =   "frmPrincipal.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
frmBuscar.Show
End Sub

Private Sub CommandButton2_Click()
frmAlta.Show
End Sub

Private Sub CommandButton3_Click()
Unload Me
End Sub

Private Sub CommandButton4_Click()
Do While TextBox1.Text = "hgguel" And TextBox2.Text = "123qweas"
 frmPrincipal.Hide
frmBuscar.Show
Loop
MsgBox "Usuario y/o Cotraseña icorrecto(s), vuelva a intentarlo"


With Me
.TextBox1 = ""
.TextBox2 = ""
End With

End Sub

Private Sub CommandButton5_Click()
Salir.Show
End Sub
