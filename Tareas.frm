VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tareas 
   Caption         =   "UserForm1"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11805
   OleObjectBlob   =   "Tareas.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "tAREAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
MsgBox "Advertencia: Los datos se han enviado a la Base de Datos"
Dim Dato1 As String
Dim Dato2 As String
Dim Dato3 As String
Dim Dato4 As String
Dim Dato5 As String
Dim Dato6 As String
Dim Dato7 As String
Dim Dato8 As String
Dim Dato9 As String
Dim Dato10 As String
Dim Dato11 As String
Dim ultimaC As Double

Dato1 = TextBox1.Value
Dato2 = TextBox2.Value
Dato3 = TextBox3.Value
Dato4 = TextBox4.Value
Dato5 = TextBox5.Value
Dato6 = TextBox6.Value
Dato7 = TextBox7.Value
Dato8 = TextBox8.Value
Dato9 = TextBox9.Value
Dato10 = TextBox10.Value
Dato11 = TextBox11.Value

Set TramsRowRng = ThisWorkbook.Worksheets(3).Cells(1, 10).CurrentRegion
ultimaC = ActiveSheet.UsedRange.Row - 1 + ActiveSheet.UsedRange.Rows.Count

Cells(ultimaC + 1, 2) = Dato1
Cells(ultimaC + 1, 3) = Dato2
Cells(ultimaC + 1, 4) = Dato3
Cells(ultimaC + 1, 5) = Dato4
Cells(ultimaC + 1, 6) = Dato5
Cells(ultimaC + 1, 7) = Dato6
Cells(ultimaC + 1, 8) = Dato7
Cells(ultimaC + 1, 9) = Dato8
Cells(ultimaC + 1, 10) = Dato9
Cells(ultimaC + 1, 11) = Dato10
Cells(ultimaC + 1, 12) = Dato11
tAREAS.Hide
MsgBox "Los Datos han sido cargados correctamente"
frmBuscar.Show

End Sub



