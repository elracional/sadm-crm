VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAlta 
   Caption         =   "EXCELeINFO - Agregar registros"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9225
   OleObjectBlob   =   "frmAlta.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmAlta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Alta de un registro
Private Sub CommandButton1_Click()
'Declaración de variables
'
Dim strTitulo As String
Dim Continuar As String
Dim TransRowRng As Range
Dim NewRow As Integer
Dim Limpiar As String
Dim Dato1 As String
Dim Dato2 As String
Dim Dato3 As String
Dim Dato4 As String
'
strTitulo = "EXCELeINFO"
'
Continuar = MsgBox("Dar de alta los datos?", vbYesNo + vbExclamation, strTitulo)
If Continuar = vbNo Then Exit Sub
'
Cuenta = Application.WorksheetFunction.CountIf(Range("A:A"), Me.txtID)
'
If Cuenta > 0 Then
    '
    MsgBox "El ID '" & Me.txtID & "' ya se encuentra registrado", vbExclamation, strTitulo
    '
Else
    '
    Set TransRowRng = ThisWorkbook.Worksheets("Hoja1").Cells(1, 1).CurrentRegion
    NewRow = TransRowRng.Rows.Count + 1
    With ThisWorkbook.Worksheets("Hoja1")
        .Cells(NewRow, 1).Value = Me.txtID
        .Cells(NewRow, 2).Value = Me.txtUsuario
        .Cells(NewRow, 3).Value = Me.txtDepartamento
        .Cells(NewRow, 4).Value = Me.txtPuesto
        .Cells(NewRow, 5).Value = Me.TextBox1
        .Cells(NewRow, 6).Value = Me.TextBox2
        .Cells(NewRow, 7).Value = Me.TextBox3
        .Cells(NewRow, 8).Value = Me.TextBox4
    End With
    '
    MsgBox "Alta exitosa.", vbInformation, strTitulo
    '
    MsgBox "Enseguida Registrara las tareas"
    
    frmAlta.Hide
    tAREAS.Show
End If
'

End Sub
'
'Cerrar formulario
Private Sub CommandButton2_Click()
frmAlta.Hide
frmBuscar.Show
End Sub

