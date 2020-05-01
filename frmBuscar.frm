VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBuscar 
   Caption         =   "EXCELeINFO - Modificar registros"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13995
   OleObjectBlob   =   "frmBuscar.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, items, xProyecto

'Cerrar formulario
Private Sub CommandButton2_Click()
frmBuscar.Hide
frmPrincipal.Show
End Sub
'

'Abrir el formulario para modificar
Private Sub CommandButton3_Click()
If Me.ListBox1.ListIndex < 0 Then
    MsgBox "No se ha elegido ningún registro", vbExclamation, "EXCELeINFO"
Else
frmModificar.Show
End If
End Sub
'
'Eliminar el registro
Private Sub CommandButton4_Click()
Pregunta = MsgBox("Está seguro de eliminar el registro?", vbYesNo + vbQuestion, "EXCELeINFO")
If Pregunta <> vbNo Then
    ActiveCell.EntireRow.Delete
End If
Call CommandButton5_Click
End Sub
'
'Mostrar resultado en ListBox
Private Sub CommandButton5_Click()
        If Me.txtFiltro1.Value = Empty Then
            MsgBox "Escriba un registro para buscar"
            Me.ListBox1.Clear
            Me.txtFiltro1.SetFocus
            Exit Sub
        End If

Me.ListBox1.Clear

items = Range("Tabla1").CurrentRegion.Rows.Count
        For i = 1 To 8
            If LCase(Cells(i, 1).Value) Like "*" & LCase(Me.txtFiltro1.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Cells(i, 7)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = Cells(i, 8)
            
                ElseIf LCase(Cells(i, 2).Value) Like "*" & LCase(Me.txtFiltro1.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Cells(i, 7)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = Cells(i, 8)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 8) = Cells(i, 9)
                
                 ElseIf LCase(Cells(i, 3).Value) Like "*" & LCase(Me.txtFiltro1.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Cells(i, 7)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = Cells(i, 8)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 8) = Cells(i, 9)
                
                 ElseIf LCase(Cells(i, 4).Value) Like "*" & LCase(Me.txtFiltro1.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Cells(i, 7)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = Cells(i, 8)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 8) = Cells(i, 9)
                
                 ElseIf LCase(Cells(i, 5).Value) Like "*" & LCase(Me.txtFiltro1.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Cells(i, 7)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = Cells(i, 8)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 8) = Cells(i, 9)
                
                 ElseIf LCase(Cells(i, 6).Value) Like "*" & LCase(Me.txtFiltro1.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Cells(i, 7)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = Cells(i, 8)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 8) = Cells(i, 9)
                
                 ElseIf LCase(Cells(i, 7).Value) Like "*" & LCase(Me.txtFiltro1.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Cells(i, 7)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = Cells(i, 8)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 8) = Cells(i, 9)
                
                 ElseIf LCase(Cells(i, 8).Value) Like "*" & LCase(Me.txtFiltro1.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Cells(i, 7)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = Cells(i, 8)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 8) = Cells(i, 9)
            End If
        Next i
        Me.txtFiltro1.SetFocus
        Me.txtFiltro1.SelStart = 0
        Me.txtFiltro1.SelLength = Len(Me.txtFiltro1.Text)
Exit Sub


End Sub

Private Sub CommandButton6_Click()
Dim x1 As String
x1 = ComboBox1.Value
        If Me.txt_Buscar2.Value = Empty Then
            MsgBox "Escriba un registro para buscar"
            Me.ListBox1.Clear
            Me.txt_Buscar2.SetFocus
            Exit Sub
        End If

Me.ListBox1.Clear

items = Range("Tabla2").CurrentRegion.Rows.Count
        For i = 1 To 9
        If ComboBox1.Text = "ID" Then
            If LCase(Cells(i, 1).Value) Like "*" & LCase(Me.txt_Buscar2.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Cells(i, 7)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = Cells(i, 8)
            End If
            End If
        
            If ComboBox1.Text = "Proyecto" Then
            If LCase(Cells(i, 2).Value) Like "*" & LCase(Me.txt_Buscar2.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Cells(i, 7)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = Cells(i, 8)
            End If
            End If
            
            If ComboBox1.Text = "Responsable1" Then
            If LCase(Cells(i, 3).Value) Like "*" & LCase(Me.txt_Buscar2.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Cells(i, 7)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = Cells(i, 8)
            End If
            End If
            
            
            If ComboBox1.Text = "Responsable2" Then
            If LCase(Cells(i, 4).Value) Like "*" & LCase(Me.txt_Buscar2.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Cells(i, 7)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = Cells(i, 8)
            End If
            End If
            
            
            If ComboBox1.Text = "Fecha-Inicio" Then
            If LCase(Cells(i, 5).Value) Like "*" & LCase(Me.txt_Buscar2.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Cells(i, 7)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = Cells(i, 8)
            End If
            End If
            
            
            If ComboBox1.Text = "Fecha-Final" Then
            If LCase(Cells(i, 6).Value) Like "*" & LCase(Me.txt_Buscar2.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Cells(i, 7)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = Cells(i, 8)
            End If
            End If
            
            
            If ComboBox1.Text = "Ingreso" Then
            If LCase(Cells(i, 7).Value) Like "*" & LCase(Me.txt_Buscar2.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Cells(i, 7)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = Cells(i, 8)
            End If
            End If
            
            If ComboBox1.Text = "Tareas" Then
            If LCase(Cells(i, 8).Value) Like "*" & LCase(Me.txt_Buscar2.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Cells(i, 7)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = Cells(i, 8)
            End If
            End If
            
            If ComboBox1.Text = "Avance" Then
            If LCase(Cells(i, 9).Value) Like "*" & LCase(Me.txt_Buscar2.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Cells(i, 7)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = Cells(i, 8)
            End If
            End If
        Next i
        Me.txt_Buscar2.SetFocus
        Me.txt_Buscar2.SelStart = 0
        Me.txt_Buscar2.SelLength = Len(Me.txt_Buscar2.Text)
Exit Sub



End Sub

Private Sub CommandButton7_Click()
If Me.ListBox1.ListIndex < 0 Then
MsgBox "No se ha elegido ningún registro", vbExclamation, "EXCELeINFO"
Else
Seguimiento.Show
End If
End Sub

Private Sub CommandButton8_Click()
frmBuscar.Hide
frmAlta.Show
End Sub

Private Sub Image1_Click()

End Sub

'Activar la celda del registro elegido
Private Sub ListBox1_Click()
Range("a2").Activate
Cuenta = Me.ListBox1.ListCount
Set Rango = Range("A1").CurrentRegion
For i = 0 To Cuenta - 1
    If Me.ListBox1.Selected(i) Then
        Valor = Me.ListBox1.List(i)
        Rango.Find(What:=Valor, LookAt:=xlWhole, After:=ActiveCell).Activate
    End If
Next i
End Sub
'
'Dar formato al ListBox y traer datos de la tabla
Private Sub UserForm_Initialize()
For i = 1 To 8
    Me.Controls("Label" & i) = Cells(1, i).Value
Next i

With ListBox1
    .ColumnCount = 8

End With
End Sub
