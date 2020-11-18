VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Revision del estado del vehiculo"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   LinkTopic       =   "Form6"
   Picture         =   "6) Revisión de Vehículos.frx":0000
   ScaleHeight     =   3285
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "GUARDAR INFORMACION"
      Height          =   615
      Left            =   6960
      TabIndex        =   9
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdVehiculo 
      Caption         =   "¿No recuerda la matrícula del vehículo?"
      Height          =   735
      Left            =   2280
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "VOLVER"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin VB.ListBox lstVehiculo 
      Height          =   1620
      Left            =   3600
      TabIndex        =   6
      Top             =   240
      Width           =   3135
   End
   Begin VB.Frame frmIngreso 
      Caption         =   "Ingrese los datos del Vehículo:"
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3255
      Begin VB.CommandButton cmdIngresar 
         Caption         =   "INGRESAR"
         Default         =   -1  'True
         Height          =   495
         Left            =   840
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtEstado 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtMatricula 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "AAAAA"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblEstado 
         Caption         =   "Estado:(utilizable=1/no utilizable=0)"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblMatricula 
         Caption         =   "Matricula:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGuardar_Click()
msj = MsgBox("¿Está seguro de que desea guardar la información?", vbYesNo)
If (Val(msj) = 6) Then
    
    sql = "select CodVehiculo from Vehiculo where matricula='" & txtMatricula.Text & "'"
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            Var = Val(rs.Fields(0))
            rs.MoveNext
        Loop
    End If
    
    If (Val(txtEstado.Text = 1)) Then
        
        sql = "update vehiculo set utilizable=1 and FRevision=curdate() where Matricula='" & txtMatricula.Text & "'"
        Call Ejecutar_Comando(sql, cn)
    Else
        sql = "update vehiculo set utilizable=0 and FRevision=curdate() where matricula='" & txtMatricula.Text & "'"
        Call Ejecutar_Comando(sql, cn)
    End If
    
    
    sql = "select codvehiculo from vehiculo where matricula='" & txtMatricula.Text & "'"
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            Var = Val(rs.Fields(0))
            rs.MoveNext
        Loop
    End If

    sql = "insert into rotura values ('', curdate(), " & Var & ", " & Var2 & ", 0, " & Var3 & ")"
    Call Ejecutar_Comando(sql, cn)
    
    MsgBox ("Los datos de la revisión fueron guardados con éxito")
Else
    MsgBox ("Los datos no han sido guardados")
End If


End Sub

Private Sub cmdIngresar_Click()
txtMatricula.Enabled = False
txtEstado.Enabled = False

cmdIngresar.Default = False
cmdGuardar.Default = True


If (txtMatricula.Text = Null Or txtEstado.Text = Null) Then
    msj = MsgBox("¡Ingrese datos!", , "ERROR")
Else
    sql = "select CodVehiculo from Vehiculo where matricula='" & txtMatricula.Text & "'"
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            Var = Val(rs.Fields(0))
            rs.MoveNext
        Loop
        
        If (txtEstado.Text <> 1 And txtEstado.Text <> 0) Then
            msj = MsgBox("El dato 'UTILIZABLE' no es correcto", vbExclamation, "ERROR")
        
        Else
                    
            lstVehiculo.Clear
                
            lstVehiculo.AddItem "Codigo del Vehiculo: " & Var
            lstVehiculo.AddItem "Matricula del Vehiculo: " & txtMatricula.Text
            
            sql = "select Marca from Vehiculo where matricula='" & txtMatricula.Text & "'"
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
                rs.MoveFirst
                Do While rs.EOF = False
                    Varalf = rs.Fields(0)
                    rs.MoveNext
                Loop
            End If
            
            lstVehiculo.AddItem "Marca: " & Varalf
            
            sql = "select Modelo from Vehiculo where matricula='" & txtMatricula.Text & "'"
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
                rs.MoveFirst
                Do While rs.EOF = False
                    Varalf = rs.Fields(0)
                    rs.MoveNext
                Loop
            End If
            
            lstVehiculo.AddItem "Modelo: " & Varalf
            
            If (Val(txtEstado.Text) = 1) Then
                lstVehiculo.AddItem "Utilizable: Sí"
            Else
                lstVehiculo.AddItem "Utilizable: No"
                msj = MsgBox("¿La rotura se debe a un daño ocasionado por un alumno?", vbYesNo)
                If (Val(msj) = 6) Then
                    msj = MsgBox("Selecccione un alumno de los de la lista a continuación", vbOKOnly)
                    With Form9
                    
                    .Show
                    sql = "select idAlumno, nombre, apellido, dni, email, tel, direccion from alumno a, persona p where p.idPersona=a.idPersona order by idAlumno"
                    Call Ejecutar_Comando(sql, cn)
                    
                    Set .dgAlumno.DataSource = rs
                    .lblAlumno.Visible = True
                    .lblAlumno2.Visible = True
                    .txtAlumno.Visible = True
                    .cmdSiguiente.Visible = True
                    .cmdCerrar.Visible = False
                    
                    End With
                    
                    
                End If
            End If
               
            lstVehiculo.AddItem "Fecha de Revision: " & Date
        End If
    Else
        msj = MsgBox("El vehículo no está registrado en la Escuela de Manejo Mascher", vbExclamation)
    
End If
End If

End Sub

Private Sub cmdVehiculo_Click()
Form9.Show
sql = "select select * from vehiculo"
Call Ejecutar_Comando(sql, cn)

Set Form9.dgAlumno.DataSource = rs
End Sub

Private Sub cmdVolver_Click()
txtEstado.Text = ""
txtMatricula.Text = ""
txtEstado.Enabled = True
txtMatricula.Enabled = True
Form0.Show
Form6.Hide
cmdVolver.Visible = False
lstVehiculo.Clear
End Sub

Private Sub Form_Activate()
If (flag = 0) Then
    txtMatricula.SetFocus
    flag = 1
End If
End Sub
