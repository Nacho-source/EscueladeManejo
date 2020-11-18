VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H80000009&
   Caption         =   "Pago de Rotura"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   LinkTopic       =   "Form10"
   ScaleHeight     =   3195
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVehiculo 
      Caption         =   "¿No recuerda el código del vehículo?"
      Height          =   615
      Left            =   3240
      TabIndex        =   10
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdMonto 
      Caption         =   "¿No recuerda costo de la rotura?"
      Height          =   615
      Left            =   3240
      TabIndex        =   9
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdAlumno 
      Caption         =   "¿No recuerda el id del Alumno?"
      Height          =   615
      Left            =   3240
      TabIndex        =   8
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "VOLVER"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "REGISTRAR PAGO"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtMonto 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtVehiculo 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtAlumno 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblMonto 
      BackColor       =   &H80000009&
      Caption         =   "Monto del Pago:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Codigo del Vehiculo:"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblAlumno 
      BackColor       =   &H80000009&
      Caption         =   "id del Alumno:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAlumno_Click()
Form9.Show
sql = "select idInstructor, nombre, apellido, dni, email, tel, direccion from instructor i, persona p where p.idPersona=i.idPersona order by idInstructor"
Call Ejecutar_Comando(sql, cn)

Set Form9.dgAlumno.DataSource = rs

End Sub

Private Sub cmdMonto_Click()
Form9.Show
sql = "select * from rotura"
Call Ejecutar_Comando(sql, cn)

Set Form9.dgAlumno.DataSource = rs

End Sub

Private Sub cmdRegistrar_Click()

If (txtAlumno.Text <> "" Or txtVehiculo.Text <> "" Or txtMonto.Text <> "") Then
    
    sql = "select idAlumno from rotura where (idAlumno=" & txtAlumno.Text & ")"
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount = 0) Then
        msj = MsgBox("LOS DATOS NO COINCIDEN. VERIFIQUE SUS INGRESOS", vbCritical, "ERROR")
    Else
        
        sql = "select idAlumno from rotura r, PagoRotura pr Where (R.CodRotura = pr.CodRotura) and (idAlumno=" & txtAlumno.Text & ") and pagada=1"
        Call Ejecutar_Comando(sql, cn)
        
        If (rs.RecordCount <> 0) Then
            msj = MsgBox("La rotura ya fue pagada")
            Else
            
            sql = "select monto from rotura where idAlumno=" & txtAlumno.Text & " and CodVehiculo= " & txtVehiculo.Text & ""
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
                rs.MoveFirst
                Do While rs.EOF = False
                    Var = Val(rs.Fields(0))
                    rs.MoveNext
                Loop
            End If
            
            If (Var > Val(txtMonto.Text)) Then
                msj = MsgBox("El monto es insuficiente para cubrir las roturas. El pago no será registrado", vbExclamation)
            Else
                sql = "select codRotura from rotura where idAlumno=" & txtAlumno.Text & " and CodVehiculo= " & txtVehiculo.Text & ""
                Call Ejecutar_Comando(sql, cn)
                
                If (rs.RecordCount <> 0) Then
                    rs.MoveFirst
                    Do While rs.EOF = False
                        Var = Val(rs.Fields(0))
                        rs.MoveNext
                    Loop
                End If
                
                sql = "insert into PagoRotura values(" & Var & ", curdate())"
                Call Ejecutar_Comando(sql, cn)
                
                sql = "update rotura set pagada=1 where CodRotura=" & Var & ""
                Call Ejecutar_Comando(sql, cn)
                
                msj = MsgBox("El pago fue ejecutado exitosamente")
                
                With formComprobantePagoRotura
                
                .Show
                
                sql = "select nombre, apellido from alumno a, persona p where p.idpersona=a.idpersona and idAlumno='" & txtAlumno.Text & "'"
                Call Ejecutar_Comando(sql, cn)
                
                If (rs.RecordCount <> 0) Then
                    rs.MoveFirst
                    Do While rs.EOF = False
                        Varalf = rs.Fields(0) & " " & rs.Fields(1)
                        rs.MoveNext
                    Loop
                End If
                
                .lblAlumno = Varalf
                
                sql = "select matricula from vehiculo v, rotura r where (r.CodVehiculo=v.CodVehiculo) and v.CodVehiculo=" & txtVehiculo.Text & ""
                Call Ejecutar_Comando(sql, cn)
                
                If (rs.RecordCount <> 0) Then
                    rs.MoveFirst
                    Do While rs.EOF = False
                        Varalf = rs.Fields(0)
                        rs.MoveNext
                    Loop
                End If
                
                .lblMatricula = Varalf
                .lblFecha = Date
                End With
                
                Form10.Hide
                
            End If
        End If
    End If
Else
    msj = MsgBox("INGRESE DATOS", vbExclamation)
End If

txtAlumno.Text = ""
txtVehiculo.Text = ""
txtMonto.Text = ""

End Sub

Private Sub cmdVehiculo_Click()
Form9.Show
sql = "select vehiculo.* from vehiculo"
Call Ejecutar_Comando(sql, cn)

Set Form9.dgAlumno.DataSource = rs

End Sub

Private Sub cmdVolver_Click()

txtAlumno.Text = ""
txtVehiculo.Text = ""
txtMonto.Text = ""

Form10.Hide
Form0.Show
End Sub
