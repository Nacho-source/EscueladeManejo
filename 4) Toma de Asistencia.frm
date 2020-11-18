VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Toma de Asistencia"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7770
   LinkTopic       =   "Form4"
   Picture         =   "4) Toma de Asistencia.frx":0000
   ScaleHeight     =   6330
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCF 
      Caption         =   "¿No sabe el id del Curso-Fecha?"
      Height          =   615
      Left            =   5280
      TabIndex        =   27
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdAlumno 
      Caption         =   "¿No recuerda el id del Alumno?"
      Height          =   615
      Left            =   3720
      TabIndex        =   26
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdInstructor 
      Caption         =   "¿No recuerda el id del Instructor?"
      Height          =   615
      Left            =   2160
      TabIndex        =   25
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "VOLVER"
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Frame frmAlumno 
      Caption         =   "Ingrese datos de llegada:"
      Height          =   3735
      Left            =   3960
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton cmdLimpiarLista 
         Caption         =   "LIMPIAR LISTA"
         Height          =   735
         Left            =   1920
         TabIndex        =   24
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ListBox lstAsistenciaAlumno 
         Height          =   1035
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   3375
      End
      Begin VB.CommandButton cmdRegistrar 
         Caption         =   "REGISTRAR LLEGADA"
         Height          =   735
         Left            =   240
         TabIndex        =   17
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtCF 
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtAlumno 
         Height          =   285
         Left            =   2400
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtHora 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   4
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Ingrese el id del Curso-Fecha:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Ingrese el id del Alumno:"
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Ingrese el horario de llegada:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame frmOpcion 
      Caption         =   "Elija una opcion:"
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3735
      Begin VB.OptionButton optAlumno 
         Caption         =   "Alumno"
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optInstructor 
         Caption         =   "Instructor"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblOpcion 
         Caption         =   "La persona es un:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame frmInstructor 
      Caption         =   "Ingrese datos de llegada:"
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton cmdlLimpiar 
         Caption         =   "LIMPIAR LISTA"
         Height          =   735
         Left            =   1920
         TabIndex        =   23
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtHorario 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   4
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtInstructor 
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtCursoFecha 
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdRegistrarLlegada 
         Caption         =   "REGISTRAR LLEGADA"
         Height          =   735
         Left            =   240
         TabIndex        =   2
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ListBox lstAsistenciaInstructor 
         Height          =   1035
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label lblHorario 
         Caption         =   "Ingrese el horario de llegada:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Ingrese el id del Instructor:"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblCurso 
         Caption         =   "Ingrese el id del Curso-Fecha:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAlumno_Click()
Form9.Show
sql = "select idAlumno, nombre, apellido, dni, email, tel, direccion from alumno a, persona p where p.idPersona=a.idPersona order by idAlumno"
Call Ejecutar_Comando(sql, cn)

Set Form9.dgAlumno.DataSource = rs
End Sub

Private Sub cmdCF_Click()
Form9.Show
sql = "select cf.idCF, c.Ncurso, c.nombre, h.dia, h.horaentrada, h.horasalida from CursoFecha cf, curso c, horario h where c.NCurso=cf.NCurso and cf.NHorario = h.NHorario"
Call Ejecutar_Comando(sql, cn)

Set Form9.dgAlumno.DataSource = rs
End Sub

Private Sub cmdInstructor_Click()
Form9.Show
sql = "select idInstructor, nombre, apellido, dni, email, tel, direccion from instructor i, persona p where p.idPersona=i.idPersona order by idInstructor"
Call Ejecutar_Comando(sql, cn)

Set Form9.dgAlumno.DataSource = rs
End Sub

Private Sub cmdLimpiarLista_Click()
lstAsistenciaAlumno.Clear
txtCF.Text = ""
txtAlumno.Text = ""
txtHora.Text = Time
End Sub

Private Sub cmdlLimpiar_Click()
lstAsistenciaInstructor.Clear
txtInstructor.Text = ""
txtCursoFecha.Text = ""
txtHorario.Text = Time
End Sub

Private Sub cmdRegistrarLlegada_Click()

If (txtHorario.Text <> "" Or txtInstructor.Text <> "" Or txtCursoFecha.Text <> "") Then
    
    sql = "select idPersona from instructor i, CursoFecha cf, CursoInstructor ci where cf.idCF=" & txtCursoFecha.Text & " and i.idInstructor = " & txtInstructor.Text & " and (cf.idCF=ci.idCF) and (ci.idInstructor=i.idInstructor)"
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount = 0) Then
        msj = MsgBox("LOS DATOS NO COINCIDEN. VERIFIQUE SUS INGRESOS", vbCritical, "ERROR")
    Else
        
        sql = "select idPersona from instructor where idInstructor = " & txtInstructor.Text & ""
        Call Ejecutar_Comando(sql, cn)
        
        If (rs.RecordCount <> 0) Then
            rs.MoveFirst
            Do While rs.EOF = False
                Var = Val(rs.Fields(0))
                rs.MoveNext
            Loop
        End If
        
        sql = "select * from RegistroAsistencia where idPersona='" & Var & "' and Fecha = curdate() and idCF='" & txtCursoFecha.Text & "'  "
        Call Ejecutar_Comando(sql, cn)
        
        If (rs.RecordCount <> 0) Then
            Varalf = MsgBox("YA FUE TOMADA ASISTENCIA A ESTA PERSONA", vbInformation)
        Else
            
            
            sql = "select HoraEntrada from horario h, CursoFecha cf, CursoInstructor ci where (ci.idCF=cf.idCF) and (h.NHorario=cf.NHorario) and (idInstructor=" & txtInstructor.Text & ") and (ci.idCF=" & txtCursoFecha.Text & ")"
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
                rs.MoveFirst
                Do While rs.EOF = False
                    Varalf = rs.Fields(0)
                    rs.MoveNext
                Loop
            End If
            
            Dim t0, t1, t2 As Variant
            t0 = Format(txtHorario.Text, "hh:mm")
            t1 = Format(Varalf, "hh:mm")
            t2 = Format(TimeValue(t1) - TimeValue(t0), "hh:mm")
            If ((t2 < "30:00") And (txtHorario.Text < Varalf)) Then
                msj = MsgBox("La persona llegó demasiado temprano para tomarle asistencia")
            Else
                
                lstAsistenciaInstructor.AddItem "Id del Instructor: " & txtInstructor.Text
                lstAsistenciaInstructor.AddItem "Curso-Fecha: " & txtCursoFecha.Text
                lstAsistenciaInstructor.AddItem "Horario de Entrada: " & Varalf
                lstAsistenciaInstructor.AddItem "Horario de Llegada: " & txtHorario.Text
                
                If ((t2 < "15:00") And (txtHorario.Text > Varalf)) Then
                    lstAsistenciaInstructor.AddItem "A horario: No"
                    msj = MsgBox("El instructor llegó fuera de horario. ¿Desea presentar un justificativo?", vbYesNo)
                    If (Val(msj) = 6) Then
                        sql = "select max(CodAsis) from RegistroAsistencia"
                        Call Ejecutar_Comando(sql, cn)
                        
                        If (rs.RecordCount <> 0) Then
                            rs.MoveFirst
                            Do While rs.EOF = False
                                Var = Val(rs.Fields(0))
                                rs.MoveNext
                            Loop
                        End If
                        
                        sql = "insert into justificativo values ('', " & Var & ", curdate())"
                        Call Ejecutar_Comando(sql, cn)
                        
                        lstAsistenciaInstructor.AddItem "Justificado: Si"
                        msj = MsgBox("El justificativo fue registrado con exito")
                    Else
                        lstAsistenciaInstructor.AddItem "Justificado: No"
                        msj = MsgBox("El justificativo no fue registrado")
                    End If
                Else
                    msj = MsgBox("El instructor llegó a tiempo.")
                    lstAsistenciaInstructor.AddItem "A horario: Si"
                End If
                sql = "insert into RegistroAsistencia values ('', " & Var & ", " & txtCursoFecha.Text & ", curdate(), time(curdate()))"
                Call Ejecutar_Comando(sql, cn)
            End If
        End If
    End If
Else
    msj = MsgBox("INGRESE DATOS", vbExclamation)
End If

End Sub

Private Sub cmdVolver_Click()
Form0.Show
Form4.Hide
frmAlumno.Visible = False
frmInstructor.Visible = False
End Sub

Private Sub cmdRegistrar_Click()
If (txtHora.Text <> "" Or txtAlumno.Text <> "" Or txtCF.Text <> "") Then
    
    sql = "select idPersona from alumno a, CursoFecha cf, AlumnoCurso ac where cf.idCF=" & txtCF.Text & " and a.idAlumno= " & txtAlumno.Text & " and (cf.idcf=ac.idcf) and (ac.idAlumno=a.idAlumno)"
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount = 0) Then
        msj = MsgBox("LOS DATOS NO COINCIDEN. VERIFIQUE SUS INGRESOS", vbCritical, "ERROR")
    Else
        
        sql = "select idPersona from alumno where idAlumno= " & txtAlumno.Text & ""
        Call Ejecutar_Comando(sql, cn)
        
        If (rs.RecordCount <> 0) Then
            rs.MoveFirst
            Do While rs.EOF = False
                Var = Val(rs.Fields(0))
                rs.MoveNext
            Loop
        End If
        
        sql = "select * from RegistroAsistencia where idPersona='" & Var & "' and Fecha = curdate() and idCF='" & txtCF.Text & "'  "
        Call Ejecutar_Comando(sql, cn)
        
        If (rs.RecordCount <> 0) Then
            Varalf = MsgBox("YA FUE TOMADA ASISTENCIA A ESTA PERSONA", vbInformation)
        Else
            
            
            sql = "select HoraEntrada from horario h, CursoFecha cf, AlumnoCurso ac where (ac.idCF=cf.idCF) and (h.NHorario=cf.NHorario) and (idAlumno=" & txtAlumno.Text & ") and (cf.idCF=" & txtCF.Text & ")"
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
                rs.MoveFirst
                Do While rs.EOF = False
                    Varalf = rs.Fields(0)
                    rs.MoveNext
                Loop
            End If
            
            Dim t0, t1, t2 As Variant
            t0 = Format(txtHorario.Text, "hh:mm")
            t1 = Format(Varalf, "hh:mm")
            t2 = Format(TimeValue(t1) - TimeValue(t0), "hh:mm")
            If ((t2 < "30:00") And (txtHorario.Text < Varalf)) Then
                msj = MsgBox("La persona llegó demasiado temprano para tomarle asistencia")
            Else
                
                lstAsistenciaAlumno.AddItem "Id del Alumno: " & txtAlumno.Text
                lstAsistenciaAlumno.AddItem "Curso-Fecha: " & txtCF.Text
                lstAsistenciaAlumno.AddItem "Horario de Entrada: " & Varalf
                lstAsistenciaAlumno.AddItem "Horario de Llegada: " & txtHorario.Text
                
                If ((t2 < "15:00") And (txtHorario.Text > Varalf)) Then
                    lstAsistenciaAlumno.AddItem "A horario: No"
                    msj = MsgBox("El Alumno llegó fuera de horario. ¿Desea presentar un justificativo?", vbYesNo)
                    If (Val(msj) = 6) Then
                        sql = "select max(CodAsis) from RegistroAsistencia"
                        Call Ejecutar_Comando(sql, cn)
                        
                        If (rs.RecordCount <> 0) Then
                            rs.MoveFirst
                            Do While rs.EOF = False
                                Var = Val(rs.Fields(0))
                                rs.MoveNext
                            Loop
                        End If
                        
                        sql = "insert into justificativo values ('', " & Var & ", curdate())"
                        Call Ejecutar_Comando(sql, cn)
                        
                        lstAsistenciaAlumno.AddItem "Justificado: Si"
                        msj = MsgBox("El justificativo fue registrado con exito")
                    Else
                        lstAsistenciaAlumno.AddItem "Justificado: No"
                        msj = MsgBox("El justificativo no fue registrado")
                    End If
                Else
                    msj = MsgBox("El alumno llegó a tiempo.")
                    lstAsistenciaAlumno.AddItem "A horario: Si"
                End If
                sql = "insert into RegistroAsistencia values ('', " & Var & ", " & txtCF.Text & ", curdate(), time(curdate()))"
                Call Ejecutar_Comando(sql, cn)
            End If
        End If
    End If
Else
    msj = MsgBox("INGRESE DATOS", vbExclamation)
End If
    
End Sub

Private Sub optAlumno_Click()
If (optAlumno.Value = True) Then
    frmAlumno.Visible = True
    frmInstructor.Visible = False
    txtAlumno.Text = ""
    txtCF.Text = ""
    txtHora.Text = Time
    lstAsistenciaInstructor.Clear
End If
End Sub

Private Sub optInstructor_Click()
If (optInstructor.Value = True) Then
    frmInstructor.Visible = True
    frmAlumno.Visible = False
    optAlumno.Value = False
    txtInstructor.Text = ""
    txtCursoFecha.Text = ""
    txtHorario.Text = Time
    lstAsistenciaAlumno.Clear
End If
End Sub
