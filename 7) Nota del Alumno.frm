VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Gestionar Nota del Alumno"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   LinkTopic       =   "Form7"
   Picture         =   "7) Nota del Alumno.frx":0000
   ScaleHeight     =   4755
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCurso 
      Caption         =   "¿No recuerda el Numero del curso?"
      Height          =   975
      Left            =   4800
      TabIndex        =   14
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdAlumno 
      Caption         =   "¿No recuerda el id del Alumno?"
      Height          =   975
      Left            =   3600
      TabIndex        =   13
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdInstructor 
      Caption         =   "¿No recuerda el id del Instructor?"
      Height          =   975
      Left            =   2400
      TabIndex        =   12
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "VOLVER"
      Height          =   735
      Left            =   240
      Picture         =   "7) Nota del Alumno.frx":55A4
      TabIndex        =   11
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Frame frmIngreso 
      Caption         =   "Ingrese los datos:"
      Height          =   3255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton cmdFiltrar 
         Caption         =   "FILTRAR NUMERO DE CURSO"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2655
      End
      Begin VB.ComboBox cboCurso 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "7) Nota del Alumno.frx":AA94
         Left            =   1560
         List            =   "7) Nota del Alumno.frx":AA96
         TabIndex        =   2
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdRegistrar 
         Caption         =   "REGISTRAR"
         Default         =   -1  'True
         Height          =   615
         Left            =   840
         TabIndex        =   5
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtCalificacion 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtInstructor 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtAlumno 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblCalificacion 
         Caption         =   "Calificación:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label lblCurso 
         Caption         =   "Numero del Curso:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblInstructor 
         Caption         =   "Id del Instructor:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblAlumno 
         Caption         =   "Id del Alumno:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form7"
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

Private Sub cmdCurso_Click()
Form9.Show
sql = "select c.*, h.dia, h.horaentrada, h.horasalida from curso c, horario h, cursofecha cf where (cf.ncurso=c.NCurso) and (cf.nhorario=h.nhorario)"
Call Ejecutar_Comando(sql, cn)

Set Form9.dgAlumno.DataSource = rs

End Sub

Private Sub cmdFiltrar_Click()
If (txtAlumno.Text = "") Then
    msj = MsgBox("INGRESE DATOS", vbExclamation)
    Else
    
    sql = "select idalumno from alumno where idalumno=" & txtAlumno.Text & ""
    Call Ejecutar_Comando(sql, cn)
    
    If rs.RecordCount = 0 Then
        msj = MsgBox("EL ALUMNO NO EXISTE. REVISE SUS DATOS", vbCritical, "ERROR")
    Else
        
        cboCurso.Enabled = True
        cboCurso.Clear
        
        sql = "select c.ncurso from curso c, alumno a, alumnocurso ac, cursofecha cf where c.ncurso=cf.ncurso and cf.idcf=ac.idcf and a.idalumno=ac.idalumno and a.idalumno=" & txtAlumno.Text & ""
        Call Ejecutar_Comando(sql, cn)
        
        If (rs.RecordCount <> 0) Then
            rs.MoveFirst
            Do While rs.EOF = False
                cboCurso.AddItem rs.Fields(0)
                rs.MoveNext
            Loop
        End If
    End If
End If

End Sub

Private Sub cmdInstructor_Click()
Form9.Show
sql = "select idInstructor, nombre, apellido, dni, email, tel, direccion from instructor i, persona p where p.idPersona=i.idPersona order by idInstructor"
Call Ejecutar_Comando(sql, cn)

Set Form9.dgAlumno.DataSource = rs

End Sub

Private Sub cmdRegistrar_Click()
sql = "select * from alumno a, curso c, alumnocurso ac, cursofecha cf where a.idAlumno=ac.idalumno and cf.idcf=ac.idcf and cf.ncurso=c.ncurso"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount = 0) Then
    msj = MsgBox("LOS DATOS NO COINCIDEN. VERIFIQUE LA INFORMACION INGRESADA", vbCritical, "ERROR")
Else
    If ((txtAlumno.Text = "") Or (cboCurso.Text = "") Or (txtInstructor.Text = "") Or (txtCalificacion.Text = "")) Then
        msj = MsgBox("INGRESE DATOS", vbExclamation, "ERROR")
    Else
        msj = MsgBox("¿Está seguro de que desea registrar la evaluación?", vbYesNo)
        If (Val(msj) = 6) Then
            sql = "insert into evaluacion values('', " & txtAlumno.Text & ", " & cboCurso.Text & ", " & txtCalificacion.Text & ", " & txtInstructor.Text & ")"
            Call Ejecutar_Comando(sql, cn)
            
            If (Val(txtCalificacion.Text) >= 6) Then
                MsgBox ("El alumno está aprobado")
                
                sql = "update AlumnoCurso set aprobado=1 where (idAlumno=" & txtAlumno.Text & ")"
                Call Ejecutar_Comando(sql, cn)
            Else
                MsgBox ("El alumno no está aprobado")
            End If
            
            MsgBox ("La evaluación fue registrada exitosamente")
        Else
            MsgBox ("La evaluación no fue registrada")
        End If
        
        txtAlumno.Text = ""
        txtInstructor.Text = ""
        txtCalificacion.Text = ""
        cboCurso.Text = ""
    End If
End If

cboCurso.Enabled = False
cboCurso.Text = ""
txtAlumno.Text = ""
txtInstructor.Text = ""
txtCalificacion.Text = ""

End Sub


Private Sub cmdVolver_Click()
Form7.Hide
Form0.Show
txtAlumno.Text = ""
txtInstructor.Text = ""
txtCalificacion.Text = ""
cboCurso.Text = ""
End Sub
Private Sub Form_Activate()
txtAlumno.SetFocus
End Sub
