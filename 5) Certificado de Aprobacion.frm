VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Gestión del Certificado de Aprobacion del Curso"
   ClientHeight    =   2385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7755
   LinkTopic       =   "Form3"
   Picture         =   "5) Certificado de Aprobacion.frx":0000
   ScaleHeight     =   2385
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCurso 
      Caption         =   "¿No recuerda el Numero del curso?"
      Height          =   615
      Left            =   5160
      TabIndex        =   8
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdAlumno 
      Caption         =   "¿No recuerda el id del Alumno?"
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtCurso 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "VOLVER"
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtidAlumno 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdVerificar 
      Caption         =   "VERIFICAR"
      Default         =   -1  'True
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblCurso 
      BackColor       =   &H80000005&
      Caption         =   "Ingrese numero del Curso:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblIng 
      BackColor       =   &H8000000E&
      Caption         =   "Ingrese id del Alumno:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblVer 
      BackColor       =   &H8000000E&
      Caption         =   "Verificar Aprobación del Curso:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form5"
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

Private Sub cmdVerificar_Click()
If (txtidAlumno.Text = "" Or txtCurso.Text = "") Then
    msj = MsgBox("INGRESE DATOS", vbExclamation, "ERROR")
Else
    
    sql = "select * from alumno a, curso c, alumnocurso ac, cursofecha cf where (a.idAlumno=ac.idAlumno) and (c.NCurso=cf.NCurso) and (a.idAlumno='" & txtidAlumno.Text & "') and (c.NCurso='" & txtCurso.Text & "')"
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount = 0) Then
        msj = MsgBox("LOS DATOS NO COINCIDEN. REVISAR INGRESOS", vbCritical, "ERROR")
    Else
        
        Form5.Hide
        txtidAlumno.Enabled = False
        txtCurso.Enabled = False
        
        sql = "select aprobado from AlumnoCurso ac, CursoFecha cf, curso c Where cf.idCF = ac.idCF and c.NCurso = cf.NCurso and idAlumno=" & txtidAlumno.Text & " and c.NCurso= " & txtCurso.Text & ""
        Call Ejecutar_Comando(sql, cn)
        
        If (rs.RecordCount <> 0) Then
            rs.MoveFirst
            Do While rs.EOF = False
                Var = Val(rs.Fields(0))
                rs.MoveNext
            Loop
        End If
        
        If (Var = 1) Then
            MsgBox ("El curso está aprobado")
            With formCertAp
            .Show
            
            sql = "select apellido from persona p, alumno a where (idAlumno= " & txtidAlumno.Text & ") and (p.idPersona=a.idPersona)"
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
                rs.MoveFirst
                Do While rs.EOF = False
                    Varalf = rs.Fields(0)
                    rs.MoveNext
                Loop
            End If
            
            .lblApellido.Caption = Varalf
            
            sql = "select nombre from persona p, alumno a where (idAlumno= " & txtidAlumno.Text & ") and (p.idPersona=a.idPersona)"
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
               rs.MoveFirst
               Do While rs.EOF = False
                   Varalf = rs.Fields(0)
                   rs.MoveNext
               Loop
            End If
            
            .lblNombre.Caption = Varalf
            
            sql = "select dni from persona p, alumno a where (idAlumno= " & txtidAlumno.Text & ") and (p.idPersona=a.idPersona)"
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
               rs.MoveFirst
               Do While rs.EOF = False
                   Varalf = rs.Fields(0)
                   rs.MoveNext
               Loop
            End If
            
            .lblDNI.Caption = Varalf
            
            sql = "select nombre from curso where NCurso=" & txtCurso.Text & ""
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
               rs.MoveFirst
               Do While rs.EOF = False
                   Varalf = rs.Fields(0)
                   rs.MoveNext
               Loop
            End If
            
            .lblCurso.Caption = Varalf
            
            sql = "select curdate()"
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
               rs.MoveFirst
               Do While rs.EOF = False
                   Varalf = rs.Fields(0)
                   rs.MoveNext
               Loop
            End If
            
            .lblFecha.Caption = Varalf
            
            sql = "select nota from evaluacion e, curso c where (idAlumno=" & Val(txtidAlumno.Text) & ") and (c.NCurso='" & txtCurso.Text & "') and (c.NCurso=e.NCurso)"
            Call Ejecutar_Comando(sql, cn)
        
            If (rs.RecordCount <> 0) Then
                rs.MoveFirst
                Do While rs.EOF = False
                    varfloat = Val(rs.Fields(0))
                    rs.MoveNext
                Loop
            End If
            
            .lblCalificacion = varfloat
            
            End With
            
        Else
            MsgBox ("El curso no está aprobado")
            With formInformeDesapruebo
            
            .Show
            
            sql = "select apellido from persona p, alumno a where (idAlumno= " & txtidAlumno.Text & ") and (p.idPersona=a.idPersona)"
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
                rs.MoveFirst
                Do While rs.EOF = False
                    Varalf = rs.Fields(0)
                    rs.MoveNext
                Loop
            End If
            
            .lblApellido.Caption = Varalf
            
            sql = "select nombre from persona p, alumno a where (idAlumno= " & txtidAlumno.Text & ") and (p.idPersona=a.idPersona)"
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
               rs.MoveFirst
               Do While rs.EOF = False
                   Varalf = rs.Fields(0)
                   rs.MoveNext
               Loop
            End If
            
            .lblNombre.Caption = Varalf
            
            sql = "select dni from persona p, alumno a where (idAlumno= " & txtidAlumno.Text & ") and (p.idPersona=a.idPersona)"
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
               rs.MoveFirst
               Do While rs.EOF = False
                   Varalf = rs.Fields(0)
                   rs.MoveNext
               Loop
            End If
            
            .lblDNI.Caption = Varalf
            
            sql = "select nombre from curso where NCurso=" & txtCurso.Text & ""
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
               rs.MoveFirst
               Do While rs.EOF = False
                   Varalf = rs.Fields(0)
                   rs.MoveNext
               Loop
            End If
            
            .lblCurso.Caption = Varalf
            
            sql = "select curdate()"
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
               rs.MoveFirst
               Do While rs.EOF = False
                   Varalf = rs.Fields(0)
                   rs.MoveNext
               Loop
            End If
            
            .lblFecha.Caption = Varalf
            
            sql = "select nota from evaluacion e, curso c where (idAlumno=" & Val(txtidAlumno.Text) & ") and (c.NCurso='" & txtCurso.Text & "') and (c.NCurso=e.NCurso)"
            Call Ejecutar_Comando(sql, cn)
        
            If (rs.RecordCount <> 0) Then
                rs.MoveFirst
                Do While rs.EOF = False
                    varfloat = Val(rs.Fields(0))
                    rs.MoveNext
                Loop
            End If
            
            .lblCalificacion = varfloat
            
            End With
        End If
    End If
End If
End Sub

Private Sub cmdVolver_Click()
Form0.Show
Form5.Hide

txtCurso.Enabled = True
txtidAlumno.Enabled = True
txtCurso.Text = ""
txtidAlumno.Text = ""

End Sub

Private Sub Form_Activate()
txtidAlumno.SetFocus
End Sub
