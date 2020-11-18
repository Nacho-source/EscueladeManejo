VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFF00&
   Caption         =   "Inscripcion"
   ClientHeight    =   6795
   ClientLeft      =   1845
   ClientTop       =   2550
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   Picture         =   "1) Inscripcion.frx":0000
   ScaleHeight     =   6795
   ScaleWidth      =   9960
   Begin VB.Frame frmQuitarEspera 
      BackColor       =   &H8000000E&
      Height          =   1575
      Left            =   6000
      TabIndex        =   27
      Top             =   2400
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton cmdQuitarEspera 
         Caption         =   "QUITAR DE LA LISTA DE ESPERA"
         Height          =   855
         Left            =   480
         TabIndex        =   28
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame frmGenerarComprobanteIns 
      BackColor       =   &H8000000E&
      Height          =   1575
      Left            =   4440
      TabIndex        =   23
      Top             =   3960
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CommandButton cmdCompIns 
         Caption         =   "GENERAR COMPROBANTE DE INSCRIPCION"
         Height          =   1095
         Left            =   360
         TabIndex        =   34
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame frmGenerarComprobante 
      BackColor       =   &H8000000E&
      Height          =   1575
      Left            =   4440
      TabIndex        =   22
      Top             =   3960
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton cmbComprobanteEspera 
         Caption         =   "GENERAR COMPROBANTE DE ESPERA"
         Height          =   855
         Left            =   360
         TabIndex        =   26
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame frmVerificarAlumno 
      BackColor       =   &H8000000E&
      Height          =   1575
      Left            =   4440
      TabIndex        =   21
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
      Begin VB.CommandButton cmdVerifAlumno 
         Caption         =   "VERIFICAR ALUMNO"
         Height          =   855
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame frmAsignarCurso 
      Caption         =   "Asignar Curso"
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   4440
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CommandButton cmdSig 
         Caption         =   "Siguiente"
         Height          =   615
         Left            =   2400
         TabIndex        =   33
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "Siguiente"
         Height          =   615
         Left            =   3120
         TabIndex        =   32
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ListBox lstListaHorario 
         Height          =   1425
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   2775
      End
      Begin VB.ComboBox cboPractico 
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Text            =   "Seleccione Curso-Fecha"
         Top             =   720
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox cboTeorico 
         Height          =   315
         ItemData        =   "1) Inscripcion.frx":54DC
         Left            =   120
         List            =   "1) Inscripcion.frx":54DE
         TabIndex        =   29
         Text            =   "Seleccione Division"
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblCurso 
         Caption         =   "Seleccion código del curso:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
      End
   End
   Begin VB.Frame frmAsignarEspera 
      BackColor       =   &H8000000E&
      Caption         =   " Lista de Espera"
      Height          =   1575
      Left            =   1560
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton cmdAsignarEspera 
         Caption         =   "ASIGNAR A LISTA DE ESPERA"
         Height          =   855
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame frmEspera 
      BackColor       =   &H8000000E&
      Caption         =   "Ver lista de Espera"
      Height          =   1575
      Left            =   4440
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton cmdVerifListaEspera 
         BackColor       =   &H000000FF&
         Caption         =   "VERIFICAR LISTA DE ESPERA"
         Height          =   855
         Left            =   480
         TabIndex        =   19
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame frmIngreso 
      Caption         =   "Ingrese Datos Personales"
      Height          =   4935
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   4335
      Begin VB.ComboBox cboCurso 
         Height          =   315
         ItemData        =   "1) Inscripcion.frx":54E0
         Left            =   1440
         List            =   "1) Inscripcion.frx":54F3
         TabIndex        =   6
         Top             =   3360
         Width           =   2415
      End
      Begin VB.CommandButton cmdVerificar 
         BackColor       =   &H000000FF&
         Caption         =   "VERIFICAR DISPONIBILIDAD DEL TIPO DE CURSO"
         Default         =   -1  'True
         Height          =   735
         Left            =   1440
         TabIndex        =   7
         Top             =   3840
         Width           =   2415
      End
      Begin VB.TextBox txtNombre 
         Height          =   375
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtApellido 
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox txtDireccion 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtTelefono 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox txtDNI 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblApellido 
         Caption         =   "Apellido:"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblDireccion 
         Caption         =   "Direccion:"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblTel 
         Caption         =   "Telefono:"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email:"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lblDNI 
         Caption         =   "DNI:"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label lblTipo 
         Caption         =   "Nombre de Curso:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "VOLVER"
      Height          =   735
      Left            =   120
      Picture         =   "1) Inscripcion.frx":5544
      TabIndex        =   8
      Top             =   5640
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboPractico_Click()
cmdSiguiente.Visible = True

sql = "select h.dia from CursoFecha cf, curso c, horario h Where cf.NCurso = C.NCurso And h.NHorario = cf.NHorario and cf.idCF=" & cboPractico.Text & ""
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Varalf = rs.Fields(0)
        rs.MoveNext
    Loop
End If

sql = "select h.HoraEntrada from CursoFecha cf, curso c, horario h Where cf.NCurso = C.NCurso And h.NHorario = cf.NHorario and cf.idCF=" & cboPractico.Text & ""
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Varalf2 = rs.Fields(0)
        rs.MoveNext
    Loop
End If

sql = "select h.HoraSalida from CursoFecha cf, curso c, horario h Where cf.NCurso = C.NCurso And h.NHorario = cf.NHorario and cf.idCF=" & cboPractico.Text & ""
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Varalf3 = rs.Fields(0)
        rs.MoveNext
    Loop
End If

lstListaHorario.Clear
lstListaHorario.AddItem ("Dia: " & Varalf)
lstListaHorario.AddItem ("Horario de Entrada: " & Varalf2)
lstListaHorario.AddItem ("Horario de Salida: " & Varalf3)

End Sub

Private Sub cboTeorico_Click()
cmdSiguiente.Visible = True

sql = "select h.dia from CursoFecha cf, curso c, teorico t, horario h Where (cf.NCurso = C.NCurso) and (t.NCurso=c.NCurso) and (h.NHorario=cf.NHorario) and t.NTeorico=" & cboTeorico.Text & ""
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Varalf = rs.Fields(0)
        rs.MoveNext
    Loop
End If

sql = "select h.HoraEntrada from CursoFecha cf, curso c, teorico t, horario h Where cf.NCurso = C.NCurso and t.NCurso=c.NCurso and h.NHorario=cf.NHorario and t.NTeorico=" & cboTeorico.Text & ""
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Varalf2 = rs.Fields(0)
        rs.MoveNext
    Loop
End If

sql = "select h.HoraSalida from CursoFecha cf, curso c, teorico t, horario h Where cf.NCurso = C.NCurso and t.NCurso=c.NCurso and h.NHorario=cf.NHorario and t.NTeorico=" & cboTeorico.Text & ""
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Varalf3 = rs.Fields(0)
        rs.MoveNext
    Loop
End If

lstListaHorario.Clear
lstListaHorario.AddItem ("Dia: " & Varalf)
lstListaHorario.AddItem ("Horario de Entrada: " & Varalf2)
lstListaHorario.AddItem ("Horario de Salida: " & Varalf3)

End Sub

Private Sub cmbComprobanteEspera_Click()
Form1.Visible = False
With formComprobanteEspera
.Visible = True

.lblNombre.Caption = Form1.txtNombre.Text
.lblApellido.Caption = Form1.txtApellido.Text
.lblDNI = Form1.txtDNI.Text
.lblEmail = Form1.txtEmail.Text
.lblNomCurso = Form1.cboCurso.Text
.lblTel = Form1.txtTelefono.Text
.lblDireccion = Form1.txtDireccion.Text

sql = "select max(NEspera) from ListaEspera"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

.lblEspera.Caption = Var

sql = "select NCurso from curso where nombre='" & cboCurso.Text & "'"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var2 = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

.lblNCurso.Caption = Var2

.lblFecha.Caption = Date

End With
End Sub

Private Sub cmdAsignarEspera_Click()

frmAsignarEspera.Visible = False

sql = "select idPersona from persona where dni='" & txtDNI.Text & "'"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount = 0) Then
    sql = "insert into persona values (' ', '" & txtNombre.Text & "', '" & txtApellido.Text & "', '" & txtDireccion.Text & "', '" & txtTelefono.Text & "', '" & txtEmail.Text & "', '" & txtDNI.Text & "')"
    Call Ejecutar_Comando(sql, cn)
End If

sql = "select MAX(idPersona) from persona"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

sql = "select NCurso from curso where nombre='" & cboCurso.Text & "' limit 1"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var2 = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

msj = MsgBox("¿Esta seguro de que desea agregar a la persona al a lista de espera?", vbYesNo, "Advertencia")
If (Val(msj) = 6) Then
    sql = "insert into ListaEspera values (' '," & Var & ", " & Var2 & ", curdate())"
    Call Ejecutar_Comando(sql, cn)
    
    frmGenerarComprobante.Visible = True
    MsgBox ("La persona ha sido asignada a la Lista de Espera exitosamente")
Else
    MsgBox ("La persona no ha sido asignada a la Lisa de Espera")
    If (cboCurso.Text = "Teorico") Then
        frmTeorico.Visible = True
    Else
        frmPractico.Visible = True
    End If
End If

End Sub

Private Sub cmdSig_Click()
lblCurso.Visible = True
cmdSig.Visible = False

If (cboCurso.Text = "Teorico") Then
    flag = 1
    
    cboTeorico.Visible = True
    
    sql = "select distinct t.NTeorico from teorico t, AlumnoCurso ac, (select count(*) cant, t.NTeorico from AlumnoCurso ac, CursoFecha cf, teorico t Where ac.idcf = cf.idcf and cf.ncurso=t.ncurso group by t.NTeorico) sc Where t.NTeorico = sc.NTeorico and t.cupo>sc.cant limit 1"
    Call Ejecutar_Comando(sql, cn)
    
    cboTeorico.Clear
    
    If (rs.RecordCount > 0) Then
        rs.MoveFirst
        Do While (rs.EOF = False)
            cboTeorico.AddItem (Val(rs.Fields(0)))
            rs.MoveNext
        Loop
    End If
    
Else
    flag = 0
    
    cboPractico.Visible = True
    
    sql = "select distinct cf.idCF from CursoFecha cf, Curso c, CursoInstructor ci Where (nombre = '" & cboCurso.Text & "') And (cf.ncurso = C.ncurso)"
    Call Ejecutar_Comando(sql, cn)
    
    cboPractico.Clear
    
    If (rs.RecordCount > 0) Then
        rs.MoveFirst
        Do While (rs.EOF = False)
            cboPractico.AddItem Val((rs.Fields(0)))
            rs.MoveNext
        Loop
    End If
    
End If

End Sub

Private Sub cmdCompIns_Click()
Form1.Visible = False

With formCompIns
.Visible = True
.lblNombre.Caption = Form1.txtNombre
.lblApellido.Caption = Form1.txtApellido
.lblDireccion.Caption = Form1.txtDireccion
.lblDNI.Caption = Form1.txtDNI
.lblEmail.Caption = Form1.txtEmail
.lblFecha.Caption = Date
.lblTel.Caption = Form1.txtTelefono

If (flagteorico = 1) Then
    sql = "select h.HoraEntrada from CursoFecha cf, curso c, teorico t, horario h Where cf.NCurso = C.NCurso and t.NCurso=c.NCurso and h.NHorario=cf.NHorario and t.NTeorico=" & cboTeorico.Text & ""
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            Varalf2 = rs.Fields(0)
            rs.MoveNext
        Loop
    End If
    
    .lblHEntrada.Caption = Varalf2
    
    sql = "select h.HoraSalida from CursoFecha cf, curso c, teorico t, horario h Where cf.NCurso = C.NCurso and t.NCurso=c.NCurso and h.NHorario=cf.NHorario and t.NTeorico=" & cboTeorico.Text & ""
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            Varalf3 = rs.Fields(0)
            rs.MoveNext
        Loop
    End If

    .lblHSalida.Caption = Varalf3
    
Else
    sql = "select h.HoraEntrada from CursoFecha cf, curso c, horario h Where cf.NCurso = C.NCurso And h.NHorario = cf.NHorario and cf.idCF=" & cboPractico.Text & ""
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            Varalf2 = rs.Fields(0)
            rs.MoveNext
        Loop
    End If
    
    .lblHEntrada.Caption = Varalf2
    
    sql = "select h.HoraSalida from CursoFecha cf, curso c, horario h Where cf.NCurso = C.NCurso And h.NHorario = cf.NHorario and cf.idCF=" & cboPractico.Text & ""
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            Varalf3 = rs.Fields(0)
            rs.MoveNext
        Loop
    End If
    
    .lblHSalida.Caption = Varalf3
End If
    
sql = "select idAlumno from persona p, alumno a where dni= ('" & Form1.txtDNI.Text & "') and (a.idPersona=p.idPersona)"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

.lblAlumno.Caption = Var
.lblNComprobante.Caption = Var

sql = "select NCurso from AlumnoCurso ac, CursoFecha cf where FIngreso =(select max(FIngreso) from AlumnoCurso) and idAlumno = " & Var & " and ac.idCF=cf.idCF"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

.lblCurso.Caption = Var

sql = "select nombre from curso where NCurso=" & Var & ""
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Varalf = rs.Fields(0)
        rs.MoveNext
    Loop
End If

.lblNCurso.Caption = Varalf

sql = "select CodVehiculo from AlumnoVehiculo where idAlumno= " & .lblAlumno.Caption & ""
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

.lblVehiculo.Caption = Var

sql = "select matricula from vehiculo where codVehiculo= " & Var & ""
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Varalf = rs.Fields(0)
        rs.MoveNext
    Loop
End If

.lblMatricula.Caption = Varalf

sql = "select idInstructor from alumnocurso ac, cursoinstructor ci Where ac.idcf = ci.idcf and idalumno=" & Var2 & ""
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

.lblInstructor.Caption = Var

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
End With
End Sub

Private Sub cmdQuitarEspera_Click()
frmQuitarEspera.Visible = False

sql = "select idPersona From persona where DNI = '" & txtDNI.Text & "'"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

msj = MsgBox("¿Está seguro de que desea borrar el dato de la lista de espera?", vbYesNo, "Advertencia")
If (Val(msj) = 6) Then
    sql = "delete from ListaEspera where idPersona=" & Var & ""
    Call Ejecutar_Comando(sql, cn)
    
    MsgBox ("El alumno ha sido borrado de la Lista de Espera exitosamente")
    frmAsignarCurso.Visible = True
Else
    MsgBox ("El alumno no ha sido borrado de la Lista de Espera")
    frmEspera.Visible = True
End If


End Sub

Private Sub cmdRegistrarAlumno_Click()

frmRegistrarAlumno.Visible = False

sql = "select idPersona from persona where dni='" & txtDNI.Text & "'"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount = 0) Then
    sql = "insert into persona values ('', '" & txtNombre.Text & "', '" & txtApellido.Text & "', '" & txtDireccion.Text & "', '" & txtTelefono.Text & "', '" & txtEmail.Text & "', '" & txtDNI.Text & "')"
    Call Ejecutar_Comando(sql, cn)
   
End If

sql = "select MAX(idPersona) from persona"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

sql = "insert into alumno values ('', " & Var & ")"
Call Ejecutar_Comando(sql, cn)

MsgBox ("Usted ha sido registrado exitosamente.")

frmAsignarCurso.Visible = True

End Sub

Private Sub cmdSiguiente_Click()
cmdSig.Visible = False
cont = cont + 1
If (flag = 0) Then
    If (cont = 1) Then
        formVehiculos.Visible = True
        Form1.Visible = False
    Else
    
        sql = "select max(idAlumno) from alumno"
        Call Ejecutar_Comando(sql, cn)
        
        If (rs.RecordCount <> 0) Then
            rs.MoveFirst
            Do While rs.EOF = False
                Var = Val(rs.Fields(0))
                rs.MoveNext
            Loop
        End If
    
        sql = "insert into AlumnoCurso values (" & Var & ",'" & cboPractico.Text & "', curdate(), 0)"
        Call Ejecutar_Comando(sql, cn)
        
        sql = "insert into AlumnoVehiculo values (" & Var & ", '" & cboPractico.Text & "', '" & formVehiculos.txtCodVehiculo.Text & "')"
        Call Ejecutar_Comando(sql, cn)
        
        frmAsignarCurso.Visible = False
        frmGenerarComprobanteIns.Visible = True
        
    End If
    
Else
    If (cont = 1) Then
        
        sql = "select p.nombre from persona p, instructor i, curso c, CursoInstructor ci, CursoFecha cf, Teorico t Where (p.idPersona = i.idPersona) And (ci.idInstructor = i.idInstructor) And (ci.idCF = cf.idCF) and (cf.NCurso=c.NCurso) and (t.NCurso=c.NCurso) and t.NTeorico= '" & cboTeorico.Text & "'"
        Call Ejecutar_Comando(sql, cn)
        
        If (rs.RecordCount <> 0) Then
            rs.MoveFirst
            Do While rs.EOF = False
                Varalf = rs.Fields(0)
                rs.MoveNext
            Loop
        End If
        
        lstListaHorario.AddItem "Nombre del Instructor: " & Varalf
        
        sql = "select p.apellido from persona p, instructor i, curso c, CursoInstructor ci, CursoFecha cf, Teorico t Where (p.idPersona = i.idPersona) And (ci.idInstructor = i.idInstructor) And (ci.idCF = cf.idCF) and (cf.NCurso=c.NCurso) and (t.NCurso=c.NCurso) and (t.NTeorico= '" & cboTeorico.Text & "')"
        Call Ejecutar_Comando(sql, cn)
    
        If (rs.RecordCount <> 0) Then
            rs.MoveFirst
            Do While rs.EOF = False
                Varalf = rs.Fields(0)
                rs.MoveNext
            Loop
        End If
        
        lstListaHorario.AddItem "Apellido del Instructor: " & Varalf
    Else
        sql = "select max(idAlumno) from alumno"
        Call Ejecutar_Comando(sql, cn)
        
        If (rs.RecordCount <> 0) Then
            rs.MoveFirst
            Do While rs.EOF = False
                Var = Val(rs.Fields(0))
                rs.MoveNext
            Loop
        End If
        
        sql = "select idCF from CursoFecha cf, teorico t, curso c Where (cf.NCurso = C.NCurso) and (t.NCurso=c.NCurso) and t.NTeorico='" & cboTeorico.Text & "'"
        Call Ejecutar_Comando(sql, cn)
        
        If (rs.RecordCount <> 0) Then
            rs.MoveFirst
            Do While rs.EOF = False
                Var2 = Val(rs.Fields(0))
                rs.MoveNext
            Loop
        End If
    
        sql = "insert into AlumnoCurso values (" & Var & ", " & Var2 & ", curdate(), 0)"
        Call Ejecutar_Comando(sql, cn)
        
        frmAsignarCurso.Visible = False
        frmGenerarComprobanteIns.Visible = True

        
    End If
    
End If

End Sub

Private Sub cmdVerifAlumno_Click()
frmVerificarAlumno.Visible = False

sql = "select idPersona from persona where dni='" & txtDNI.Text & "'"
Comando.ActiveConnection = cn

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

sql = "select idAlumno from alumno where idPersona=" & Var & ""
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var2 = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

If (Var <> 0 And Var2 <> 0) Then
    MsgBox ("Usted ya es alumno de la escuela.")
Else
    MsgBox ("Usted no es alumno de la escuela. Se lo registrara automaticamente")
    sql = "select idPersona from persona where dni='" & txtDNI.Text & "'"
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount = 0) Then
        sql = "insert into persona values ('', '" & txtNombre.Text & "', '" & txtApellido.Text & "', '" & txtDireccion.Text & "', '" & txtTelefono.Text & "', '" & txtEmail.Text & "', '" & txtDNI.Text & "')"
        Call Ejecutar_Comando(sql, cn)
       
    End If
    
    sql = "select MAX(idPersona) from persona"
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            Var = Val(rs.Fields(0))
            rs.MoveNext
        Loop
    End If
    
    sql = "insert into alumno values ('', " & Var & ")"
    Call Ejecutar_Comando(sql, cn)
    
    MsgBox ("Usted ha sido registrado exitosamente.")
    
End If
frmAsignarCurso.Visible = True
cmdSig.Visible = True

End Sub

Private Sub cmdVerificar_Click()
If (txtNombre.Text <> "" And txtApellido.Text <> "" And txtDireccion.Text <> "" And txtTelefono.Text <> "" And txtEmail.Text <> "" And txtDNI.Text <> "" And cboCurso.Text <> "") Then
    cmdVerificar.Default = False
        
    frmIngreso.Visible = False
    
    sql = "select idAlumno from alumno a, persona p where (p.idPersona=a.idPersona) and (dni = '" & txtDNI.Text & "')"
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            msj = MsgBox("El alumno ya está inscripto", vbExclamation)
            Form1.Hide
            Form0.Show
            rs.MoveNext
        Loop
    Else
        
        If (cboCurso.Text = "Teorico") Then
            flagteorico = 1
            sql = "select count(*) from AlumnoCurso ac, CursoFecha cf, teorico t Where ac.idcf = cf.idcf and cf.ncurso=t.ncurso"
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
                rs.MoveFirst
                Do While rs.EOF = False
                    Var = Val(rs.Fields(0))
                    rs.MoveNext
                Loop
            End If
            
            sql = "select sum(cupo) from teorico"
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
                rs.MoveFirst
                Do While rs.EOF = False
                    Var2 = Val(rs.Fields(0))
                    rs.MoveNext
                Loop
            End If
            
            resta = Var2 - Var
            
            If (resta > 0) Then
                msj = MsgBox("Existe cupo para los cursos teoricos. Se verificara si usted existe en la lista de espera")
                frmEspera.Visible = True
                cmdVerifListaEspera.Default = True
                
            Else
                msj = MsgBox("No hay cupo para los cursos teoricos. Se lo asignara a la lista de espera")
                frmAsignarEspera.Visible = True
                cmdAsignarEspera.Default = True
                
            End If
        
                    
        Else
            flagteorico = 0
            sql = "select cf.idCF from CursoFecha cf, Curso c, CursoInstructor ci Where (nombre = '" & cboCurso.Text & "') And (cf.ncurso = C.ncurso) And (ci.idCF = cf.idCF) limit 1"
            Call Ejecutar_Comando(sql, cn)
            
            If (rs.RecordCount <> 0) Then
                rs.MoveFirst
                Do While rs.EOF = False
                    Var = Val(rs.Fields(0))
                    rs.MoveNext
                Loop
            End If
            
            
            If (Var <> 0) Then
                MsgBox ("Existe disponibilidad para el curso que usted eligió. Se verificará si usted existe en la lista de espera")
                frmEspera.Visible = True
                cmdVerifListaEspera.Default = True
                
            Else
                MsgBox ("No hay cupo para el curso elegido. Se lo asignara a la lista de espera")
                frmAsignarEspera.Visible = True
                cmdAsignarEspera.Default = True
                
            End If
        End If
    End If
Else
    MsgBox ("INGRESE DATOS")
End If
    
End Sub

Private Sub cmdVerifListaEspera_Click()
frmEspera.Visible = False

sql = "select count(*) from persona p, listaespera le Where p.idpersona = le.idpersona and p.dni= '" & txtDNI.Text & "'"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

If (Var = 1) Then
    MsgBox ("Usted ya existe en la lista de espera. Se lo quitará de la misma y se le asignará un curso")
    cmdVerifListaEspera.Visible = True
Else
    MsgBox ("Usted no existe en la lista de espera. Se verificará si ya es alumno de la escuela.")
    frmVerificarAlumno.Visible = True
End If

End Sub

Private Sub cmdVerifPractico_Click()
frmPractico.Visible = False

sql = "select cf.idCF from CursoFecha cf, Curso c, CursoInstructor ci Where (nombre = '" & cboCurso.Text & "') And (cf.ncurso = C.ncurso) And (ci.idCF = cf.idCF) limit 1"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If


If (Var <> 0) Then
    MsgBox ("Existe disponibilidad para el curso que usted eligió. Se verificará si usted existe en la lista de espera")
    frmEspera.Visible = True
Else
    MsgBox ("No hay cupo para el curso elegido. Se lo asignara a la lista de espera")
    frmAsignarEspera.Visible = True
End If

End Sub

Private Sub cmdVerifTeorico_Click()
frmTeorico.Visible = False

sql = "select count(*) from AlumnoCurso ac, CursoFecha cf, teorico t Where ac.idcf = cf.idcf and cf.ncurso=t.ncurso"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

sql = "select sum(cupo) from teorico"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var2 = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

resta = Var2 - Var

If (resta > 0) Then
    msj = MsgBox("Existe cupo para los cursos teoricos. Se verificara si usted existe en la lista de espera")
    frmEspera.Visible = True
    
Else
    msj = MsgBox("No hay cupo para los cursos teoricos. Se lo asignara a la lista de espera")
    frmAsignarEspera.Visible = True
    
End If

End Sub

Private Sub cmdVolver_Click()
Form1.Hide
Form0.Show

frmIngreso.Visible = True

frmVerificarAlumno.Visible = False
txtNombre.Text = ""
txtApellido.Text = ""
txtDireccion.Text = ""
txtDNI.Text = ""
txtEmail.Text = ""
txtTelefono.Text = ""

End Sub

Private Sub Form_Activate()

    Set rs = New ADODB.Recordset
    With rs
        .CursorLocation = adUseClient
        .Open "curso", cn, adOpenStatic, adLockReadOnly, adCmdTable
    End With
End Sub

Private Sub Form_Load()
Form1.Hide
DESBLOQUEO.Show

cont = 0

Call Iniciar_Conexion
End Sub
