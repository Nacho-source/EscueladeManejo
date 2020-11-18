VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Liquidacion de Sueldo"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9435
   LinkTopic       =   "Form3"
   Picture         =   "3) Liquidacion de Sueldos.frx":0000
   ScaleHeight     =   6075
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdListaInstructor 
      Caption         =   "¿No recuerda el id del Instructor?"
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "VOLVER"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid dgJustificativo 
      Height          =   2295
      Left            =   4080
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tardes y Faltas Justificadas"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dgAsistencia 
      Height          =   2655
      Left            =   4080
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Registro de Asistencia"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstInfo 
      Height          =   1035
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "OBTENER DATOS"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtInstructor 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "Datos del Instructor:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblInstructor 
      BackColor       =   &H8000000E&
      Caption         =   "Ingrese el id del Instructor:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInfo_Click()
If (txtInstructor.Text = "") Then
    msj = MsgBox("INGRESE DATOS", vbExclamation, "ERROR")
Else
    sql = "select idInstructor from instructor where idinstructor=" & txtInstructor.Text & ""
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount = 0) Then
        msj = MsgBox("EL INSTRUCTOR NO EXISTE. REVISE SUS INGRESOS", vbCritical, "ERROR")
            
    Else
        lstInfo.Clear
        lstInfo.Visible = True
        dgJustificativo.ClearFields
        dgJustificativo.Visible = True
        dgAsistencia.ClearFields
        dgAsistencia.Visible = True
        
        sql = "select nombre from instructor i, persona p Where (i.idPersona = p.idPersona) and idInstructor= " & txtInstructor.Text & ""
        Call Ejecutar_Comando(sql, cn)
        
        If (rs.RecordCount <> 0) Then
            rs.MoveFirst
            Do While rs.EOF = False
                Varalf = rs.Fields(0)
                rs.MoveNext
            Loop
        End If
        
        lstInfo.AddItem "Nombre del Instructor: " & Varalf
        
        sql = "select apellido from instructor i, persona p where (i.idPersona=p.idPersona) and idInstructor= " & txtInstructor.Text & ""
        Call Ejecutar_Comando(sql, cn)
        
        If (rs.RecordCount <> 0) Then
            rs.MoveFirst
            Do While rs.EOF = False
                Varalf = rs.Fields(0)
                rs.MoveNext
            Loop
        End If
        
        lstInfo.AddItem "Apellido del Instructor: " & Varalf
        
        sql = "select FIngreso from instructor where idInstructor=" & txtInstructor & ""
        Call Ejecutar_Comando(sql, cn)
        
        If (rs.RecordCount <> 0) Then
            rs.MoveFirst
            Do While rs.EOF = False
                Varalf = rs.Fields(0)
                rs.MoveNext
            Loop
        End If
        
        lstInfo.AddItem "Fecha de ingreso a la escuela: " & Varalf
        
        sql = "select fecha, ra.HoraEntrada Hora_de_llegada, h.HoraEntrada Horario_de_entrada from RegistroAsistencia ra, instructor i, persona p, CursoFecha cf, horario h Where (ra.idCF = cf.idCF) and (cf.NHorario=h.NHorario) and (ra.idPersona=p.idPersona) and (p.idPersona=i.idPersona) and (i.idInstructor= " & txtInstructor.Text & ")"
        Call Ejecutar_Comando(sql, cn)
        
        Set dgAsistencia.DataSource = rs 'arreglar error
        
        sql = "select NJustificativo, FJustificada from justificativo j, RegistroAsistencia ra, persona p, instructor i where (ra.CodAsis=j.CodAsis) and (p.idPersona=i.idPersona) and (ra.idPersona=p.idPersona) and (i.idInstructor=" & txtInstructor.Text & ")"
        Call Ejecutar_Comando(sql, cn)
        
        Set dgJustificativo.DataSource = rs
    End If
End If

End Sub

Private Sub cmdListaInstructor_Click()
Form9.Show
sql = "select idInstructor, nombre, apellido, dni, email, tel, direccion from instructor i, persona p where p.idPersona=i.idPersona order by idInstructor"
Call Ejecutar_Comando(sql, cn)

Set Form9.dgAlumno.DataSource = rs
End Sub

Private Sub cmdVolver_Click()
txtInstructor.Text = ""
lblInfo.Visible = False
lstInfo.Visible = False
dgAsistencia.Visible = False
dgJustificativo.Visible = False
lstInfo.Clear
dgAsistencia.ClearFields
dgJustificativo.ClearFields

Form3.Visible = False
Form0.Visible = True
End Sub

Private Sub Form_Activate()
txtInstructor.SetFocus
End Sub

