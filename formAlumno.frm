VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form9 
   BackColor       =   &H80000009&
   Caption         =   "Lista de Alumnos"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   LinkTopic       =   "Form9"
   ScaleHeight     =   6015
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSiguiente 
      Caption         =   "SIGUIENTE"
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAlumno 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "CERRAR"
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   5160
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid dgAlumno 
      Height          =   4575
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8070
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
            LCID            =   11274
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
            LCID            =   11274
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
   Begin VB.Label lblAlumno2 
      BackColor       =   &H80000009&
      Caption         =   "id del Alumno:"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblAlumno 
      BackColor       =   &H80000009&
      Caption         =   "Haga doble click en el alumno a seleccionar:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   4455
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
Form9.Hide
End Sub

Private Sub cmdSiguiente_Click()
If (txtAlumno.Text = "") Then
    msj = MsgBox("Seleccione un alumno", vbExclamation)
    Else
    Var3 = InputBox("Ingrese monto de la rotura")
    Form6.lstVehiculo.AddItem "Monto de la Rotura: " & Var3
    
    sql = "select nombre, apellido from alumno a, persona p where p.idpersona=a.idpersona and idalumno='" & Var & "'"
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While (rs.EOF = False)
            Varalf = rs.Fields(0) & " " & rs.Fields(1)
            rs.MoveFirst
        Loop
    End If
    
    Form6.lstVehiculo.AddItem "Alumno que causó la rotura: " & Varalf
    
    lblAlumno.Visible = False
    lblAlumno2.Visible = False
    txtAlumno.Visible = False
    cmdSiguiente.Visible = False
    cmdCerrar.Visible = True
    Form9.Hide
End If
End Sub

Private Sub dgAlumno_Click()
dgAlumno.MarqueeStyle = dbgHighlightRowRaiseCell
Var2 = dgAlumno.Columns(0)
txtAlumno.Text = Var2

End Sub
