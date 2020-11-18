VERSION 5.00
Begin VB.Form Form0 
   BackColor       =   &H00000000&
   Caption         =   "INICIO"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form2"
   ScaleHeight     =   7005
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmInicio 
      BackColor       =   &H00FFFFFF&
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.CommandButton cmdPagoRotura 
         Caption         =   "PROCESAR PAGO DE ROTURA"
         Height          =   855
         Left            =   5760
         TabIndex        =   11
         Top             =   4440
         Width           =   1935
      End
      Begin VB.CommandButton cmdVolver 
         Caption         =   "VOLVER"
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   5880
         Width           =   1455
      End
      Begin VB.CommandButton cmdProceso 
         Caption         =   "PROCESAR CIERRE DE CAJA"
         Height          =   855
         Left            =   3600
         TabIndex        =   10
         Top             =   4440
         Width           =   1935
      End
      Begin VB.CommandButton cmdGestNota 
         Caption         =   "GESTIONAR NOTA DE APRENDIZ"
         Height          =   855
         Left            =   1320
         TabIndex        =   9
         Top             =   4440
         Width           =   1935
      End
      Begin VB.CommandButton cmdRevision 
         Caption         =   "REVISAR VEHICULO PARA CURSO"
         Height          =   855
         Left            =   5760
         TabIndex        =   8
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CommandButton cmdGestion 
         Caption         =   "GESTIONAR CERTIFICADO DE APROBACION "
         Height          =   855
         Left            =   3600
         TabIndex        =   7
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CommandButton cmdToma 
         Caption         =   "TOMAR ASISTENCIA"
         Height          =   855
         Left            =   1320
         TabIndex        =   6
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CommandButton cmdLiquidacion 
         Caption         =   "LIQUIDAR SUELDO"
         Height          =   855
         Left            =   5760
         TabIndex        =   5
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CommandButton cmdCobro 
         Caption         =   "COBRAR CURSO"
         Height          =   855
         Left            =   3600
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CommandButton cmdInscripcion 
         Caption         =   "INSCRIBIR ALUMNO"
         Height          =   855
         Left            =   1320
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblElija 
         BackColor       =   &H80000005&
         Caption         =   "Elija la accion que desea realizar:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   4095
      End
      Begin VB.Label lblINICIO 
         BackColor       =   &H80000005&
         Caption         =   "BIENVENIDOS A LA ESCUELA DE MANEJO ""MASCHER"""
         BeginProperty Font 
            Name            =   "Berlin Sans FB Demi"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   8175
      End
   End
End
Attribute VB_Name = "Form0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCobro_Click()
Form0.Hide
Form2.Show
flag = 0
End Sub

Private Sub cmdGestion_Click()
Form0.Hide
With Form5

.Show
.txtCurso.Enabled = True
.txtidAlumno.Enabled = True
.txtCurso.Text = ""
.txtidAlumno.Text = ""

End With
End Sub

Private Sub cmdGestNota_Click()
Form7.Show
Form0.Hide
End Sub

Private Sub cmdInscripcion_Click()
Form0.Hide
With Form1
.Show

cont = 0

.frmIngreso.Visible = True


.frmVerificarAlumno.Visible = False
.txtNombre.Text = ""
.txtApellido.Text = ""
.txtDireccion.Text = ""
.txtDNI.Text = ""
.txtEmail.Text = ""
.txtTelefono.Text = ""
.frmAsignarCurso.Visible = False
.frmAsignarEspera.Visible = False
.frmEspera.Visible = False
.frmGenerarComprobante.Visible = False
.frmGenerarComprobanteIns.Visible = False
.frmQuitarEspera.Visible = False
.frmVerificarAlumno.Visible = False

End With

End Sub

Private Sub cmdLiquidacion_Click()
Form0.Hide
Form3.Show
End Sub

Private Sub cmdPagoRotura_Click()
Form10.Show
Form0.Hide

End Sub

Private Sub cmdProceso_Click()

Form0.Hide
Form8.Show

With Form8

    .lstCierreCaja.Clear
    
    Dia = Format(Date, "d")
    Mes = Format(Date, "m")
    Select Case Mes
        Case 1, 3, 5, 7, 8, 10, 11, 12
        If (Dia = 31) Then
            .cmdMontoSalida.Visible = True
        Else
            .cmdMontoEntrada.Visible = True
        End If
        
        Case 4, 6, 9, 11
        If (Dia = 30) Then
            .cmdMontoSalida.Visible = True
        Else
            .cmdMontoEntrada.Visible = True
        End If
        
        Case 2
        If (Dia = 28 Or Dia = 29) Then
            .cmdMontoSalida.Visible = True
        Else
            .cmdMontoEntrada.Visible = True
        End If
    End Select
End With

End Sub

Private Sub cmdRegistro_Click()
Form0.Hide
Form7.Show
End Sub

Private Sub cmdRevision_Click()
Form0.Hide
Form6.Show
flag = 0
End Sub

Private Sub cmdToma_Click()
Form0.Hide
Form4.Show
Form4.txtHora = Time
Form4.txtHorario = Time
End Sub

Private Sub cmdVolver_Click()
Form0.Hide
DESBLOQUEO.Show
End Sub

Private Sub Form_Load()
Form0.Show
Form1.Hide
Form2.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
formCertAp.Hide
formCompIns.Hide
formComprobanteEspera.Hide
formFacturaCobro.Hide
formInformeDesapruebo.Hide
formVehiculos.Hide
End Sub

