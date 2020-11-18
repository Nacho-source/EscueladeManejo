VERSION 5.00
Begin VB.Form formFacturaCobro 
   Caption         =   "Factura C"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form3"
   Picture         =   "formFacturaCobro.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "IMPRIMIR"
      Height          =   735
      Left            =   8160
      TabIndex        =   9
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblNomTarj 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   9360
      TabIndex        =   17
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label lblTarj 
      BackColor       =   &H80000009&
      Caption         =   "Tarjeta:"
      Height          =   255
      Left            =   8280
      TabIndex        =   16
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label lblNTarjeta 
      BackColor       =   &H80000009&
      Caption         =   "Numero de la Tarjeta:"
      Height          =   375
      Left            =   8280
      TabIndex        =   15
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label lblTarjeta 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   14
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label lblFPago 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   13
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Forma de Pago:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label lblObser 
      BackColor       =   &H80000009&
      Height          =   1455
      Left            =   1320
      TabIndex        =   11
      Top             =   6960
      Width           =   6855
   End
   Begin VB.Label lblFactura 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   10
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblCant 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label lblTotal2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   7
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   6
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label lblPrecio 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   5
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label lblCurso 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   6240
      Width           =   6855
   End
   Begin VB.Label lblFecha 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label lblDireccion 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   3840
      Width           =   5175
   End
   Begin VB.Label lblApellido 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   3360
      Width           =   2175
   End
End
Attribute VB_Name = "formFacturaCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
formFacturaCobro.Visible = False
Form0.Visible = True
cmdImprimir.Visible = False
formFacturaCobro.PrintForm
Form2.txtidAlumno.Enabled = True
End Sub
