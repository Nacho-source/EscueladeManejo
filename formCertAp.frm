VERSION 5.00
Begin VB.Form formCertAp 
   Caption         =   "Certificado de Aprobacion"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9975
   LinkTopic       =   "Form3"
   Picture         =   "formCertAp.frx":0000
   ScaleHeight     =   7035
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "IMPRIMIR"
      Height          =   375
      Left            =   8160
      TabIndex        =   6
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label lblCalificacion 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   8640
      TabIndex        =   5
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblCurso 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label lblDNI 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   8280
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblApellido 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   6360
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   4800
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblFecha 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   2280
      Width           =   2055
   End
End
Attribute VB_Name = "formCertAp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
cmdImprimir.Visible = False
formCertAp.PrintForm
formCertAp.Visible = False
Form0.Visible = True
End Sub
