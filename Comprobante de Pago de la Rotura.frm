VERSION 5.00
Begin VB.Form formComprobantePagoRotura 
   Caption         =   "Comprobante de Pago de la Rotura"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   LinkTopic       =   "Form11"
   Picture         =   "Comprobante de Pago de la Rotura.frx":0000
   ScaleHeight     =   6345
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "IMPRIMIR"
      Height          =   735
      Left            =   7320
      TabIndex        =   0
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label lblFecha 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label lblMatricula 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   6720
      TabIndex        =   2
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label lblAlumno 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   5520
      TabIndex        =   1
      Top             =   3720
      Width           =   2655
   End
End
Attribute VB_Name = "formComprobantePagoRotura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
formComprobantePagoRotura.PrintForm
Form11.Hide
Form0.Show
End Sub
