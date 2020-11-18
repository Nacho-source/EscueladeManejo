VERSION 5.00
Begin VB.Form formInformeDesapruebo 
   Caption         =   "Informe de Desapruebo"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9960
   LinkTopic       =   "Form3"
   Picture         =   "formInformeDesapruebo.frx":0000
   ScaleHeight     =   7065
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "IMPRIMIR"
      Height          =   495
      Left            =   7440
      TabIndex        =   6
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label lblCalificacion 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label lblCurso 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label lblDNI 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblApellido 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label lblFecha 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   2040
      Width           =   1935
   End
End
Attribute VB_Name = "formInformeDesapruebo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
cmdImprimir.Visible = False
formInformeDesapruebo.PrintForm
formInformeDesapruebo.Hide
Form0.Show
End Sub
