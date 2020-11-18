VERSION 5.00
Begin VB.Form formTarjeta 
   BackColor       =   &H80000009&
   Caption         =   "Tarjeta"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   LinkTopic       =   "Form10"
   ScaleHeight     =   3285
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSiguiente 
      Caption         =   "Siguiente"
      Height          =   615
      Left            =   4920
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "VOLVER"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox cboCuotas 
      Height          =   315
      ItemData        =   "formTarjeta.frx":0000
      Left            =   3240
      List            =   "formTarjeta.frx":0010
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.ListBox lstTarjeta 
      Height          =   1425
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   5895
   End
   Begin VB.ComboBox cboTarjeta 
      Height          =   315
      ItemData        =   "formTarjeta.frx":0021
      Left            =   360
      List            =   "formTarjeta.frx":0037
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblCuotas 
      BackColor       =   &H80000009&
      Caption         =   "Elija cantidad de cuotas:"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblTarjeta 
      BackColor       =   &H80000009&
      Caption         =   "Elija tarjeta con la cual abonar:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "formTarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCuotas_Click()
lstTarjeta.AddItem "Cantidad de Cuotas: " & cboCuotas.Text
flag = 0
flag2 = 1

sql = "select precio from curso c, AlumnoCurso ac, CursoFecha cf Where (C.NCurso = cf.NCurso) And (cf.idCF = ac.idCF) And (ac.idAlumno =" & Form2.txtidAlumno.Text & ")"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

If ((cboTarjeta.Text = "Cencosud") Or (cboTarjeta.Text = "Tarjeta Shopping") Or (cboTarjeta.Text = "Nativa")) Then
    If (Val(cboCuotas.Text) = 3) Then
        flag = 1
        total = Var + (Var * 10 / 100)
        lstTarjeta.AddItem "El pago tendrá un recargo del 10%. Total a pagar: " & total
    Else
        If (Val(cboCuotas.Text) > 3) Then
            flag = 1
            total = Var + (Var * 15 / 100)
            lstTarjeta.AddItem "El pago tendrá un recargo del 15%. Total a Pagar: " & total
        End If
    End If
End If

End Sub

Private Sub cboTarjeta_Click()
formTarjeta.lstTarjeta.Clear
formTarjeta.lstTarjeta.AddItem "Beneficios con tarjetas: "
formTarjeta.lstTarjeta.AddItem "Visa, Mastercard y American Express hasta 12 cuotas sin interés"
formTarjeta.lstTarjeta.AddItem "Cencosud, Tarjeta Shopping y Nativa con"
formTarjeta.lstTarjeta.AddItem "10% de recargo en 3 pagos y 15% de recargo en 6 pagos"
formTarjeta.lstTarjeta.AddItem "---------------------------------"
formTarjeta.lstTarjeta.AddItem ""
lstTarjeta.AddItem "Tarjeta: " & cboTarjeta.Text
msj = InputBox("Ingrese número de tarjeta: ", "TARJETA")
lstTarjeta.AddItem "Numero de la tarjeta: " & msj

lstTarjeta.AddItem "Tarjeta: " & cboTarjeta.Text
End Sub

Private Sub cmdSiguiente_Click()
If (flag2 = 0) Then
    Varalf = MsgBox("INGRESE DATOS", vbExclamation, "ERROR")
Else
    formTarjeta.Hide
    formFacturaCobro.Show
    formFacturaCobro.lblNomTarj.Caption = cboTarjeta.Text
    formFacturaCobro.lblTarjeta.Caption = msj
    
    If (flag = 1) Then
        formFacturaCobro.lblTotal = total
        formFacturaCobro.lblTotal2 = total
    End If
            
    If (total < Var) Then
        Var2 = total - Var
        Varalf = MsgBox("Usted deberá un total de " & Var2, vbExclamation)
    End If
End If
End Sub

Private Sub cmdVolver_Click()
formTarjeta.Hide
Form2.Show
End Sub

Private Sub Form_Load()
flag2 = 0
End Sub
