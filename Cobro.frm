VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Cobro"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7290
   LinkTopic       =   "Form2"
   Picture         =   "Cobro.frx":0000
   ScaleHeight     =   5415
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdListAlumno 
      Caption         =   "¿No recuerda el id del Alumno?"
      Height          =   615
      Left            =   4920
      TabIndex        =   10
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "VOLVER"
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Frame frmFormaPago 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione forma de pago:"
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   2295
      Begin VB.OptionButton optDebito 
         BackColor       =   &H80000009&
         Caption         =   "Débito"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "SIGUIENTE"
         Height          =   495
         Left            =   1080
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optCredito 
         BackColor       =   &H8000000E&
         Caption         =   "Crédito"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optEfectivo 
         BackColor       =   &H8000000E&
         Caption         =   "Efectivo"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame frmVerificarAlumno 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Verificar Alumno"
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4215
      Begin VB.CommandButton cmdEmitirFactura 
         Caption         =   "EMITIR FACTURA"
         Height          =   855
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "VERIFICAR"
         Default         =   -1  'True
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtidAlumno 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblidAlumno 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ingrese id del Alumno:"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEmitirFactura_Click()
sql = "insert into factura values ('', " & txtidAlumno.Text & ", curdate())"
Call Ejecutar_Comando(sql, cn)

sql = "select max(NFactura) from factura"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

sql = "select distinct c.NCurso from Curso c, AlumnoCurso ac, CursoFecha cf, Alumno a Where (a.idAlumno = ac.idAlumno) and (ac.idCF=cf.idCF) and (cf.NCurso=c.NCurso) and (a.idAlumno= " & txtidAlumno.Text & ")"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var2 = rs.Fields(0)
        rs.MoveNext
    Loop
End If

msj = InputBox("Ingrese monto pagado:")
Var3 = CSng(msj)

msj = InputBox("Ingrese observaciones:")
Varalf = msj

formFacturaCobro.lblObser.Caption = Varalf

sql = "insert into DetalleFactura values (" & Var & ", " & Var2 & ", " & Var3 & ", '" & Varalf & "')"
Call Ejecutar_Comando(sql, cn)

frmFormaPago.Visible = True
cmdEmitirFactura.Enabled = False
cmdVerificar.Enabled = False
txtidAlumno.Enabled = False


End Sub

Private Sub cmdListAlumno_Click()
Form9.Show
sql = "select idAlumno, nombre, apellido, dni, email, tel, direccion from alumno a, persona p where p.idPersona=a.idPersona order by idAlumno"
Call Ejecutar_Comando(sql, cn)

Set Form9.dgAlumno.DataSource = rs

End Sub

Private Sub cmdSiguiente_Click()
msj = MsgBox("¿Seguro que desea emitir la factura?", vbYesNo)
If (Val(msj) = 6) Then
    
    With formFacturaCobro
    
    sql = "select precio from curso c, AlumnoCurso ac, CursoFecha cf Where (C.NCurso = cf.NCurso) And (cf.idCF = ac.idCF) And (ac.idAlumno =" & txtidAlumno.Text & ")"
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            Var2 = Val(rs.Fields(0))
            rs.MoveNext
        Loop
    End If
    
    .lblPrecio.Caption = Var2
    .lblTotal2.Caption = Var3
    .lblTotal.Caption = Var3
     
    
    If (optEfectivo.Value = True) Then
        sql = "insert into pago values (''," & Var & ",'Efectivo', curdate(), " & Var3 & ")"
        Call Ejecutar_Comando(sql, cn)
        
        .lblFPago.Caption = "Efectivo"
        
        If (Var3 < Var2) Then
            Var = Var3 - Var2
            Varalf = MsgBox("Usted deberá un total de " & Var2, vbExclamation)
        End If
        
        formFacturaCobro.Visible = True
    Else
        If (optCredito.Value = True) Then
            sql = "insert into pago values (''," & Var & ",'Credito', curdate(), " & Var3 & ")"
            Call Ejecutar_Comando(sql, cn)
            
            .lblNomTarj.Visible = True
            .lblTarj.Visible = True
            .lblFPago = "Credito"
            .lblTarjeta.Visible = True
            .lblNTarjeta.Visible = True
                        
            formTarjeta.Show
            formTarjeta.lstTarjeta.Clear
            formTarjeta.lstTarjeta.AddItem "Beneficios con tarjetas: "
            formTarjeta.lstTarjeta.AddItem "Visa, Mastercard y American Express hasta 12 cuotas sin interés"
            formTarjeta.lstTarjeta.AddItem "Cencosud, Tarjeta Shopping y Nativa con"
            formTarjeta.lstTarjeta.AddItem "10% de recargo en 3 pagos y 15% de recargo en 6 pagos"
            formTarjeta.lstTarjeta.AddItem "---------------------------------"
            formTarjeta.lstTarjeta.AddItem ""
            
            formTarjeta.cboCuotas.Text = ""
            formTarjeta.cboTarjeta.Text = ""
            
            formTarjeta.Show
        Else
            If (optDebito.Value = True) Then
                sql = "insert into pago values (''," & Var & ",'Debito', curdate(), " & Var3 & ")"
                Call Ejecutar_Comando(sql, cn)
                
                msj = InputBox("¿Qué tarjeta va a utilizar?", "TARJETA")
                .lblNomTarj = msj
                .lblNomTarj.Visible = True
                .lblTarj.Visible = True
                .lblFPago.Caption = "Debito"
                msj = InputBox("Ingrese número de tarjeta: ", "TARJETA")
                .lblTarjeta.Visible = True
                .lblNTarjeta.Visible = True
                .lblTarjeta.Caption = msj
                
                If (Var3 < Var) Then
                    Var = Var3 - Var2
                    Varalf = MsgBox("Usted deberá un total de " & Var2, vbExclamation)
                End If
        
                
            Else
                msj = MsgBox("SELECCIONE UN METODO DE PAGO", , "ERROR")
            End If
        End If
    End If
        
    Form2.Visible = False
    
    sql = "select max(Nfactura) from factura"
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            Var2 = Val(rs.Fields(0))
            rs.MoveNext
        Loop
    End If
    
    .lblFactura.Caption = Var2
    
    .lblFecha.Caption = Date
    
    sql = "select apellido from persona p, alumno a where (a.idPersona=p.idPersona) and (a.idAlumno=" & txtidAlumno.Text & ")"
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            Varalf = rs.Fields(0)
            rs.MoveNext
        Loop
    End If
    
    .lblApellido.Caption = Varalf
    
    sql = "select nombre from persona p, alumno a where (a.idPersona=p.idPersona) and (a.idAlumno=" & txtidAlumno.Text & ")"
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            Varalf = rs.Fields(0)
            rs.MoveNext
        Loop
    End If
    
    .lblNombre.Caption = Varalf
    
    sql = "select direccion from persona p, alumno a where (a.idPersona=p.idPersona) and (a.idAlumno='" & txtidAlumno.Text & "')"
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            Varalf = rs.Fields(0)
            rs.MoveNext
        Loop
    End If
    
    .lblDireccion.Caption = Varalf
    
    sql = "select nombre from curso c, AlumnoCurso ac, CursoFecha cf Where (C.NCurso = cf.NCurso) And (cf.idCF = ac.idCF) And (ac.idAlumno =" & txtidAlumno.Text & ")"
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            Varalf = rs.Fields(0)
            rs.MoveNext
        Loop
    End If
    
    .lblCurso.Caption = Varalf
    

    End With
Else
    MsgBox ("La factura no será emitida")
    
End If
End Sub

Private Sub cmdVerificar_Click()
sql = "select f.NFactura from factura f, pago p where f.nfactura = p.nfactura and idAlumno=" & txtidAlumno.Text & ""
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    MsgBox ("Su pago ya ha sido abonado")
Else
    cmdVerificar.Default = False
    cmdEmitirFactura.Default = True
    
    sql = "select idAlumno from Alumno where idAlumno=" & txtidAlumno.Text & ""
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        MsgBox ("El Comprobante es válido")
        cmdEmitirFactura.Visible = True
    Else
        msj = MsgBox("El comprobante no es válido", vbExclamation, "ERROR")
    End If
End If
End Sub

Private Sub cmdVolver_Click()
Form0.Show
Form2.Hide

frmFormaPago.Visible = False
cmdEmitirFactura.Enabled = True
cmdVerificar.Enabled = True
txtidAlumno.Enabled = True
txtidAlumno.Text = ""
cmdEmitirFactura.Visible = False

formFacturaCobro.lblObser.Caption = ""

End Sub

Private Sub Form_Activate()
If (flag = 0) Then
    txtidAlumno.SetFocus
    flag = 1
End If

End Sub

Private Sub optCredito_Click()
If (optCredito.Value = True) Then
    optEfectivo.Value = False
    optDebito.Value = False
End If

End Sub

Private Sub optDebito_Click()
If (optDebito.Value = True) Then
    optEfectivo.Value = False
    optCredito.Value = False
End If
End Sub

Private Sub optEfectivo_Click()
If (optEfectivo.Value = True) Then
    optCredito.Value = False
    optDebito.Value = False
End If

End Sub
