VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Cierre de Caja"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form3"
   Picture         =   "8) Cierre de Caja.frx":0000
   ScaleHeight     =   6375
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVolver 
      Caption         =   "VOLVER"
      Height          =   855
      Left            =   360
      TabIndex        =   8
      Top             =   5280
      Width           =   1455
   End
   Begin VB.ListBox lstCierreCaja 
      Height          =   1815
      Left            =   360
      TabIndex        =   7
      Top             =   3240
      Width           =   3015
   End
   Begin VB.CommandButton cmdSaldoFinal 
      Caption         =   "CALCULAR SALDO FINAL"
      Height          =   975
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtSalida 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtEntrada 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdMontoSalida 
      Caption         =   "CALCULAR MONTO DE SALIDA"
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdMontoEntrada 
      Caption         =   "CALCULAR MONTO DE ENTRADA"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblSalida 
      BackColor       =   &H8000000E&
      Caption         =   "Monto de Salida:"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblEntrada 
      BackColor       =   &H8000000E&
      Caption         =   "Monto de Entrada:"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMontoEntrada_Click()
lblEntrada.Visible = True
txtEntrada.Visible = True

sql = "select NPago from pago where FPago = curdate()"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    flag = 1
End If

sql = "select CodRotura from pagorotura where FPago=curdate()"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0 Or flag = 1) Then
    If (rs.RecordCount <> 0) Then
        sql = "select monto from rotura r, pagorotura pr where r.codrotura=pr.codrotura and fpago=curdate()"
        Call Ejecutar_Comando(sql, cn)
        
        If (rs.RecordCount <> 0) Then
            rs.MoveFirst
            Do While rs.EOF = False
                Var2 = Var2 + Val(rs.Fields(0))
                rs.MoveNext
            Loop
        End If
    End If
    
    If (flag = 1) Then
        sql = "select (Monto) from Pago Where FPago=(select curdate())"
        Call Ejecutar_Comando(sql, cn)
        
        rs.MoveFirst
        Do While rs.EOF = False
            Var = Var + Val(rs.Fields(0))
            rs.MoveNext
        Loop
    End If
    
    Var3 = Var + Var2
    
Else
    msj = MsgBox("No hubo entradas en el día.", vbExclamation)
End If

txtEntrada.Text = Var3
Entrada = Val(Var3)

cmdMontoEntrada.Visible = False
cmdSaldoFinal.Visible = True

End Sub

Private Sub cmdMontoSalida_Click()
lblSalida.Visible = True
txtSalida.Visible = True

Mes = Format(Date, "m")

sql = "select sum(monto) from sueldo where (month(curdate())=" & Mes & ")"
Call Ejecutar_Comando

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

txtSalida.Text = Var
Salida = Val(Var)

cmdMontoEntrada.Visible = True
cmdMontoSalida.Visible = False

End Sub

Private Sub cmdSaldoFinal_Click()

sql = "select SaldoInicial from caja where FCaja=curdate()"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Var = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If

Var = Val(Var)
Var2 = Var + Entrada


sql = "insert into Caja values ('', curdate(), 0, " & Var2 & ", " & Entrada & ", " & Salida & ")"
Call Ejecutar_Comando(sql, cn)

lstCierreCaja.AddItem "Fecha: " & Date
lstCierreCaja.AddItem "Saldo Inicial del día: " & 0
lstCierreCaja.AddItem "Saldo Final del día: " & Var2
lstCierreCaja.AddItem "Total de ingresos del día: " & Entrada
lstCierreCaja.AddItem "Total de egresos del día: " & Salida

lblEntrada.Visible = False
lblSalida.Visible = False
txtEntrada.Visible = False
txtSalida.Visible = False
cmdSaldoFinal.Visible = False

End Sub

Private Sub cmdVolver_Click()
Form8.Hide
Form0.Show
End Sub
