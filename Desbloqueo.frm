VERSION 5.00
Begin VB.Form DESBLOQUEO 
   Caption         =   "Form9"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6405
   LinkTopic       =   "Form9"
   Picture         =   "Desbloqueo.frx":0000
   ScaleHeight     =   4470
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtContrase�a 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdDesbloquar 
      Caption         =   "DESBLOQUEAR"
      Default         =   -1  'True
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   3480
      Width           =   1695
   End
End
Attribute VB_Name = "DESBLOQUEO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDesbloquar_Click()
Dim PassWord As String

sql = "select contrase�a from contrase�a where NContrase�a=(select max(NContrase�a) from contrase�a)"
Call Ejecutar_Comando(sql, cn)

If (rs.RecordCount <> 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        varlong = Val(rs.Fields(0))
        rs.MoveNext
    Loop
End If


If ((Val(txtContrase�a.Text) = varlong)) Then
    MsgBox ("Bienvenido a la Escuela de Manejo Mascher")
    If (flagpass = 0) Then
        Call inputbox_Password(DESBLOQUEO, "*")
        
        PassWord = InputBox("Ingrese nueva contrase�a(num�rica):", App.Title)
    
        sql = "insert into contrase�a values(''," & Val(PassWord) & ")"
        Call Ejecutar_Comando(sql, cn)
        
        flagpass = 1
    End If
    
    Form0.Show
    DESBLOQUEO.Hide
Else
    
    msj = MsgBox("La contrase�a no es correcta", , "ERROR")
End If

End Sub


Private Sub Form_Activate()
txtContrase�a.SetFocus
Call Iniciar_Conexion
End Sub
