VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form formVehiculos 
   BackColor       =   &H8000000E&
   Caption         =   "Vehiculos"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7860
   LinkTopic       =   "Form2"
   ScaleHeight     =   5985
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dgVehiculos 
      Height          =   3975
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7011
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
            LCID            =   3082
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
            LCID            =   3082
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
   Begin VB.TextBox txtCodVehiculo 
      Enabled         =   0   'False
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton cmdSiguiente 
      Caption         =   "SIGUIENTE"
      Height          =   735
      Left            =   6000
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCargarGrilla 
      Caption         =   "CARGAR GRILLA"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label lblVehi 
      BackColor       =   &H80000009&
      Caption         =   "Haga doble click en el vehículo a seleccionar:"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label lblInserte 
      BackColor       =   &H8000000E&
      Caption         =   "CODIGO DEL VEHICULO:"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   5040
      Width           =   1455
   End
End
Attribute VB_Name = "formVehiculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCargarGrilla_Click()
lblVehi.Visible = True

sql = "select v.CodVehiculo, matricula, marca, modelo from vehiculo v where utilizable=1 and  v.CodVehiculo not in(select v.CodVehiculo from vehiculo v, AlumnoVehiculo av where v.CodVehiculo=av.CodVehiculo)"
Call Ejecutar_Comando(sql, cn)


Set dgVehiculos.DataSource = rs

    

End Sub


Private Sub cmdSiguiente_Click()

If (Val(txtCodVehiculo.Text) = 0) Then
    MsgBox ("SELEECIONE VEHICULO A UTILIZAR")
Else
    formVehiculos.Visible = False
    Form1.Visible = True
    
    sql = "select marca from vehiculo where CodVehiculo = '" & txtCodVehiculo & "' "
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            Varalf = rs.Fields(0)
            rs.MoveNext
        Loop
    End If
    
    sql = "select modelo from vehiculo where CodVehiculo = '" & txtCodVehiculo & "' "
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            Varalf2 = rs.Fields(0)
            rs.MoveNext
        Loop
    End If
    
    sql = "select matricula from vehiculo where CodVehiculo = " & txtCodVehiculo & ""
    Call Ejecutar_Comando(sql, cn)
    
    If (rs.RecordCount <> 0) Then
        rs.MoveFirst
        Do While rs.EOF = False
            Varalf3 = rs.Fields(0)
            rs.MoveNext
        Loop
    End If
    
    Form1.lstListaHorario.AddItem ""
    Form1.lstListaHorario.AddItem "Marca del Vehiculo: " & Varalf
    Form1.lstListaHorario.AddItem "Modelo del Vehiculo: " & Varalf2
    Form1.lstListaHorario.AddItem "Matricula del Vehiculo: " & Varalf3
      
End If

End Sub


Private Sub dgVehiculos_Click()
dgVehiculos.MarqueeStyle = dbgHighlightRowRaiseCell
Var = dgVehiculos.Columns(0)
txtCodVehiculo.Text = Var
End Sub
