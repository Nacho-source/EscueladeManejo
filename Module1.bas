Attribute VB_Name = "Module1"
Option Explicit

Public cn As ADODB.Connection
Public rs As ADODB.Recordset
Public rs2 As ADODB.Recordset
Public rs3 As ADODB.Recordset
Public rsConsulta  As New ADODB.Recordset
Public Comando As New ADODB.Command
Public sql As String, sql2 As String, sql3 As String, Varalf As String, Varalf2 As String, Varalf3 As String, msj As String
Public Modi As Boolean
Public resta As Integer, aux As Integer
Public Entrada As Single, Salida As Single, total As Single
Public varlong As Long
Public varfloat As Single
Public flagteorico As Integer, flagpass As Integer, flag2 As Integer, Mes As Integer, Dia As Integer, flag As Integer, R As Integer, M As Integer, C As Integer, Var As Integer, Var2 As Integer, Var3 As Integer, Var4 As Integer, cont As Integer


Public Sub Iniciar_Conexion()

    If cn Is Nothing Then
        Set cn = New ADODB.Connection
        With cn
            .CursorLocation = adUseClient
         
            .Open "Driver={MySQL ODBC 5.1 Driver};" & "Server=localhost;" & "Port=3306;Database=Escuela_Manejo;" & _
                      "User=root;Password=;Option=3;"
        End With
    End If

End Sub

Public Sub Ejecutar_Comando(consulta As String, conexion As ADODB.Connection)

    With Comando
        .ActiveConnection = conexion
        .CommandText = consulta
        
        Set rs = .Execute()
    End With
    
    Set Comando = Nothing
    
End Sub



