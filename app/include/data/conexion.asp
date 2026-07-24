<%
' Punto UNICO de apertura de conexion: si la BD no responde (red caida, credenciales
' incorrectas, servidor parado, etc.) no debe quedar nunca en manos de la configuracion
' de IIS decidir si se ve un error tecnico o una pagina en blanco - lo controlamos aqui,
' con timeout corto para fallar rapido en vez de colgar la peticion, y redirigiendo
' siempre a un aviso generico. El detalle tecnico se manda al log de IIS (AppendToLog),
' nunca al navegador.
Function NuevaConexion()
    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.ConnectionString = Application("DBConnectionString")
    conn.ConnectionTimeout = 10

    On Error Resume Next
    conn.Open
    If Err.Number <> 0 Then
        Dim numeroError, descripcionError
        numeroError = Err.Number
        descripcionError = Err.Description
        On Error Goto 0
        Response.AppendToLog " ToperBTR_ErrorConexionBD=" & numeroError & ":" & descripcionError
        Response.Redirect "/errors/error.asp?msg=bd"
        Response.End
    End If
    On Error Goto 0

    Set NuevaConexion = conn
End Function

Function TipoAdoParaValor(valor)
    Select Case VarType(valor)
        Case 2, 3, 17 ' vbInteger, vbLong, vbByte
            TipoAdoParaValor = 3 ' adInteger
        Case 5 ' vbDouble
            TipoAdoParaValor = 5 ' adDouble
        Case 7 ' vbDate
            TipoAdoParaValor = 7 ' adDate
        Case 11 ' vbBoolean
            TipoAdoParaValor = 11 ' adBoolean
        Case Else
            TipoAdoParaValor = 200 ' adVarChar
    End Select
End Function

Function CrearComando(conn, sql, parametros)
    Dim cmd, i, valor, tipo, tamano
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = 1 ' adCmdText
    If IsArray(parametros) Then
        For i = 0 To UBound(parametros)
            valor = parametros(i)
            tipo = TipoAdoParaValor(valor)
            If tipo = 200 Then
                tamano = Len(valor & "") + 1
            Else
                tamano = 0
            End If
            If IsNull(valor) Then
                cmd.Parameters.Append cmd.CreateParameter("p" & i, tipo, 1, tamano)
                cmd.Parameters("p" & i).Value = Null
            Else
                cmd.Parameters.Append cmd.CreateParameter("p" & i, tipo, 1, tamano, valor)
            End If
        Next
    End If
    Set CrearComando = cmd
End Function

' Variantes "Tx": NO capturan errores, los dejan propagar via el objeto Err (global en
' VBScript) para que el llamador decida que hacer. USAR SOLO dentro de funciones que ya
' gestionan su propia transaccion (BeginTrans/CommitTrans/RollbackTrans) y necesitan
' comprobar Err.Number paso a paso para decidir si hacen rollback - si esas funciones
' llamaran a las variantes seguras de abajo, un fallo redirigiria y terminaria la
' respuesta ANTES de que la funcion transaccional pudiera hacer su propio RollbackTrans,
' dejando la transaccion a medias. Tambien las usan los catalogos (catalogos.asp), porque
' se llaman indistintamente desde dentro y fuera de transacciones.
Function EjecutarConsultaTx(conn, sql, parametros)
    Dim cmd
    Set cmd = CrearComando(conn, sql, parametros)
    Set EjecutarConsultaTx = cmd.Execute()
End Function

Sub EjecutarNonQueryTx(conn, sql, parametros)
    Dim cmd
    Set cmd = CrearComando(conn, sql, parametros)
    cmd.Execute
End Sub

' Variantes seguras por defecto: si la consulta falla (bloqueo de tabla, desconexion a
' mitad de peticion, timeout, etc.) no debe quedar nunca como un error sin control (igual
' que ya se corrigio para el fallo de apertura de conexion en NuevaConexion) - se registra
' el detalle en el log de IIS y se redirige a un aviso generico, nunca se expone nada
' tecnico al navegador. Son las que debe usar cualquier lectura que NO forme parte de una
' transaccion propia (la inmensa mayoria de listados, detalles y busquedas).
Function EjecutarConsulta(conn, sql, parametros)
    On Error Resume Next
    Dim resultado
    Set resultado = EjecutarConsultaTx(conn, sql, parametros)
    Call RedirigirSiErrorSQL(sql)
    On Error Goto 0
    Set EjecutarConsulta = resultado
End Function

Sub EjecutarNonQuery(conn, sql, parametros)
    On Error Resume Next
    Call EjecutarNonQueryTx(conn, sql, parametros)
    Call RedirigirSiErrorSQL(sql)
    On Error Goto 0
End Sub

Sub RedirigirSiErrorSQL(sql)
    If Err.Number <> 0 Then
        Dim numeroError, descripcionError
        numeroError = Err.Number
        descripcionError = Err.Description
        Response.AppendToLog " ToperBTR_ErrorSQL=" & numeroError & ":" & descripcionError & " SQL=" & Left(sql, 200)
        Response.Redirect "/errors/error.asp?msg=consulta"
        Response.End
    End If
End Sub

Function ObtenerUltimoId(conn)
    Dim rs
    Set rs = conn.Execute("SELECT LAST_INSERT_ID() AS ultimo_id")
    ObtenerUltimoId = CLng(rs("ultimo_id"))
    rs.Close
End Function
%>
