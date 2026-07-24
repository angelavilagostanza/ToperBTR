<%
' Registro de acciones sensibles: usuario, accion, entidad, fecha, IP, resultado.
' Nunca se registra aqui contrasenas ni hashes, solo el hecho realizado.
' Se envuelve en On Error Resume Next para que un fallo al auditar (p.ej. tabla
' bloqueada un instante) nunca bloquee la operacion de negocio real que se esta auditando -
' por eso usa EjecutarNonQueryTx (la variante que NO redirige sola por error): si usara la
' variante segura por defecto, un fallo aqui redirigiria a la pagina de error incluso
' cuando la operacion real (p.ej. un alta de usuario) ya se completo correctamente.
Sub RegistrarAuditoria(conn, codIndiceUsuario, accion, entidad, entidadId, resultadoTexto, detalle)
    Dim ip
    ip = Request.ServerVariables("REMOTE_ADDR")
    On Error Resume Next
    Call EjecutarNonQueryTx(conn, _
        "INSERT INTO AUDITORIA_ACCION (COD_INDICE_USUARIO, ACCION, ENTIDAD, ENTIDAD_ID, FECHA, IP, RESULTADO, DETALLE) VALUES (?, ?, ?, ?, NOW(), ?, ?, ?)", _
        Array(codIndiceUsuario, accion, entidad, entidadId, ip, resultadoTexto, detalle))
    On Error Goto 0
End Sub
%>
