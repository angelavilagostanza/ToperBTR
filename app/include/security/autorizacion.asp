<%
Sub RequiereSesion()
    If Not (Session("Autenticado") = True) Then
        Response.Redirect "/modules/login/login.asp"
        Response.End
    End If
End Sub

' Gate por perfil, para los modulos que el legacy no controla via ACCION_PERMITIDA_REF
' (Usuarios y Bloqueos): en ToperBTR se comprueba siempre en servidor, nunca solo
' ocultando el enlace de menu (ver docs/04-matriz-permisos.md).
Sub RequierePerfil(perfilesPermitidos)
    Dim i, encontrado
    Call RequiereSesion()
    encontrado = False
    For i = 0 To UBound(perfilesPermitidos)
        If Session("CodPerfil") = perfilesPermitidos(i) Then
            encontrado = True
        End If
    Next
    If Not encontrado Then
        Response.Redirect "/errors/error.asp?msg=sinpermiso"
        Response.End
    End If
End Sub

Sub RequiereAccion(conn, codAccionRequerida)
    Call RequiereSesion()
    If Not TienePermiso(Session("AccionesPermitidas"), codAccionRequerida) Then
        Response.Redirect "/errors/error.asp?msg=sinpermiso"
        Response.End
    End If
End Sub
%>
