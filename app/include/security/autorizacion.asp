<%
' Ademas de Session("Autenticado"), exige que la cookie ToperBTR_Fp (HttpOnly, fijada en
' login_do.asp) coincida con Session("Fingerprint"). Esto es lo que de verdad impide que
' un ASPSESSIONID robado via XSS sirva para suplantar la sesion: el atacante puede copiar
' el ASPSESSIONID (no tiene HttpOnly, limitacion de ASP clasico), pero nunca pudo leer el
' valor de ToperBTR_Fp con JavaScript, asi que no puede reenviarlo. Una sesion sin huella
' valida (incluida cualquier sesion ya abierta antes de este cambio) se trata como no
' autenticada y se fuerza a volver a iniciar sesion.
Sub RequiereSesion()
    Dim fpCookie, fpSesion
    fpCookie = "" & Request.Cookies("ToperBTR_Fp")
    fpSesion = "" & Session("Fingerprint")
    If Not (Session("Autenticado") = True) Or fpSesion = "" Or fpCookie <> fpSesion Then
        Session.Abandon
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
