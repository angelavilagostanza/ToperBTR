<%
' Mensaje flash: cualquier controlador puede dejar un aviso en sesion (EstablecerMensajeFlash,
' en common/errores.asp) antes de redirigir; se muestra una unica vez aqui y se limpia.
Dim flashTexto, flashTipo, claseFlash
flashTexto = Session("FlashTexto")
flashTipo = Session("FlashTipo")
Session("FlashTexto") = ""
Session("FlashTipo") = ""

Select Case flashTipo
    Case "exito"
        claseFlash = "flash-exito"
    Case "error"
        claseFlash = "flash-error"
    Case Else
        claseFlash = "flash-info"
End Select
%>
<header class="cabecera">
    <div class="cabecera-marca">
        <img src="/images/ToperBTR_Logo_peque_v3.png" alt="ToperBTR" class="cabecera-logo">
        <span class="cabecera-tagline">Bloqueo de Terminales Robados</span>
    </div>
    <div class="cabecera-mensaje">
        <% If EsCadenaNoVacia(flashTexto) Then %>
        <span class="<%= claseFlash %>"><%= Server.HTMLEncode(flashTexto) %></span>
        <% End If %>
    </div>
    <% If Session("Autenticado") = True Then %>
    <div class="cabecera-cuenta">
        <span class="cabecera-usuario"><%= Server.HTMLEncode(Session("CodUsuario")) %></span>
        <div class="cabecera-cuenta-sep" aria-hidden="true"></div>
        <a href="/modules/usuarios/cambio_password.asp" class="cabecera-icono" title="Cambiar contraseña">
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <circle cx="8" cy="15" r="4"/><path d="M12 11l7-7m-3 0l3 3"/>
            </svg>
            <span class="cabecera-icono-label">Contraseña</span>
        </a>
        <a href="/modules/login/logout.asp" class="cabecera-icono" title="Cerrar sesión">
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/>
            </svg>
            <span class="cabecera-icono-label">Salir</span>
        </a>
    </div>
    <% End If %>
</header>
