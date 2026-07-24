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
    <div class="cabecera-marca"><%= Application("NombreApp") %></div>
    <div class="cabecera-mensaje">
        <% If EsCadenaNoVacia(flashTexto) Then %>
        <span class="<%= claseFlash %>"><%= Server.HTMLEncode(flashTexto) %></span>
        <% End If %>
    </div>
    <% If Session("Autenticado") = True Then %>
    <div class="cabecera-cuenta">
        <span class="cabecera-usuario"><%= Server.HTMLEncode(Session("CodUsuario")) %></span>
        <a href="/modules/usuarios/cambio_password.asp">Cambiar contrasena</a>
        <a href="/modules/login/logout.asp">Desconectar</a>
    </div>
    <% End If %>
</header>
