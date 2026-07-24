<%
Function MensajeErrorGenerico()
    MensajeErrorGenerico = "Se ha producido un error. Intentelo de nuevo mas tarde."
End Function

' Deja un aviso de una sola vez para la franja superior (ver include/layout/header.asp).
' tipo: "exito" | "error" | "info"
Sub EstablecerMensajeFlash(tipo, texto)
    Session("FlashTipo") = tipo
    Session("FlashTexto") = texto
End Sub
%>
