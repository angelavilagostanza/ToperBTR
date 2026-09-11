<%@ Language=VBScript %>
<%
Option Explicit

Dim accion
accion = Request.Form("accion")
%>

<html>
<head>
    <title>Prueba MySQL</title>
</head>
<body>

<h2>Conectando a BTR</h2>

<form method="post">
    <input type="submit" name="accion" value="Probar">
</form>

<%
If accion <> "" Then

    On Error Resume Next

    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")

    ' Ajustar según tu driver
    conn.Open "Driver={MySQL ODBC 8.0 Unicode Driver};" & _
              "Server=172.30.170.27;" & _
              "Port=3306" & _
              "Database=eir_test;" & _
              "User=toperbtr_app;" & _
              "Password=lsdfuoQ26!;" & _
              "Option=3;"

    If Err.Number <> 0 Then

        Response.Write "<p style='color:red'>Error de conexión: " & _
                       Server.HTMLEncode(Err.Description) & "</p>"

    Else

        Response.Write "<p style='color:green'>✅ Conexión OK</p>"

        Dim rs

        Set rs = conn.Execute( _
            "SELECT VERSION() AS version_mysql," & _
            " DATABASE() AS base_datos," & _
            " NOW() AS fecha_hora")

        If Not rs.EOF Then

            Response.Write "<b>Versión:</b> " & rs("version_mysql") & "<br>"
            Response.Write "<b>Base de datos:</b> " & rs("base_datos") & "<br>"
            Response.Write "<b>Fecha servidor:</b> " & rs("fecha_hora") & "<br>"

        End If

        rs.Close
        Set rs = Nothing

        conn.Close

    End If

    Set conn = Nothing

End If
%>

</body>
</html>