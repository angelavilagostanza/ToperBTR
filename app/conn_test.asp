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

<h2>Diagnostico PKs sin AUTO_INCREMENT</h2>

<form method="post">
    <input type="submit" name="accion" value="Probar">
</form>

<%
If accion <> "" Then

    On Error Resume Next

    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open Application("DBConnectionString")

    If Err.Number <> 0 Then
        Response.Write "<p style='color:red'>Error de conexion: " & Server.HTMLEncode(Err.Description) & "</p>"
        Response.End
    End If

    Dim tablas(7)
    tablas(0) = "COMENTARIO_INCIDENCIA"
    tablas(1) = "AUDITORIA_ACCION"
    tablas(2) = "SOLICITUD"
    tablas(3) = "HISTORICO_ESTADO_SOLICITUD"
    tablas(4) = "HISTORICO_ESTADO_INCIDENCIA"
    tablas(5) = "BLOQUEOS"
    tablas(6) = "USUARIOS"
    tablas(7) = "INCIDENCIA"

    Dim i, rs, pkCols, extraCols
    For i = 0 To UBound(tablas)
        Err.Clear
        Set rs = conn.Execute("DESCRIBE " & tablas(i))
        If Err.Number <> 0 Then
            Response.Write "<h3>" & tablas(i) & " — ERROR: " & Server.HTMLEncode(Err.Description) & "</h3>"
        Else
            Response.Write "<h3>" & tablas(i) & "</h3>"
            Response.Write "<table border=1 cellpadding=3><tr><th>Field</th><th>Type</th><th>Null</th><th>Key</th><th>Default</th><th>Extra</th></tr>"
            pkCols = ""
            extraCols = ""
            Do While Not rs.EOF
                Dim isKey
                isKey = (rs(3) & "" = "PRI")
                If isKey Then
                    pkCols = pkCols & rs(0) & " "
                    extraCols = extraCols & (rs(5) & "")
                End If
                Dim rowColor
                rowColor = ""
                If isKey And InStr(LCase(rs(5) & ""), "auto_increment") = 0 Then
                    rowColor = " bgcolor='#ffcccc'"
                End If
                Response.Write "<tr" & rowColor & ">"
                Response.Write "<td>" & Server.HTMLEncode(rs(0) & "") & "</td>"
                Response.Write "<td>" & Server.HTMLEncode(rs(1) & "") & "</td>"
                Response.Write "<td>" & Server.HTMLEncode(rs(2) & "") & "</td>"
                Response.Write "<td>" & Server.HTMLEncode(rs(3) & "") & "</td>"
                Response.Write "<td>" & Server.HTMLEncode(rs(4) & "") & "</td>"
                Response.Write "<td>" & Server.HTMLEncode(rs(5) & "") & "</td>"
                Response.Write "</tr>"
                rs.MoveNext
            Loop
            rs.Close
            Response.Write "</table>"
            If pkCols <> "" And InStr(LCase(extraCols), "auto_increment") = 0 Then
                Response.Write "<p style='color:red'><b>PROBLEMA:</b> PK [" & Trim(pkCols) & "] sin AUTO_INCREMENT &mdash; ALTER TABLE " & tablas(i) & " MODIFY " & Trim(pkCols) & " INT NOT NULL AUTO_INCREMENT;</p>"
            Else
                Response.Write "<p style='color:green'>PK OK (AUTO_INCREMENT o sin PK simple)</p>"
            End If
        End If
        Err.Clear
    Next

    conn.Close
    Set conn = Nothing

End If
%>

</body>
</html>
