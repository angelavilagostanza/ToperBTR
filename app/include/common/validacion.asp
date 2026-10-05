<%
Function EsCadenaNoVacia(valor)
    EsCadenaNoVacia = (Trim(valor & "") <> "")
End Function

Function LongitudEntre(valor, minimo, maximo)
    Dim longitud
    longitud = Len(valor & "")
    LongitudEntre = (longitud >= minimo And longitud <= maximo)
End Function

Function CumpleFormato(valor, patron)
    Dim re
    Set re = New RegExp
    re.Pattern = patron
    re.IgnoreCase = False
    CumpleFormato = re.Test(valor & "")
End Function

' VBScript no tiene IIf (es de VBA/VB6, no de VBScript) - este es el equivalente propio.
Function SiVerdadero(condicion, valorSiTrue, valorSiFalse)
    If condicion Then
        SiVerdadero = valorSiTrue
    Else
        SiVerdadero = valorSiFalse
    End If
End Function

' Solo para los paneles "Ver consulta SQL generada" de los buscadores (bloqueo.asp,
' consulta.asp, buscar.asp de Incidencias/Ficheros). Los parametros de fecha llegan ya
' convertidos con CDate() (VarType 7 = vbDate), por eso EjecutarConsulta los enlaza como
' adDate (ver TipoAdoParaValor en conexion.asp) y no como texto: MySQL recibe un valor de
' fecha/hora real, no una cadena con el formato regional del servidor. Pero "valor & ''"
' SI usa el formato regional del servidor para un Date - esta funcion evita esa confusion
' mostrando siempre ISO en el panel, independientemente del locale de IIS.
Function TextoDebugParametro(valor)
    If VarType(valor) = 7 Then ' vbDate
        TextoDebugParametro = Year(valor) & "-" & Right("0" & Month(valor), 2) & "-" & Right("0" & Day(valor), 2)
        If Hour(valor) <> 0 Or Minute(valor) <> 0 Or Second(valor) <> 0 Then
            TextoDebugParametro = TextoDebugParametro & " " & Right("0" & Hour(valor), 2) & ":" & _
                Right("0" & Minute(valor), 2) & ":" & Right("0" & Second(valor), 2)
        End If
    Else
        TextoDebugParametro = valor & ""
    End If
End Function
%>
