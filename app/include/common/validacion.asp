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
%>
