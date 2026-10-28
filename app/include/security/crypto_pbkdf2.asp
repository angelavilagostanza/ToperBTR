<%
' Implementacion 100% nativa (sin componentes COM/DLL externos) de SHA-1, SHA-256,
' HMAC-SHA256, PBKDF2-HMAC-SHA256 y Base64, en VBScript puro.
'
' Antes esto estaba en JScript (sus operadores de 32 bits nativos no tienen problemas
' de desbordamiento), pero se ha comprobado en el servidor de PRO (Windows Server 2025)
' que mezclar motores VBScript+JScript dentro de la misma pagina ASP falla de forma
' intermitente (~50% de las peticiones) con un crash nativo del motor de scripting de
' IIS (C0000005 en IActiveScript::SetScriptState/Close) - no es un bug de este codigo,
' es una incompatibilidad del propio servidor. La solucion de aplicacion es no depender
' nunca de JScript: todo el motor de hashing se reescribe aqui en VBScript, que no tiene
' operadores de bits de 32 sin signo nativos, asi que los words de 32 bits se representan
' como Double no negativos (un Double representa enteros exactos hasta 2^53, de sobra),
' la suma modulo 2^32 se hace por aritmetica (resta condicional, nunca con el operador
' Mod de VBScript sobre valores grandes - Mod internamente los convierte a Long y
' revienta con "Overflow" por encima de 2^31), y las operaciones a nivel de bit (And/Or/
' Xor/Not) se hacen convirtiendo el Double sin signo a Long con signo (misma huella de
' bits, complemento a dos) justo antes de usar los operadores nativos de VBScript, que si
' operan correctamente sobre los 32 bits de un Long.
'
' SHA-1 se mantiene unicamente para poder verificar (y migrar) los hashes del sistema
' legacy (SHA-1 + Base64, sin sal). Nunca se debe usar SHA-1 para generar hashes nuevos.
'
' Los valores de salida de esta funciones (hashes PBKDF2SHA256$.../SHA1+Base64) son
' exactamente los mismos que generaba la version en JScript: son implementaciones
' estandar (RFC 3174, FIPS 180-4, RFC 2898), no hay margen de interpretacion, asi que
' los hashes ya almacenados en BD por la version anterior siguen verificando igual.

Dim SHA256_K(63)
SHA256_K(0)=1116352408 : SHA256_K(1)=1899447441 : SHA256_K(2)=3049323471 : SHA256_K(3)=3921009573
SHA256_K(4)=961987163 : SHA256_K(5)=1508970993 : SHA256_K(6)=2453635748 : SHA256_K(7)=2870763221
SHA256_K(8)=3624381080 : SHA256_K(9)=310598401 : SHA256_K(10)=607225278 : SHA256_K(11)=1426881987
SHA256_K(12)=1925078388 : SHA256_K(13)=2162078206 : SHA256_K(14)=2614888103 : SHA256_K(15)=3248222580
SHA256_K(16)=3835390401 : SHA256_K(17)=4022224774 : SHA256_K(18)=264347078 : SHA256_K(19)=604807628
SHA256_K(20)=770255983 : SHA256_K(21)=1249150122 : SHA256_K(22)=1555081692 : SHA256_K(23)=1996064986
SHA256_K(24)=2554220882 : SHA256_K(25)=2821834349 : SHA256_K(26)=2952996808 : SHA256_K(27)=3210313671
SHA256_K(28)=3336571891 : SHA256_K(29)=3584528711 : SHA256_K(30)=113926993 : SHA256_K(31)=338241895
SHA256_K(32)=666307205 : SHA256_K(33)=773529912 : SHA256_K(34)=1294757372 : SHA256_K(35)=1396182291
SHA256_K(36)=1695183700 : SHA256_K(37)=1986661051 : SHA256_K(38)=2177026350 : SHA256_K(39)=2456956037
SHA256_K(40)=2730485921 : SHA256_K(41)=2820302411 : SHA256_K(42)=3259730800 : SHA256_K(43)=3345764771
SHA256_K(44)=3516065817 : SHA256_K(45)=3600352804 : SHA256_K(46)=4094571909 : SHA256_K(47)=275423344
SHA256_K(48)=430227734 : SHA256_K(49)=506948616 : SHA256_K(50)=659060556 : SHA256_K(51)=883997877
SHA256_K(52)=958139571 : SHA256_K(53)=1322822218 : SHA256_K(54)=1537002063 : SHA256_K(55)=1747873779
SHA256_K(56)=1955562222 : SHA256_K(57)=2024104815 : SHA256_K(58)=2227730452 : SHA256_K(59)=2361852424
SHA256_K(60)=2428436474 : SHA256_K(61)=2756734187 : SHA256_K(62)=3204031479 : SHA256_K(63)=3329325298

Const B64_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Const HEX_CHARS = "0123456789abcdef"

' ---------------------------------------------------------------------------
' Helpers de aritmetica/bits de 32 bits sin signo (word representado como
' Double no negativo, rango [0, 4294967295]).
' ---------------------------------------------------------------------------

' BAnd32/BOr32/BXor32/BNot32 convierten Double sin signo -> Long con signo (misma huella
' de bits, complemento a dos), aplican el operador nativo de VBScript (que si opera bien
' sobre los 32 bits de un Long) y deshacen la conversion. Antes esto se apoyaba en dos
' funciones auxiliares U2S/S2U; se ha dejado la conversion en linea (duplicada en las 4
' funciones) porque son, con diferencia, las funciones mas llamadas de todo el fichero
' (varios miles de veces por login con 200 iteraciones de PBKDF2) y cada llamada a
' funcion en VBScript tiene su propio coste fijo - eliminar 3 llamadas de por medio en
' cada una es la optimizacion de mayor impacto de todo el fichero. Se valido que el
' resultado es identico antes/despues con la misma bateria de vectores NIST/RFC + .NET.
Function BAnd32(ByVal a, ByVal b)
    Dim sa, sb, r
    If a >= 2147483648 Then sa = CLng(a - 4294967296) Else sa = CLng(a)
    If b >= 2147483648 Then sb = CLng(b - 4294967296) Else sb = CLng(b)
    r = sa And sb
    If r < 0 Then BAnd32 = CDbl(r) + 4294967296 Else BAnd32 = CDbl(r)
End Function

Function BOr32(ByVal a, ByVal b)
    Dim sa, sb, r
    If a >= 2147483648 Then sa = CLng(a - 4294967296) Else sa = CLng(a)
    If b >= 2147483648 Then sb = CLng(b - 4294967296) Else sb = CLng(b)
    r = sa Or sb
    If r < 0 Then BOr32 = CDbl(r) + 4294967296 Else BOr32 = CDbl(r)
End Function

Function BXor32(ByVal a, ByVal b)
    Dim sa, sb, r
    If a >= 2147483648 Then sa = CLng(a - 4294967296) Else sa = CLng(a)
    If b >= 2147483648 Then sb = CLng(b - 4294967296) Else sb = CLng(b)
    r = sa Xor sb
    If r < 0 Then BXor32 = CDbl(r) + 4294967296 Else BXor32 = CDbl(r)
End Function

Function BNot32(ByVal a)
    Dim sa, r
    If a >= 2147483648 Then sa = CLng(a - 4294967296) Else sa = CLng(a)
    r = Not sa
    If r < 0 Then BNot32 = CDbl(r) + 4294967296 Else BNot32 = CDbl(r)
End Function

' Suma modulo 2^32 (equivalente al add32(...) variadico del JScript original). VBScript
' clasico no soporta ParamArray (es solo de VBA/VB6), de ahi las 3 variantes de aridad
' fija usadas en el codigo. Nunca se usa el operador Mod de VBScript sobre valores
' grandes: la reduccion es por resta condicional, ya que tras cada suma parcial el
' resultado es siempre < 2^33.
Function Add32_2(ByVal a, ByVal b)
    Dim s
    s = a + b
    If s >= 4294967296 Then s = s - 4294967296
    Add32_2 = s
End Function

' Misma reduccion en linea que en los BAnd32/BOr32/... de arriba, por el mismo motivo:
' evitar 3-4 llamadas a funcion de por medio en una de las operaciones mas repetidas.
Function Add32_4(ByVal a, ByVal b, ByVal c, ByVal d)
    Dim s
    s = a + b
    If s >= 4294967296 Then s = s - 4294967296
    s = s + c
    If s >= 4294967296 Then s = s - 4294967296
    s = s + d
    If s >= 4294967296 Then s = s - 4294967296
    Add32_4 = s
End Function

Function Add32_5(ByVal a, ByVal b, ByVal c, ByVal d, ByVal e)
    Dim s
    s = a + b
    If s >= 4294967296 Then s = s - 4294967296
    s = s + c
    If s >= 4294967296 Then s = s - 4294967296
    s = s + d
    If s >= 4294967296 Then s = s - 4294967296
    s = s + e
    If s >= 4294967296 Then s = s - 4294967296
    Add32_5 = s
End Function

' Desplazamiento logico a la derecha. x/2^n es exacto en coma flotante (dividir por
' una potencia de 2 no redondea), por eso esto nunca pierde precision.
Function ShiftRight32(ByVal x, ByVal n)
    If n <= 0 Then
        ShiftRight32 = x
    Else
        ShiftRight32 = Int(x / (2 ^ n))
    End If
End Function

' Desplazamiento logico a la izquierda, resultado truncado a 32 bits. Se calcula
' quedandose primero solo con los bits de x que no se van a desbordar (lo = x mod
' 2^(32-n)) y desplazando esos - asi el producto final nunca supera 2^32 y no hay
' perdida de precision por trabajar con numeros mayores que 2^53.
Function ShiftLeft32(ByVal x, ByVal n)
    Dim divisor, lo
    If n <= 0 Then
        ShiftLeft32 = x
        Exit Function
    End If
    divisor = 2 ^ (32 - n)
    lo = x - Int(x / divisor) * divisor
    ShiftLeft32 = lo * (2 ^ n)
End Function

Function RotR32(ByVal x, ByVal n)
    RotR32 = BOr32(ShiftRight32(x, n), ShiftLeft32(x, 32 - n))
End Function

Function RotL32(ByVal x, ByVal n)
    RotL32 = BOr32(ShiftLeft32(x, n), ShiftRight32(x, 32 - n))
End Function

' Extrae un byte (8 bits) de un word de 32 bits empezando en shiftBits, por
' aritmetica pura (nunca con And sobre el word completo, que podria superar 2^31
' y hacer que el operador nativo de VBScript intente convertirlo a Long y reviente).
Function ExtractByte(ByVal word32, ByVal shiftBits)
    Dim shifted
    If shiftBits <= 0 Then
        shifted = word32
    Else
        shifted = Int(word32 / (2 ^ shiftBits))
    End If
    ExtractByte = shifted - Int(shifted / 256) * 256
End Function

' ---------------------------------------------------------------------------
' Utilidades de bytes: UTF-8, Base64, hex, concatenacion.
' ---------------------------------------------------------------------------

' Codifica una cadena VBScript (UTF-16 interno) a bytes UTF-8, incluyendo pares
' subrogados para puntos de codigo fuera del BMP. AscW puede devolver negativo
' para codigos >= &H8000 (quirk conocido de VBScript) - se corrige sumando 65536.
Function Utf8Bytes(ByVal str)
    Dim len1, i, code, nextCode, cp, totalBytes, pos
    len1 = Len(str)

    ' Los umbrales de par subrogado (D800/DBFF/DC00/DFFF) se escriben en decimal a
    ' proposito: un literal &H... que cabe en 16 bits pero tiene el bit alto activado
    ' (rango &H8000-&HFFFF) VBScript lo interpreta como Integer NEGATIVO (p.ej. &HD800
    ' vale -10240, no 55296), lo que rompia silenciosamente esta comparacion.
    totalBytes = 0
    i = 1
    Do While i <= len1
        code = AscW(Mid(str, i, 1))
        If code < 0 Then code = code + 65536
        If code < &H80 Then
            totalBytes = totalBytes + 1
            i = i + 1
        ElseIf code < &H800 Then
            totalBytes = totalBytes + 2
            i = i + 1
        ElseIf code >= 55296 And code <= 56319 And i < len1 Then  ' D800..DBFF: surrogate alto
            nextCode = AscW(Mid(str, i + 1, 1))
            If nextCode < 0 Then nextCode = nextCode + 65536
            If nextCode >= 56320 And nextCode <= 57343 Then       ' DC00..DFFF: surrogate bajo
                totalBytes = totalBytes + 4
                i = i + 2
            Else
                totalBytes = totalBytes + 3
                i = i + 1
            End If
        Else
            totalBytes = totalBytes + 3
            i = i + 1
        End If
    Loop

    If totalBytes = 0 Then
        Utf8Bytes = Array()
        Exit Function
    End If

    Dim bytes()
    ReDim bytes(totalBytes - 1)
    pos = 0
    i = 1
    Do While i <= len1
        code = AscW(Mid(str, i, 1))
        If code < 0 Then code = code + 65536
        If code < &H80 Then
            bytes(pos) = code : pos = pos + 1
            i = i + 1
        ElseIf code < &H800 Then
            bytes(pos) = &HC0 Or (code \ &H40) : pos = pos + 1
            bytes(pos) = &H80 Or (code And &H3F) : pos = pos + 1
            i = i + 1
        ElseIf code >= 55296 And code <= 56319 And i < len1 Then  ' D800..DBFF: surrogate alto
            nextCode = AscW(Mid(str, i + 1, 1))
            If nextCode < 0 Then nextCode = nextCode + 65536
            If nextCode >= 56320 And nextCode <= 57343 Then       ' DC00..DFFF: surrogate bajo
                cp = ((code - 55296) * &H400) + (nextCode - 56320) + &H10000
                bytes(pos) = &HF0 Or (cp \ &H40000) : pos = pos + 1
                bytes(pos) = &H80 Or ((cp \ &H1000) And &H3F) : pos = pos + 1
                bytes(pos) = &H80 Or ((cp \ &H40) And &H3F) : pos = pos + 1
                bytes(pos) = &H80 Or (cp And &H3F) : pos = pos + 1
                i = i + 2
            Else
                bytes(pos) = &HE0 Or (code \ &H1000) : pos = pos + 1
                bytes(pos) = &H80 Or ((code \ &H40) And &H3F) : pos = pos + 1
                bytes(pos) = &H80 Or (code And &H3F) : pos = pos + 1
                i = i + 1
            End If
        Else
            bytes(pos) = &HE0 Or (code \ &H1000) : pos = pos + 1
            bytes(pos) = &H80 Or ((code \ &H40) And &H3F) : pos = pos + 1
            bytes(pos) = &H80 Or (code And &H3F) : pos = pos + 1
            i = i + 1
        End If
    Loop

    Utf8Bytes = bytes
End Function

Function ConcatBytes(ByRef a, ByRef b)
    Dim na, nb, out, i
    na = UBound(a) + 1
    nb = UBound(b) + 1
    If na + nb = 0 Then
        ConcatBytes = Array()
        Exit Function
    End If
    Dim outArr()
    ReDim outArr(na + nb - 1)
    For i = 0 To na - 1
        outArr(i) = a(i)
    Next
    For i = 0 To nb - 1
        outArr(na + i) = b(i)
    Next
    ConcatBytes = outArr
End Function

Function BytesToBase64(ByRef bytes)
    Dim n, i, b0, b1, b2, triple, out
    n = UBound(bytes) + 1
    out = ""
    i = 0
    Do While i < n
        b0 = bytes(i)
        If i + 1 < n Then b1 = bytes(i + 1) Else b1 = 0
        If i + 2 < n Then b2 = bytes(i + 2) Else b2 = 0
        triple = b0 * 65536 + b1 * 256 + b2
        out = out & Mid(B64_CHARS, ((triple \ 262144) And 63) + 1, 1)
        out = out & Mid(B64_CHARS, ((triple \ 4096) And 63) + 1, 1)
        If i + 1 < n Then
            out = out & Mid(B64_CHARS, ((triple \ 64) And 63) + 1, 1)
        Else
            out = out & "="
        End If
        If i + 2 < n Then
            out = out & Mid(B64_CHARS, (triple And 63) + 1, 1)
        Else
            out = out & "="
        End If
        i = i + 3
    Loop
    BytesToBase64 = out
End Function

' Decodifica Base64 a bytes. El "buffer" acumulador se resetea a su resto tras cada
' byte extraido (en vez de crecer sin limite como en JS, donde el truncado a 32 bits
' de los operadores nativos lo hace gratis) - evita perder precision en cadenas largas.
Function Base64ToBytes(ByVal b64)
    Dim s, n, i, ch, val, buffer, bitsCollected, divisor, byteVal, outCount
    Dim tempBytes()

    s = b64
    Do While Len(s) > 0 And Right(s, 1) = "="
        s = Left(s, Len(s) - 1)
    Loop

    n = Len(s)
    If n = 0 Then
        Base64ToBytes = Array()
        Exit Function
    End If

    ReDim tempBytes(n)
    buffer = 0
    bitsCollected = 0
    outCount = 0

    For i = 1 To n
        ch = Mid(s, i, 1)
        val = InStr(1, B64_CHARS, ch, 0) - 1
        If val >= 0 Then
            buffer = buffer * 64 + val
            bitsCollected = bitsCollected + 6
            If bitsCollected >= 8 Then
                bitsCollected = bitsCollected - 8
                divisor = 2 ^ bitsCollected
                byteVal = Int(buffer / divisor)
                tempBytes(outCount) = byteVal
                outCount = outCount + 1
                buffer = buffer - byteVal * divisor
            End If
        End If
    Next

    If outCount = 0 Then
        Base64ToBytes = Array()
    Else
        ReDim Preserve tempBytes(outCount - 1)
        Base64ToBytes = tempBytes
    End If
End Function

Function ToHex(ByRef bytes)
    Dim n, i, s, b
    n = UBound(bytes) + 1
    s = ""
    For i = 0 To n - 1
        b = bytes(i)
        s = s & Mid(HEX_CHARS, (b \ 16) + 1, 1) & Mid(HEX_CHARS, (b Mod 16) + 1, 1)
    Next
    ToHex = s
End Function

' ---------------------------------------------------------------------------
' Relleno de mensaje (comun a SHA-1 y SHA-256: bloque de 64 bytes, 0x80, ceros,
' longitud en bits a 32 bits big-endian).
' ---------------------------------------------------------------------------

Function PadMensaje(ByRef msgBytes)
    Dim msgLen, bitLen, totalLen, padded, i

    msgLen = UBound(msgBytes) + 1
    totalLen = ((msgLen + 9 + 63) \ 64) * 64

    ReDim padded(totalLen - 1)
    For i = 0 To msgLen - 1
        padded(i) = msgBytes(i)
    Next
    padded(msgLen) = &H80
    ' El resto queda a 0 por defecto (ReDim inicializa los elementos numericos a 0).

    bitLen = msgLen * 8
    padded(totalLen - 4) = ExtractByte(bitLen, 24)
    padded(totalLen - 3) = ExtractByte(bitLen, 16)
    padded(totalLen - 2) = ExtractByte(bitLen, 8)
    padded(totalLen - 1) = ExtractByte(bitLen, 0)

    PadMensaje = padded
End Function

' ---------------------------------------------------------------------------
' SHA-256 / SHA-1
' ---------------------------------------------------------------------------

Function Sha256Bytes(ByRef msgBytes)
    Dim h0, h1, h2, h3, h4, h5, h6, h7
    h0 = 1779033703 : h1 = 3144134277 : h2 = 1013904242 : h3 = 2773480762
    h4 = 1359893119 : h5 = 2600822924 : h6 = 528734635 : h7 = 1541459225

    Dim padded, numChunks, chunk
    padded = PadMensaje(msgBytes)
    numChunks = (UBound(padded) + 1) \ 64

    Dim w(63)
    Dim i, j, a, b, c, d, e, f, g, hh, T1, T2, sig0, sig1, bigSig0, bigSig1, ch, maj

    For chunk = 0 To numChunks - 1
        For i = 0 To 15
            j = chunk * 64 + i * 4
            ' Cada byte (0-255) ocupa un carril de 8 bits disjunto una vez colocado en su
            ' posicion (multiplo de 256) - sumarlos por aritmetica equivale exactamente a
            ' desplazar+OR (no puede haber acarreo entre carriles), y evita 3 llamadas a
            ' ShiftLeft32/BOr32 por cada word de entrada.
            w(i) = padded(j) * 16777216 + padded(j + 1) * 65536 + padded(j + 2) * 256 + padded(j + 3)
        Next
        For i = 16 To 63
            sig0 = BXor32(BXor32(RotR32(w(i - 15), 7), RotR32(w(i - 15), 18)), ShiftRight32(w(i - 15), 3))
            sig1 = BXor32(BXor32(RotR32(w(i - 2), 17), RotR32(w(i - 2), 19)), ShiftRight32(w(i - 2), 10))
            w(i) = Add32_4(w(i - 16), sig0, w(i - 7), sig1)
        Next

        a = h0 : b = h1 : c = h2 : d = h3 : e = h4 : f = h5 : g = h6 : hh = h7

        For i = 0 To 63
            bigSig1 = BXor32(BXor32(RotR32(e, 6), RotR32(e, 11)), RotR32(e, 25))
            ch = BXor32(BAnd32(e, f), BAnd32(BNot32(e), g))
            T1 = Add32_5(hh, bigSig1, ch, SHA256_K(i), w(i))
            bigSig0 = BXor32(BXor32(RotR32(a, 2), RotR32(a, 13)), RotR32(a, 22))
            maj = BXor32(BXor32(BAnd32(a, b), BAnd32(a, c)), BAnd32(b, c))
            T2 = Add32_2(bigSig0, maj)
            hh = g : g = f : f = e : e = Add32_2(d, T1) : d = c : c = b : b = a : a = Add32_2(T1, T2)
        Next

        h0 = Add32_2(h0, a) : h1 = Add32_2(h1, b) : h2 = Add32_2(h2, c) : h3 = Add32_2(h3, d)
        h4 = Add32_2(h4, e) : h5 = Add32_2(h5, f) : h6 = Add32_2(h6, g) : h7 = Add32_2(h7, hh)
    Next

    Dim outb(31), words(7)
    words(0) = h0 : words(1) = h1 : words(2) = h2 : words(3) = h3
    words(4) = h4 : words(5) = h5 : words(6) = h6 : words(7) = h7
    For i = 0 To 7
        outb(i * 4)     = ExtractByte(words(i), 24)
        outb(i * 4 + 1) = ExtractByte(words(i), 16)
        outb(i * 4 + 2) = ExtractByte(words(i), 8)
        outb(i * 4 + 3) = ExtractByte(words(i), 0)
    Next
    Sha256Bytes = outb
End Function

Function Sha1Bytes(ByRef msgBytes)
    Dim h0, h1, h2, h3, h4
    h0 = 1732584193 : h1 = 4023233417 : h2 = 2562383102 : h3 = 271733878 : h4 = 3285377520

    Dim padded, numChunks, chunk
    padded = PadMensaje(msgBytes)
    numChunks = (UBound(padded) + 1) \ 64

    Dim w(79)
    Dim i, j, a, b, c, d, e, f, k, temp

    For chunk = 0 To numChunks - 1
        For i = 0 To 15
            j = chunk * 64 + i * 4
            ' Cada byte (0-255) ocupa un carril de 8 bits disjunto una vez colocado en su
            ' posicion (multiplo de 256) - sumarlos por aritmetica equivale exactamente a
            ' desplazar+OR (no puede haber acarreo entre carriles), y evita 3 llamadas a
            ' ShiftLeft32/BOr32 por cada word de entrada.
            w(i) = padded(j) * 16777216 + padded(j + 1) * 65536 + padded(j + 2) * 256 + padded(j + 3)
        Next
        For i = 16 To 79
            w(i) = RotL32(BXor32(BXor32(BXor32(w(i - 3), w(i - 8)), w(i - 14)), w(i - 16)), 1)
        Next

        a = h0 : b = h1 : c = h2 : d = h3 : e = h4

        For i = 0 To 79
            If i < 20 Then
                f = BOr32(BAnd32(b, c), BAnd32(BNot32(b), d))
                k = 1518500249
            ElseIf i < 40 Then
                f = BXor32(BXor32(b, c), d)
                k = 1859775393
            ElseIf i < 60 Then
                f = BOr32(BOr32(BAnd32(b, c), BAnd32(b, d)), BAnd32(c, d))
                k = 2400959708
            Else
                f = BXor32(BXor32(b, c), d)
                k = 3395469782
            End If
            temp = Add32_5(RotL32(a, 5), f, e, k, w(i))
            e = d : d = c : c = RotL32(b, 30) : b = a : a = temp
        Next

        h0 = Add32_2(h0, a) : h1 = Add32_2(h1, b) : h2 = Add32_2(h2, c) : h3 = Add32_2(h3, d) : h4 = Add32_2(h4, e)
    Next

    Dim outb(19), words(4)
    words(0) = h0 : words(1) = h1 : words(2) = h2 : words(3) = h3 : words(4) = h4
    For i = 0 To 4
        outb(i * 4)     = ExtractByte(words(i), 24)
        outb(i * 4 + 1) = ExtractByte(words(i), 16)
        outb(i * 4 + 2) = ExtractByte(words(i), 8)
        outb(i * 4 + 3) = ExtractByte(words(i), 0)
    Next
    Sha1Bytes = outb
End Function

' ---------------------------------------------------------------------------
' HMAC-SHA256 / PBKDF2-HMAC-SHA256
' ---------------------------------------------------------------------------

Function HmacSha256Bytes(ByRef keyBytes, ByRef msgBytes)
    Dim blockSize, key, keyLen, i, kb
    Dim oKeyPad(63), iKeyPad(63), inner

    blockSize = 64
    keyLen = UBound(keyBytes) + 1
    If keyLen > blockSize Then
        key = Sha256Bytes(keyBytes)
        keyLen = UBound(key) + 1
    Else
        key = keyBytes
    End If

    For i = 0 To blockSize - 1
        If i < keyLen Then kb = key(i) Else kb = 0
        oKeyPad(i) = kb Xor &H5C
        iKeyPad(i) = kb Xor &H36
    Next

    inner = Sha256Bytes(ConcatBytes(iKeyPad, msgBytes))
    HmacSha256Bytes = Sha256Bytes(ConcatBytes(oKeyPad, inner))
End Function

Function Pbkdf2Block(ByRef passwordBytes, ByRef saltBytes, ByVal iterations, ByVal blockIndex)
    Dim idxBytes(3), saltPlusIndex, u, tArr(31), i, j, tn

    idxBytes(0) = ExtractByte(blockIndex, 24)
    idxBytes(1) = ExtractByte(blockIndex, 16)
    idxBytes(2) = ExtractByte(blockIndex, 8)
    idxBytes(3) = ExtractByte(blockIndex, 0)
    saltPlusIndex = ConcatBytes(saltBytes, idxBytes)

    u = HmacSha256Bytes(passwordBytes, saltPlusIndex)
    tn = UBound(u) + 1
    For i = 0 To tn - 1
        tArr(i) = u(i)
    Next

    For i = 1 To iterations - 1
        u = HmacSha256Bytes(passwordBytes, u)
        For j = 0 To tn - 1
            tArr(j) = tArr(j) Xor u(j)
        Next
    Next

    Pbkdf2Block = tArr
End Function

Function Pbkdf2Sha256Bytes(ByRef passwordBytes, ByRef saltBytes, ByVal iterations, ByVal dkLenBytes)
    Dim hLen, l, i, k, t, tn, pos, dkArr, outArr

    hLen = 32
    l = Int((dkLenBytes + hLen - 1) / hLen)

    ReDim dkArr(l * hLen - 1)
    pos = 0
    For i = 1 To l
        t = Pbkdf2Block(passwordBytes, saltBytes, iterations, i)
        tn = UBound(t) + 1
        For k = 0 To tn - 1
            dkArr(pos) = t(k)
            pos = pos + 1
        Next
    Next

    ReDim outArr(dkLenBytes - 1)
    For i = 0 To dkLenBytes - 1
        outArr(i) = dkArr(i)
    Next
    Pbkdf2Sha256Bytes = outArr
End Function

' Comparacion en tiempo constante (evita filtrar por temporizacion cuanto coincide un hash).
Function ComparacionSegura(ByVal a, ByVal b)
    Dim diff, i, n
    If Len(a) <> Len(b) Then
        ComparacionSegura = False
        Exit Function
    End If
    diff = 0
    n = Len(a)
    For i = 1 To n
        diff = diff Or (AscW(Mid(a, i, 1)) Xor AscW(Mid(b, i, 1)))
    Next
    ComparacionSegura = (diff = 0)
End Function

' Hallazgo de seguridad M3: antes la aleatoriedad salia de Rnd() de VBScript (un PRNG
' predecible, no criptografico), mezclando Now()/Timer()/Rnd() y pasandolo por SHA-256.
' Esa mitigacion tiene un limite real: pasar una entrada de poca entropia por una funcion
' hash no anade entropia que no estuviera ya ahi, solo la revuelve - si alguien acota la
' hora del servidor (la cabecera HTTP Date se la da gratis) y el rango de Timer(), el
' espacio a probar para reproducir un token deja de ser astronomico.
'
' Se sustituye por Scriptlet.TypeLib.Guid: genera un GUID v4 real (formato de Windows,
' CoCreateGuid) que usa el generador criptografico del sistema operativo por debajo
' (~122 bits de entropia real por GUID, los otros 6 bits son fijos de version/variante).
' Verificado en este proyecto que SI funciona desde VBScript puro (a diferencia de
' CAPICOM.Utilities, no disponible en el servidor, y de
' System.Security.Cryptography.RNGCryptoServiceProvider via COM, cuyo objeto se crea
' pero el metodo GetBytes falla siempre desde VBScript con error 5 "argumento o llamada
' a procedimiento no validos" - el mismo patron de "parece disponible pero no funciona
' de verdad" que Response.Cookies.HttpOnly en el hallazgo A5).
Function GuidAleatorio()
    Dim tl
    Set tl = CreateObject("Scriptlet.TypeLib")
    GuidAleatorio = tl.Guid
End Function

' 2 GUIDs (no 1) por margen: aunque un GUID ya tiene entropia de sobra, concatenar dos
' no cuesta nada y deja margen frente a cualquier duda futura sobre la implementacion
' exacta de CoCreateGuid en una version de Windows distinta.
Function EntropiaFuerte()
    EntropiaFuerte = Utf8Bytes(GuidAleatorio() & GuidAleatorio())
End Function

' Expande la entropia de EntropiaFuerte() a exactamente n bytes mediante SHA-256 en modo
' contador (bloque_i = SHA256(entropia & i), concatenados y truncados a n) - una tecnica
' estandar de expansion de claves, necesaria porque SHA-256 solo da 32 bytes por llamada
' y aqui se piden desde 6 bytes (placeholder de solicitud) hasta 32 (CSRF/huella de
' sesion/sal de password).
Function BytesAleatorios(ByVal n)
    Dim entropia, outArr(), pos, contador, bloque, k
    If n <= 0 Then
        BytesAleatorios = Array()
        Exit Function
    End If

    entropia = EntropiaFuerte()
    ReDim outArr(n - 1)
    pos = 0
    contador = 0
    Do While pos < n
        contador = contador + 1
        bloque = Sha256Bytes(ConcatBytes(entropia, Utf8Bytes(CStr(contador))))
        For k = 0 To UBound(bloque)
            If pos >= n Then Exit For
            outArr(pos) = bloque(k)
            pos = pos + 1
        Next
    Loop
    BytesAleatorios = outArr
End Function

Function GenerarTokenAleatorio(ByVal numBytes)
    GenerarTokenAleatorio = ToHex(BytesAleatorios(numBytes))
End Function

Function CalcularSha1Base64(ByVal password)
    CalcularSha1Base64 = BytesToBase64(Sha1Bytes(Utf8Bytes(password)))
End Function

Function CrearHashPbkdf2(ByVal password, ByVal iteraciones)
    Dim saltBytes, dk
    saltBytes = BytesAleatorios(16)
    dk = Pbkdf2Sha256Bytes(Utf8Bytes(password), saltBytes, iteraciones, 32)
    CrearHashPbkdf2 = "PBKDF2SHA256$" & iteraciones & "$" & BytesToBase64(saltBytes) & "$" & BytesToBase64(dk)
End Function

Function VerificarHashPbkdf2(ByVal password, ByVal valorAlmacenado)
    Dim partes, iteraciones, saltBytes, dk

    partes = Split(CStr(valorAlmacenado), "$")
    If UBound(partes) <> 3 Then
        VerificarHashPbkdf2 = False
        Exit Function
    End If
    If partes(0) <> "PBKDF2SHA256" Then
        VerificarHashPbkdf2 = False
        Exit Function
    End If
    If Not IsNumeric(partes(1)) Then
        VerificarHashPbkdf2 = False
        Exit Function
    End If

    iteraciones = CLng(partes(1))
    saltBytes = Base64ToBytes(partes(2))
    dk = Pbkdf2Sha256Bytes(Utf8Bytes(password), saltBytes, iteraciones, 32)
    VerificarHashPbkdf2 = ComparacionSegura(BytesToBase64(dk), partes(3))
End Function

' Punto de entrada unico para verificar login: decide el algoritmo segun la columna
' usuarios.algoritmo_password, para que legacy (SHA1B64) y ToperBTR (PBKDF2SHA256) convivan.
Function VerificarPassword(ByVal password, ByVal hashAlmacenado, ByVal algoritmo)
    If algoritmo = "PBKDF2SHA256" Then
        VerificarPassword = VerificarHashPbkdf2(password, hashAlmacenado)
    Else
        VerificarPassword = ComparacionSegura(CalcularSha1Base64(password), CStr(hashAlmacenado))
    End If
End Function
%>
