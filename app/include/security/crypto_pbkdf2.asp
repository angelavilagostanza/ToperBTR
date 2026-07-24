<script language="JScript" runat="server">
// Implementacion 100% nativa (sin componentes COM/DLL externos) de SHA-1, SHA-256,
// HMAC-SHA256, PBKDF2-HMAC-SHA256 y Base64, en JScript embebido dentro de una pagina
// ASP clasica. Se usa JScript (no VBScript) porque sus operadores de 32 bits (&,|,^,~,
// <<,>>>) no tienen los problemas de desbordamiento de VBScript. Las funciones definidas
// aqui quedan disponibles para el codigo VBScript del resto de la aplicacion, ya que
// todos los bloques de script de servidor de una misma pagina ASP comparten namespace
// (aviso: no escribir aqui la etiqueta literal de apertura de script de servidor ni
// siquiera en un comentario, el parser de ASP la detecta a nivel de texto plano e
// ignora que este dentro de un comentario JScript - eso es justo lo que causo el
// error "ASP 0138 Nested Script Block" la primera vez).
//
// SHA-1 se mantiene unicamente para poder verificar (y migrar) los hashes del sistema
// legacy (SHA-1 + Base64, sin sal). Nunca se debe usar SHA-1 para generar hashes nuevos.

var SHA256_K = [
 0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,
 0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,
 0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,
 0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,
 0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,
 0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,
 0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,
 0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2
];

var B64_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

function rotr32(x, n) { return ((x >>> n) | (x << (32 - n))) >>> 0; }
function rotl32(x, n) { return ((x << n) | (x >>> (32 - n))) >>> 0; }

function add32() {
    var s = 0, i;
    for (i = 0; i < arguments.length; i++) { s = (s + arguments[i]) >>> 0; }
    return s;
}

function utf8Bytes(str) {
    var bytes = [], i, code, next, cp;
    for (i = 0; i < str.length; i++) {
        code = str.charCodeAt(i);
        if (code < 0x80) {
            bytes.push(code);
        } else if (code < 0x800) {
            bytes.push(0xC0 | (code >> 6));
            bytes.push(0x80 | (code & 0x3F));
        } else if (code >= 0xD800 && code <= 0xDBFF && i + 1 < str.length) {
            next = str.charCodeAt(i + 1);
            if (next >= 0xDC00 && next <= 0xDFFF) {
                cp = ((code - 0xD800) << 10) + (next - 0xDC00) + 0x10000;
                i++;
                bytes.push(0xF0 | (cp >> 18));
                bytes.push(0x80 | ((cp >> 12) & 0x3F));
                bytes.push(0x80 | ((cp >> 6) & 0x3F));
                bytes.push(0x80 | (cp & 0x3F));
            } else {
                bytes.push(0xE0 | (code >> 12));
                bytes.push(0x80 | ((code >> 6) & 0x3F));
                bytes.push(0x80 | (code & 0x3F));
            }
        } else {
            bytes.push(0xE0 | (code >> 12));
            bytes.push(0x80 | ((code >> 6) & 0x3F));
            bytes.push(0x80 | (code & 0x3F));
        }
    }
    return bytes;
}

function bytesToBase64(bytes) {
    var out = "", i, b0, b1, b2, triple;
    for (i = 0; i < bytes.length; i += 3) {
        b0 = bytes[i];
        b1 = (i + 1 < bytes.length) ? bytes[i + 1] : 0;
        b2 = (i + 2 < bytes.length) ? bytes[i + 2] : 0;
        triple = (b0 << 16) | (b1 << 8) | b2;
        out += B64_CHARS.charAt((triple >>> 18) & 0x3F);
        out += B64_CHARS.charAt((triple >>> 12) & 0x3F);
        out += (i + 1 < bytes.length) ? B64_CHARS.charAt((triple >>> 6) & 0x3F) : "=";
        out += (i + 2 < bytes.length) ? B64_CHARS.charAt(triple & 0x3F) : "=";
    }
    return out;
}

function base64ToBytes(b64) {
    var s = String(b64).replace(/=+$/, "");
    var bytes = [], buffer = 0, bitsCollected = 0, i, val;
    for (i = 0; i < s.length; i++) {
        val = B64_CHARS.indexOf(s.charAt(i));
        if (val < 0) { continue; }
        buffer = (buffer << 6) | val;
        bitsCollected += 6;
        if (bitsCollected >= 8) {
            bitsCollected -= 8;
            bytes.push((buffer >>> bitsCollected) & 0xFF);
        }
    }
    return bytes;
}

function toHex(bytes) {
    var hex = "0123456789abcdef", s = "", i;
    for (i = 0; i < bytes.length; i++) {
        s += hex.charAt((bytes[i] >> 4) & 0xF) + hex.charAt(bytes[i] & 0xF);
    }
    return s;
}

function padMensaje(msgBytes) {
    var bitLen = (msgBytes.length * 8) >>> 0;
    var padded = msgBytes.slice();
    padded.push(0x80);
    while ((padded.length % 64) !== 56) { padded.push(0); }
    padded.push(0, 0, 0, 0);
    padded.push((bitLen >>> 24) & 0xFF, (bitLen >>> 16) & 0xFF, (bitLen >>> 8) & 0xFF, bitLen & 0xFF);
    return padded;
}

function sha256Bytes(msgBytes) {
    var h0=0x6a09e667,h1=0xbb67ae85,h2=0x3c6ef372,h3=0xa54ff53a,
        h4=0x510e527f,h5=0x9b05688c,h6=0x1f83d9ab,h7=0x5be0cd19;

    var padded = padMensaje(msgBytes);
    var numChunks = padded.length / 64;
    var chunk, w, i, j, a,b,c,d,e,f,g,hh,T1,T2,s0,s1,S0,S1,ch,maj;

    for (chunk = 0; chunk < numChunks; chunk++) {
        w = new Array(64);
        for (i = 0; i < 16; i++) {
            j = chunk * 64 + i * 4;
            w[i] = ((padded[j] << 24) | (padded[j+1] << 16) | (padded[j+2] << 8) | padded[j+3]) >>> 0;
        }
        for (i = 16; i < 64; i++) {
            s0 = rotr32(w[i-15], 7) ^ rotr32(w[i-15], 18) ^ (w[i-15] >>> 3);
            s1 = rotr32(w[i-2], 17) ^ rotr32(w[i-2], 19) ^ (w[i-2] >>> 10);
            w[i] = add32(w[i-16], s0, w[i-7], s1);
        }

        a=h0; b=h1; c=h2; d=h3; e=h4; f=h5; g=h6; hh=h7;

        for (i = 0; i < 64; i++) {
            S1 = rotr32(e,6) ^ rotr32(e,11) ^ rotr32(e,25);
            ch = (e & f) ^ ((~e) & g);
            T1 = add32(hh, S1, ch, SHA256_K[i], w[i]);
            S0 = rotr32(a,2) ^ rotr32(a,13) ^ rotr32(a,22);
            maj = (a & b) ^ (a & c) ^ (b & c);
            T2 = add32(S0, maj);
            hh=g; g=f; f=e; e=add32(d,T1); d=c; c=b; b=a; a=add32(T1,T2);
        }

        h0=add32(h0,a); h1=add32(h1,b); h2=add32(h2,c); h3=add32(h3,d);
        h4=add32(h4,e); h5=add32(h5,f); h6=add32(h6,g); h7=add32(h7,hh);
    }

    var out = [], words = [h0,h1,h2,h3,h4,h5,h6,h7];
    for (i = 0; i < 8; i++) {
        out.push((words[i]>>>24)&0xFF,(words[i]>>>16)&0xFF,(words[i]>>>8)&0xFF,words[i]&0xFF);
    }
    return out;
}

function sha1Bytes(msgBytes) {
    var h0=0x67452301,h1=0xEFCDAB89,h2=0x98BADCFE,h3=0x10325476,h4=0xC3D2E1F0;

    var padded = padMensaje(msgBytes);
    var numChunks = padded.length / 64;
    var chunk, w, i, j, a,b,c,d,e,f,k,temp;

    for (chunk = 0; chunk < numChunks; chunk++) {
        w = new Array(80);
        for (i = 0; i < 16; i++) {
            j = chunk * 64 + i * 4;
            w[i] = ((padded[j] << 24) | (padded[j+1] << 16) | (padded[j+2] << 8) | padded[j+3]) >>> 0;
        }
        for (i = 16; i < 80; i++) {
            w[i] = rotl32(w[i-3] ^ w[i-8] ^ w[i-14] ^ w[i-16], 1);
        }

        a=h0; b=h1; c=h2; d=h3; e=h4;

        for (i = 0; i < 80; i++) {
            if (i < 20) { f = (b & c) | ((~b) & d); k = 0x5A827999; }
            else if (i < 40) { f = b ^ c ^ d; k = 0x6ED9EBA1; }
            else if (i < 60) { f = (b & c) | (b & d) | (c & d); k = 0x8F1BBCDC; }
            else { f = b ^ c ^ d; k = 0xCA62C1D6; }
            temp = add32(rotl32(a,5), f, e, k, w[i]);
            e=d; d=c; c=rotl32(b,30); b=a; a=temp;
        }

        h0=add32(h0,a); h1=add32(h1,b); h2=add32(h2,c); h3=add32(h3,d); h4=add32(h4,e);
    }

    var out = [], words = [h0,h1,h2,h3,h4];
    for (i = 0; i < 5; i++) {
        out.push((words[i]>>>24)&0xFF,(words[i]>>>16)&0xFF,(words[i]>>>8)&0xFF,words[i]&0xFF);
    }
    return out;
}

function hmacSha256Bytes(keyBytes, msgBytes) {
    var blockSize = 64, key, i, oKeyPad = [], iKeyPad = [], inner;
    key = (keyBytes.length > blockSize) ? sha256Bytes(keyBytes) : keyBytes.slice();
    while (key.length < blockSize) { key.push(0); }
    for (i = 0; i < blockSize; i++) {
        oKeyPad.push(key[i] ^ 0x5c);
        iKeyPad.push(key[i] ^ 0x36);
    }
    inner = sha256Bytes(iKeyPad.concat(msgBytes));
    return sha256Bytes(oKeyPad.concat(inner));
}

function pbkdf2Block(passwordBytes, saltBytes, iterations, blockIndex) {
    var saltPlusIndex = saltBytes.concat([
        (blockIndex >>> 24) & 0xFF, (blockIndex >>> 16) & 0xFF,
        (blockIndex >>> 8) & 0xFF, blockIndex & 0xFF
    ]);
    var u = hmacSha256Bytes(passwordBytes, saltPlusIndex);
    var t = u.slice();
    var i, j;
    for (i = 1; i < iterations; i++) {
        u = hmacSha256Bytes(passwordBytes, u);
        for (j = 0; j < t.length; j++) { t[j] ^= u[j]; }
    }
    return t;
}

function pbkdf2Sha256Bytes(passwordBytes, saltBytes, iterations, dkLenBytes) {
    var hLen = 32, l = Math.ceil(dkLenBytes / hLen), dk = [], i, t;
    for (i = 1; i <= l; i++) {
        t = pbkdf2Block(passwordBytes, saltBytes, iterations, i);
        dk = dk.concat(t);
    }
    return dk.slice(0, dkLenBytes);
}

// Comparacion en tiempo constante (evita filtrar por temporizacion cuanto coincide un hash)
function ComparacionSegura(a, b) {
    if (a.length !== b.length) { return false; }
    var diff = 0, i;
    for (i = 0; i < a.length; i++) { diff |= (a.charCodeAt(i) ^ b.charCodeAt(i)); }
    return diff === 0;
}

// Generador de aleatoriedad 100% nativo (sin dependencias externas): Math.random() no es
// criptograficamente fuerte, pero aqui solo hace falta que la sal sea UNICA por contrasena
// (no secreta) para anular tablas rainbow, y el token CSRF se refuerza pasandolo por SHA-256
// junto con reloj+sesion. Si en el futuro se dispone de un generador mas fuerte, sustituir aqui.
function bytesAleatorios(n) {
    var bytes = [], i;
    for (i = 0; i < n; i++) { bytes.push(Math.floor(Math.random() * 256)); }
    return bytes;
}

function GenerarTokenAleatorio(numBytes) {
    var semilla = utf8Bytes(String(new Date().getTime()) + String(Math.random()) + String(Math.random()));
    var mezcla = sha256Bytes(semilla.concat(bytesAleatorios(numBytes)));
    return toHex(mezcla).substr(0, numBytes * 2);
}

function CalcularSha1Base64(password) {
    return bytesToBase64(sha1Bytes(utf8Bytes(password)));
}

function CrearHashPbkdf2(password, iteraciones) {
    var saltBytes = bytesAleatorios(16);
    var dk = pbkdf2Sha256Bytes(utf8Bytes(password), saltBytes, iteraciones, 32);
    return "PBKDF2SHA256$" + iteraciones + "$" + bytesToBase64(saltBytes) + "$" + bytesToBase64(dk);
}

function VerificarHashPbkdf2(password, valorAlmacenado) {
    var partes = String(valorAlmacenado).split("$");
    if (partes.length !== 4 || partes[0] !== "PBKDF2SHA256") { return false; }
    var iteraciones = parseInt(partes[1], 10);
    var saltBytes = base64ToBytes(partes[2]);
    var dk = pbkdf2Sha256Bytes(utf8Bytes(password), saltBytes, iteraciones, 32);
    return ComparacionSegura(bytesToBase64(dk), partes[3]);
}

// Punto de entrada unico para verificar login: decide el algoritmo segun la columna
// usuarios.algoritmo_password, para que legacy (SHA1B64) y ToperBTR (PBKDF2SHA256) convivan.
function VerificarPassword(password, hashAlmacenado, algoritmo) {
    if (algoritmo === "PBKDF2SHA256") {
        return VerificarHashPbkdf2(password, hashAlmacenado);
    }
    return ComparacionSegura(CalcularSha1Base64(password), String(hashAlmacenado));
}
</script>
