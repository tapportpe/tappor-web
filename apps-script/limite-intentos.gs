/**
 * Archivo NUEVO: crea "LimiteIntentos.gs" en Apps Script y pega esto. No cambia nada por sí solo.
 * Protege logins y códigos de correo contra fuerza bruta (máx. 5 fallos -> bloqueo 15 min).
 *
 * USO en Code.gs (3 pasos, mismo patrón en loginMesero, loginGerente y verificarCodigo):
 *
 *   function loginMesero(body) {
 *     const clave = "login_" + String(body.email || "").toLowerCase();
 *     const bloqueo = intentosBloqueado_(clave);
 *     if (bloqueo) return respuesta({ ok: false, error: bloqueo });      // 1) al inicio
 *     ...validar credenciales...
 *     if (!credencialesOk) { intentoFallido_(clave); return respuesta({ ok:false, error:"Datos incorrectos" }); }  // 2) en cada fallo
 *     intentosLimpiar_(clave);                                            // 3) al acertar
 *     ...
 *   }
 *
 * Códigos de correo: guárdalos con codigoGuardar_(email, codigo) y compruébalos con
 * codigoValidar_(email, ingresado). Expiran solos a los 10 min y se borran al usarse o tras 5 fallos.
 */
const LIM_MAX = 5, LIM_BLOQUEO_SEG = 900, COD_VIDA_SEG = 600;

function intentosBloqueado_(clave) {
  const n = Number(CacheService.getScriptCache().get("int_" + clave) || 0);
  return n >= LIM_MAX ? "Demasiados intentos. Espera 15 minutos e inténtalo de nuevo." : "";
}
function intentoFallido_(clave) {
  const c = CacheService.getScriptCache(), k = "int_" + clave;
  c.put(k, String(Number(c.get(k) || 0) + 1), LIM_BLOQUEO_SEG);
}
function intentosLimpiar_(clave) { CacheService.getScriptCache().remove("int_" + clave); }

function codigoGuardar_(email, codigo) {
  const e = String(email).toLowerCase();
  CacheService.getScriptCache().put("cod_" + e, String(codigo), COD_VIDA_SEG);
  intentosLimpiar_("cod_" + e);
}
// Devuelve {ok:true} o {ok:false, error:"..."}
function codigoValidar_(email, ingresado) {
  const e = String(email).toLowerCase(), clave = "cod_" + e, c = CacheService.getScriptCache();
  const bloqueo = intentosBloqueado_(clave);
  if (bloqueo) return { ok: false, error: bloqueo };
  const real = c.get("cod_" + e);
  if (!real) return { ok: false, error: "El código venció. Pide uno nuevo." };
  if (String(ingresado).trim() !== real) {
    intentoFallido_(clave);
    if (intentosBloqueado_(clave)) c.remove("cod_" + e);
    return { ok: false, error: "Código incorrecto." };
  }
  c.remove("cod_" + e); intentosLimpiar_(clave);
  return { ok: true };
}
