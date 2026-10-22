/**
 * Tapport — Editar datos de la cuenta del administrador + enviar el resumen mensual por correo.
 * Archivo NUEVO: crea "Cuenta.gs" en Apps Script y pega esto. Además, en Code.gs agrega estas
 * 2 líneas en doPost, junto a las demás acciones:
 *    if (body.action === "actualizar_cuenta_gerente") return actualizarCuentaGerente_(body);
 *    if (body.action === "enviar_reporte") return enviarReporte_(body);
 * Luego: Implementar > Administrar implementaciones > ✏️ > Nueva versión.
 * La primera vez que se envíe un correo, Google pedirá autorizar el permiso de Gmail (acéptalo).
 */
function cuentaAuth_(body) {
  const r = JSON.parse(doPost({ postData: { contents: JSON.stringify({ action: "login_gerente", email: body.email, password: body.password }) } }).getContent());
  return r && r.ok ? r : null;
}

function actualizarCuentaGerente_(body) {
  if (!cuentaAuth_(body)) return respuesta({ ok: false, error: "Sesión no válida." });
  const nombre = String(body.nombre || "").trim().replace(/\s+/g, " ").slice(0, 60);
  const tel = String(body.telefono || "").replace(/[^\d+]/g, "").slice(0, 15);
  if (nombre.length < 3) return respuesta({ ok: false, error: "Escribe tu nombre completo." });
  if (tel.replace(/\D/g, "").length < 6) return respuesta({ ok: false, error: "Escribe un teléfono válido." });
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_GERENTES);
  const datos = hoja.getDataRange().getValues();
  const email = String(body.email).trim().toLowerCase();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0] || "").trim().toLowerCase() === email) {
      hoja.getRange(i + 1, 3).setValue(nombre);        // C Nombre
      hoja.getRange(i + 1, 4).setNumberFormat("@").setValue(tel); // D Telefono
      return respuesta({ ok: true });
    }
  }
  return respuesta({ ok: false, error: "No encontramos tu cuenta." });
}

// Solo envía al correo del propio administrador (nunca a otra dirección).
function enviarReporte_(body) {
  if (!cuentaAuth_(body)) return respuesta({ ok: false, error: "Sesión no válida." });
  const html = String(body.html || "").slice(0, 60000);
  if (!html) return respuesta({ ok: false, error: "Reporte vacío." });
  const cache = CacheService.getScriptCache(), k = "rep_" + String(body.email).toLowerCase();
  const n = Number(cache.get(k) || 0);
  if (n >= 10) return respuesta({ ok: false, error: "Llegaste al límite de envíos por hoy." });
  cache.put(k, String(n + 1), 60 * 60 * 24);
  MailApp.sendEmail({ to: String(body.email).trim(), subject: String(body.asunto || "Resumen mensual Tapport").slice(0, 150), htmlBody: html, name: "Tapport" });
  return respuesta({ ok: true });
}
