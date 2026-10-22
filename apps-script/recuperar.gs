/**
 * Tapport — Recuperar contraseña de GERENTES (los meseros ya usan establecer_password).
 * Archivo NUEVO: crea "Recuperar.gs" en Apps Script y pega esto.
 *
 * Flujo: el panel pide un código al correo (enviar_codigo_verificacion), lo verifica
 * (verificar_codigo) y finalmente llama a "restablecer_password_gerente" con la nueva
 * contraseña. Solo funciona si el correo acaba de verificarse (igual que crear cuenta).
 *
 * Además de pegar este archivo, en doPost(e) agrega junto a "ia_asistente":
 *    if (body.action === "restablecer_password_gerente") return restablecerPasswordGerente_(body);
 * (el Code.gs que te entregué ya lo trae). Luego: Implementar > Administrar implementaciones >
 * ✏️ > Nueva versión > Implementar.
 */
function restablecerPasswordGerente_(body) {
  const email = String(body.email || "").trim().toLowerCase();
  const password = String(body.password || "");
  if (password.length < 6) return respuesta({ ok: false, error: "La contraseña debe tener al menos 6 caracteres" });

  const cache = CacheService.getScriptCache();
  if (cache.get("verif_ok_" + email) !== "true") {
    return respuesta({ ok: false, error: "Primero debes verificar tu correo con el código" });
  }

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_GERENTES);
  if (!hoja) return respuesta({ ok: false, error: "No encontramos una cuenta de administrador con este correo" });
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0] || "").trim().toLowerCase() === email) {
      // Cuenta dada de baja desde la hoja Ingresos (restaurante Inactivo): no se puede recuperar ni entrar.
      if (String(datos[i][7] || "").indexOf("SUSPENDIDO|") === 0) {
        return respuesta({ ok: false, error: "Tu cuenta está inactiva. Escríbenos por WhatsApp al +51 992 346 227 para reactivarla." });
      }
      const salt = generarSalt();
      hoja.getRange(i + 1, 8).setValue(hashPassword(password, salt)); // columna H: hash
      hoja.getRange(i + 1, 9).setValue(salt);                          // columna I: salt
      cache.remove("verif_ok_" + email);
      return respuesta({ ok: true });
    }
  }
  return respuesta({ ok: false, error: "No encontramos una cuenta de administrador con este correo" });
}
