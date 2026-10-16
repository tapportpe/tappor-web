/**
 * Tapport — Asistente IA para el panel de administrador (gratis con Gemini).
 *
 * PASOS (5 min):
 * 1. Saca una API key gratis en https://aistudio.google.com/apikey
 * 2. En Apps Script: Configuración del proyecto > Propiedades del script >
 *    agrega  GEMINI_API_KEY = (tu key)
 * 3. Pega TODO este archivo en tu proyecto (archivo nuevo, ej. IA.gs).
 * 4. Al INICIO de tu doPost(e), justo después de leer el JSON, agrega:
 *      if (data.action === "ia_asistente") return iaAsistente_(data);
 *    (usa el mismo nombre de variable con el que ya lees el body, ej. data / body / req)
 * 5. Implementar > Administrar implementaciones > Editar > Nueva versión > Implementar.
 *
 * Seguridad: solo responde si el correo/contraseña de gerente son válidos
 * (reutiliza tu propio "login_gerente") y limita el uso por gerente.
 */
const IA_MODELO = "gemini-2.5-flash";   // si Google lo renombra, cámbialo aquí
const IA_LIMITE_DIARIO = 40;            // consultas por gerente al día
const IA_MAX_DATOS = 30000;             // caracteres máximos de datos aceptados

function iaAsistente_(d) {
  try {
    // 1) Verificar credenciales con tu login existente
    const login = JSON.parse(doPost({ postData: { contents: JSON.stringify({
      action: "login_gerente", email: d.email, password: d.password }) } }).getContent());
    if (!login || !login.ok) return iaJson_({ ok: false, error: "Sesión no válida." });

    // 2) Límite diario por gerente
    const cache = CacheService.getScriptCache();
    const clave = "ia_" + String(d.email).toLowerCase() + "_" + Utilities.formatDate(new Date(), "America/Lima", "yyyyMMdd");
    const usadas = Number(cache.get(clave) || 0);
    if (usadas >= IA_LIMITE_DIARIO) return iaJson_({ ok: false, error: "Llegaste al límite diario de consultas." });
    cache.put(clave, String(usadas + 1), 60 * 60 * 26);

    const key = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!key) return iaJson_({ ok: false, error: "Falta configurar GEMINI_API_KEY." });

    const datos = String(d.datos || "").slice(0, IA_MAX_DATOS);
    const sistema =
      "Eres el asistente de Tapport, plataforma de reputación de meseros por reseñas de clientes. " +
      "Hablas con el gerente de '" + (login.restaurante || "un restaurante") + "'. " +
      "Responde SIEMPRE en español, claro, breve y accionable, con tono profesional y cercano. " +
      "Usa SOLO los datos entregados; si algo no está en los datos, dilo. No inventes cifras ni nombres. " +
      "Formato: frases cortas, viñetas con '- ' cuando ayude, **negrita** solo para lo importante.";
    let tarea;
    if (d.modo === "resumen") {
      tarea = "Redacta un resumen ejecutivo del equipo: 1) panorama general, 2) puntos fuertes, " +
              "3) problemas detectados según los comentarios, 4) 3 acciones concretas para esta semana. Máximo 180 palabras.";
    } else {
      tarea = "Pregunta del gerente: " + String(d.pregunta || "").slice(0, 500);
    }
    const hist = (d.historial || []).slice(-6).map(function (m) {
      return { role: m.rol === "ia" ? "model" : "user", parts: [{ text: String(m.texto).slice(0, 800) }] };
    });
    const contents = hist.concat([{ role: "user", parts: [{ text: "DATOS DEL EQUIPO:\n" + datos + "\n\n" + tarea }] }]);

    const resp = UrlFetchApp.fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/" + IA_MODELO + ":generateContent?key=" + key, {
        method: "post", contentType: "application/json", muteHttpExceptions: true,
        payload: JSON.stringify({
          systemInstruction: { parts: [{ text: sistema }] },
          contents: contents,
          generationConfig: { temperature: 0.4, maxOutputTokens: 700 }
        })
      });
    const j = JSON.parse(resp.getContentText());
    const texto = j.candidates && j.candidates[0] && j.candidates[0].content &&
                  j.candidates[0].content.parts.map(function (p) { return p.text || ""; }).join("");
    if (!texto) return iaJson_({ ok: false, error: (j.error && j.error.message) || "La IA no devolvió respuesta." });
    return iaJson_({ ok: true, texto: texto });
  } catch (err) {
    return iaJson_({ ok: false, error: "Error del asistente: " + err });
  }
}

function iaJson_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
