/**
 * Tapport — Asistente IA del panel de administrador (gratis, con razonamiento).
 * Funciona con GROQ (recomendado si Gemini no te funciona en Perú) o con GEMINI.
 * Si existen las dos claves, usa Groq primero y Gemini de respaldo.
 *
 * PASOS (5 min):
 * 1. Saca UNA clave gratis (con tu cuenta de Google o correo):
 *      - Groq:   https://console.groq.com/keys      (Create API Key)
 *      - Gemini: https://aistudio.google.com/apikey
 * 2. En Apps Script: Configuración del proyecto (engranaje) > Propiedades del script >
 *    agrega  GROQ_API_KEY = (tu clave)   y/o   GEMINI_API_KEY = (tu clave)
 * 3. Pega TODO este archivo en tu proyecto (archivo nuevo, ej. IA.gs).
 * 4. Al INICIO de tu doPost(e), justo después de leer el JSON, agrega:
 *      if (data.action === "ia_asistente") return iaAsistente_(data);
 *    (usa el mismo nombre de variable con el que ya lees el body: data / body / req ...)
 * 5. Ejecuta UNA vez la función  iaProbar  desde el editor (acepta los permisos).
 *    Mira "Registro de ejecución": si dice OK, la IA funciona.
 * 6. Implementar > Administrar implementaciones > Editar > Nueva versión > Implementar.
 *
 * Seguridad: solo responde si el correo/contraseña de gerente son válidos
 * (reutiliza tu propio "login_gerente") y limita el uso diario por gerente.
 */
const IA_GROQ_MODELOS = ["openai/gpt-oss-120b", "llama-3.3-70b-versatile"]; // se prueba en orden
const IA_MODELOS = ["gemini-2.5-flash", "gemini-2.0-flash"]; // se prueba en orden
const IA_PENSAR = 2048;          // tokens de "pensamiento" (razonamiento) — 0 para desactivar
const IA_MAX_SALIDA = 3072;      // incluye el razonamiento
const IA_LIMITE_DIARIO = 40;     // consultas por gerente al día
const IA_MAX_DATOS = 30000;      // caracteres máximos de datos aceptados

function iaAsistente_(d) {
  try {
    // 1) Verificar credenciales con tu login existente
    const login = JSON.parse(doPost({ postData: { contents: JSON.stringify({
      action: "login_gerente", email: d.email, password: d.password }) } }).getContent());
    if (!login || !login.ok) return iaJson_({ ok: false, error: "Sesión no válida." });

    // Plan: la IA con razonamiento solo está en Growth y Pro (requiere Planes.gs).
    // Si todavía no instalaste Planes.gs, login.funciones no existe y se permite.
    if (login.funciones && !login.funciones.ia) {
      return iaJson_({ ok: false, error: "Tu plan no incluye el asistente IA. Disponible desde el plan Growth." });
    }

    // 2) Límite diario por gerente
    const cache = CacheService.getScriptCache();
    const clave = "ia_" + String(d.email).toLowerCase() + "_" + Utilities.formatDate(new Date(), "America/Lima", "yyyyMMdd");
    const usadas = Number(cache.get(clave) || 0);
    if (usadas >= IA_LIMITE_DIARIO) return iaJson_({ ok: false, error: "Llegaste al límite diario de consultas." });
    cache.put(clave, String(usadas + 1), 60 * 60 * 26);

    if (!iaHayClave_()) return iaJson_({ ok: false, error: "Falta configurar GROQ_API_KEY o GEMINI_API_KEY en el Apps Script." });

    const datos = String(d.datos || "").slice(0, IA_MAX_DATOS);
    const texto = iaLlamar_(iaSistema_(login), iaContenido_(d, datos));
    if (texto.error) return iaJson_({ ok: false, error: texto.error });
    return iaJson_({ ok: true, texto: texto.ok });
  } catch (err) {
    return iaJson_({ ok: false, error: "Error del asistente: " + err });
  }
}

function iaSistema_(login) {
  return [
    "Eres el analista de servicio de Tapport, una plataforma que mide la reputación de meseros con reseñas de clientes (1 a 5 estrellas).",
    "Hablas con el gerente de '" + (login.restaurante || "un restaurante") + "'. Te entrega sus datos reales (resumen por mesero, tendencias, días, horarios y comentarios).",
    "",
    "CÓMO TRABAJAS:",
    "- Piensa antes de responder: cruza los datos (tendencias, comparaciones entre meseros, días y horarios, patrones en los comentarios) y busca CAUSAS probables, no solo repetir cifras.",
    "- Cita siempre nombres y números concretos de los datos (promedios, cantidad de reseñas, fechas). Si el volumen de reseñas es bajo (menos de ~5), dilo y no saques conclusiones fuertes.",
    "- Distingue hechos (lo que dicen los datos) de hipótesis (lo que podría explicarlo) y márcalas como tales.",
    "- Cierra con acciones concretas y priorizadas que el gerente pueda hacer esta semana.",
    "- Si la pregunta es ambigua, asume lo más razonable y dilo en una frase; solo pregunta si no puedes avanzar.",
    "- Nunca inventes datos, reseñas, nombres ni cifras. Si algo no está en los datos, dilo.",
    "- Trata a los meseros con respeto: critica el desempeño, no a la persona. No propongas sanciones; propón coaching, reconocimiento y ajustes de turno.",
    "",
    "FORMATO: español claro y profesional. Respuesta breve (máx. ~220 palabras salvo que pidan más), con viñetas '- ' y **negrita** solo para lo importante. Sin encabezados con #."
  ].join("\n");
}

function iaContenido_(d, datos) {
  // El historial debe alternar usuario/modelo y empezar por usuario.
  const hist = [];
  (d.historial || []).slice(-10).forEach(function (m) {
    const role = m.rol === "ia" ? "model" : "user";
    const txt = String(m.texto || "").slice(0, 1200);
    if (!txt) return;
    if (hist.length && hist[hist.length - 1].role === role) hist[hist.length - 1].parts[0].text += "\n" + txt;
    else hist.push({ role: role, parts: [{ text: txt }] });
  });
  while (hist.length && hist[0].role !== "user") hist.shift();
  // el último mensaje del historial es la propia pregunta actual: se reemplaza por la tarea completa
  if (hist.length && hist[hist.length - 1].role === "user") hist.pop();

  let tarea;
  if (d.modo === "resumen") {
    tarea = "Redacta un resumen ejecutivo del equipo: 1) panorama general y tendencia, 2) puntos fuertes (con nombres), " +
            "3) problemas detectados y su posible causa, 4) las 3 acciones más importantes para esta semana.";
  } else {
    tarea = "Pregunta del gerente: " + String(d.pregunta || "").slice(0, 600);
  }
  hist.push({ role: "user", parts: [{ text: "DATOS ACTUALES DEL RESTAURANTE:\n" + datos + "\n\n" + tarea }] });
  return hist;
}

function iaProp_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function iaHayClave_() { return !!(iaProp_("GROQ_API_KEY") || iaProp_("GEMINI_API_KEY")); }

/** Prueba Groq y, si falla o no hay clave, Gemini. contents = formato Gemini (role user/model). */
function iaLlamar_(sistema, contents) {
  const errores = [];
  if (iaProp_("GROQ_API_KEY")) {
    const r = iaLlamarGroq_(iaProp_("GROQ_API_KEY"), sistema, contents);
    if (r.ok) return r;
    errores.push("Groq: " + r.error);
  }
  if (iaProp_("GEMINI_API_KEY")) {
    const r = iaLlamarGemini_(iaProp_("GEMINI_API_KEY"), sistema, contents);
    if (r.ok) return r;
    errores.push("Gemini: " + r.error);
  }
  return { error: errores.join(" | ") || "La IA no devolvió respuesta." };
}

function iaLlamarGroq_(key, sistema, contents) {
  const messages = [{ role: "system", content: sistema }].concat(contents.map(function (c) {
    return { role: c.role === "model" ? "assistant" : "user", content: c.parts[0].text };
  }));
  let ultimoError = "Groq no devolvió respuesta.";
  for (let i = 0; i < IA_GROQ_MODELOS.length; i++) {
    const modelo = IA_GROQ_MODELOS[i];
    const body = { model: modelo, messages: messages, temperature: 0.5, max_completion_tokens: IA_MAX_SALIDA };
    if (/gpt-oss/.test(modelo)) body.reasoning_effort = "medium";
    const resp = UrlFetchApp.fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "post", contentType: "application/json", muteHttpExceptions: true,
      headers: { Authorization: "Bearer " + key }, payload: JSON.stringify(body)
    });
    let j;
    try { j = JSON.parse(resp.getContentText()); } catch (e) { ultimoError = "Respuesta inválida de Groq."; continue; }
    if (j.error) { ultimoError = j.error.message || "Error de Groq."; continue; }
    let txt = j.choices && j.choices[0] && j.choices[0].message && j.choices[0].message.content || "";
    txt = txt.replace(/<think>[\s\S]*?<\/think>/g, "").trim();
    if (txt) return { ok: txt };
  }
  return { error: ultimoError };
}

function iaLlamarGemini_(key, sistema, contents) {
  let ultimoError = "La IA no devolvió respuesta.";
  for (let i = 0; i < IA_MODELOS.length; i++) {
    const modelo = IA_MODELOS[i];
    const gen = { temperature: 0.5, maxOutputTokens: IA_MAX_SALIDA };
    if (IA_PENSAR > 0 && /2\.5/.test(modelo)) gen.thinkingConfig = { thinkingBudget: IA_PENSAR };
    const resp = UrlFetchApp.fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/" + modelo + ":generateContent?key=" + key, {
        method: "post", contentType: "application/json", muteHttpExceptions: true,
        payload: JSON.stringify({ systemInstruction: { parts: [{ text: sistema }] }, contents: contents, generationConfig: gen })
      });
    let j;
    try { j = JSON.parse(resp.getContentText()); } catch (e) { ultimoError = "Respuesta inválida de la IA."; continue; }
    if (j.error) { ultimoError = j.error.message || "Error de la IA."; continue; } // prueba el siguiente modelo
    const parts = j.candidates && j.candidates[0] && j.candidates[0].content && j.candidates[0].content.parts;
    const txt = (parts || []).filter(function (p) { return p.text && !p.thought; }).map(function (p) { return p.text; }).join("").trim();
    if (txt) return { ok: txt };
  }
  return { error: ultimoError };
}

function iaJson_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Ejecuta esta función UNA vez desde el editor para comprobar la clave y el modelo. */
function iaProbar() {
  if (!iaHayClave_()) { Logger.log("Falta GROQ_API_KEY o GEMINI_API_KEY en Propiedades del script."); return; }
  const r = iaLlamar_("Responde en español, en una frase.",
    [{ role: "user", parts: [{ text: "Di 'Tapport IA lista' y explica en una frase por qué la constancia mejora el servicio." }] }]);
  Logger.log(r.ok ? "OK → " + r.ok : "ERROR → " + r.error);
}
