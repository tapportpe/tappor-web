/**
 * Tapport — Soporte con IA.
 * Archivo NUEVO: crea "Soporte.gs" en Apps Script y pega esto.
 *
 * Qué hace: el gerente o el mesero escribe su problema en el panel; la IA lee SUS datos
 * en el Sheet (cuenta, plan, equipo, reseñas, vacantes) + una guía de ayuda, y responde.
 * Si no puede resolverlo, devuelve escalar:true y el panel muestra el botón de WhatsApp.
 * Cada consulta queda registrada en una hoja "Soporte" (se crea sola) para que la revises.
 *
 * Requiere IA.gs (usa su iaLlamar_ y GROQ_API_KEY / GEMINI_API_KEY ya configuradas).
 *
 * PASOS:
 * 1. Pega este archivo (Soporte.gs).
 * 2. En doPost(e), junto a la línea de "ia_asistente", agrega:
 *      if (body.action === "soporte_ia") return soporteIA_(body);
 * 3. Ejecuta UNA vez  soporteProbar  desde el editor (acepta permisos) y mira el registro.
 * 4. Implementar > Administrar implementaciones > ✏️ > Nueva versión > Implementar.
 *
 * Privacidad: a la IA nunca se le envían contraseñas, hashes, códigos, DNI ni teléfonos,
 * ni datos de otros restaurantes. Solo ve lo del usuario que escribe.
 */
const SOPORTE_LIMITE_DIARIO = 25;
const SOPORTE_HOJA = "Soporte";
const SOPORTE_WSP = "51992346227";
const SOPORTE_SENSIBLE = /contrase|password|hash|salt|c[oó]digo|token|dni|documento|tel[eé]fono|celular|whatsapp/i;

const SOPORTE_GUIA = [
  "GUÍA DE AYUDA DE TAPPORT (úsala para responder):",
  "- Tapport mide la reputación de meseros con reseñas de clientes (1 a 5 estrellas). Cada mesero tiene una pulsera/tarjeta NFC: el cliente la acerca al celular, califica y deja un comentario.",
  "- Planes: Starter S/125 al mes (1 a 5 meseros), Growth S/175 (6 a 10 meseros, incluye asistente IA con razonamiento), Pro S/225 (11 o más). Todos incluyen pulseras NFC y panel de administración. Cambiar de plan o renovar se hace por WhatsApp con el equipo.",
  "- Cada mesero se registra por su cuenta en 'Crear mi cuenta' (perfil de mesero): verifica su correo con un código de 6 dígitos, crea su contraseña y elige el restaurante de la lista. El gerente no crea cuentas de meseros.",
  "- Si el restaurante llegó al límite de meseros de su plan, los nuevos meseros no pueden registrarse hasta subir de plan.",
  "- Gerentes: se registran en 'Crear cuenta de gerente' con foto, restaurante, datos personales, correo verificado y contraseña. El restaurante debe estar creado por el equipo Tapport; si no aparece en la lista, hay que escribir por WhatsApp.",
  "- Si alguien olvidó su contraseña o no puede entrar: revisar que use el mismo correo con el que se registró, que haya creado su contraseña (si solo verificó el correo pero no la creó, debe terminar 'Crear mi cuenta'), y que no haya espacios extra. Para restablecerla hay que escribir al soporte humano.",
  "- La sesión se cierra sola tras un tiempo de inactividad: basta volver a iniciar sesión.",
  "- Cambiar foto: en el perfil, 'Editar perfil' > cambiar foto. La foto anterior se reemplaza. Puede tardar unos segundos en verse; recargar con Ctrl+F5.",
  "- Si el panel de gerente muestra 'Sin plan asignado': el plan lo activa el equipo Tapport; si ya pagó, escribir por WhatsApp para activarlo.",
  "- Plan vencido: el panel sigue visible pero las funciones del plan pueden bloquearse; renovar por WhatsApp.",
  "- Reseñas: aparecen cuando el cliente completa la calificación. Si un mesero no ve reseñas: confirmar que su pulsera esté asignada a su código y que haya habido calificaciones reales.",
  "- Vacantes: los restaurantes publican vacantes desde el panel de gerente; los meseros postulan desde su perfil. Se editan o eliminan desde la sección Administrar.",
  "- Asistente IA del panel (gerente): disponible en planes Growth y Pro, con límite diario de consultas.",
  "- El Feed/Directorio es público: muestra meseros, restaurantes con mejor atención y vacantes."
].join("\n");

function soporteIA_(d) {
  try {
    const tipo = d.tipo === "mesero" ? "mesero" : "gerente";
    const email = String(d.email || "").trim().toLowerCase();
    const pregunta = String(d.pregunta || "").trim().slice(0, 600);
    if (!pregunta) return iaJson_({ ok: false, error: "Escribe tu consulta." });

    // 1) Autenticación reutilizando tus logins existentes
    const accion = tipo === "mesero" ? "login" : "login_gerente";
    const login = JSON.parse(doPost({ postData: { contents: JSON.stringify({
      action: accion, email: email, password: d.password }) } }).getContent());
    if (!login || !login.ok) return iaJson_({ ok: false, error: "Tu sesión no es válida. Vuelve a iniciar sesión." });

    // 2) Límite diario
    const cache = CacheService.getScriptCache();
    const clave = "sop_" + email + "_" + Utilities.formatDate(new Date(), "America/Lima", "yyyyMMdd");
    const usadas = Number(cache.get(clave) || 0);
    if (usadas >= SOPORTE_LIMITE_DIARIO) return iaJson_({ ok: true, texto: "Llegaste al límite de consultas de hoy. Escríbenos por WhatsApp y te ayudamos.", escalar: true });
    cache.put(clave, String(usadas + 1), 60 * 60 * 26);

    if (!iaHayClave_()) return iaJson_({ ok: true, texto: "El asistente de soporte aún no está disponible. Escríbenos por WhatsApp.", escalar: true });

    // 3) Datos del usuario (sin información sensible)
    const contexto = tipo === "mesero" ? soporteContextoMesero_(email, login) : soporteContextoGerente_(email, login);

    // 4) Historial corto
    const contents = [];
    (d.historial || []).slice(-8).forEach(function (m) {
      const role = m.rol === "ia" ? "model" : "user", txt = String(m.texto || "").slice(0, 800);
      if (!txt) return;
      if (contents.length && contents[contents.length - 1].role === role) contents[contents.length - 1].parts[0].text += "\n" + txt;
      else contents.push({ role: role, parts: [{ text: txt }] });
    });
    while (contents.length && contents[0].role !== "user") contents.shift();
    if (contents.length && contents[contents.length - 1].role === "user") contents.pop();
    contents.push({ role: "user", parts: [{ text: "DATOS DE LA CUENTA DE ESTE USUARIO:\n" + contexto + "\n\nCONSULTA: " + pregunta }] });

    const sistema = [
      "Eres el asistente de soporte de Tapport. Atiendes a un " + (tipo === "mesero" ? "MESERO" : "GERENTE/ADMINISTRADOR de restaurante") + " dentro de su panel.",
      "Resuelve su problema usando (1) los DATOS DE SU CUENTA que se te entregan y (2) la GUÍA. Diagnostica con datos concretos: por ejemplo, si el plan está vencido, si llegó al límite de meseros, si falta una foto, si su cuenta no tiene contraseña creada.",
      "REGLAS:",
      "- Responde en español, claro y amable, en pocas líneas (máx. ~120 palabras) con pasos numerados si hay que hacer algo.",
      "- Nunca inventes datos, funciones, precios ni plazos. Si algo no está en los datos o la guía, dilo.",
      "- Nunca pidas ni muestres contraseñas, códigos, DNI ni datos de otras personas o restaurantes.",
      "- No hagas tablas ni gráficos. Si el usuario quiere analizar sus reseñas, dile que use el Asistente IA del panel.",
      "- Si NO puedes resolverlo (cambios de plan o pagos, restablecer contraseña, errores técnicos, reclamos, o no tienes datos suficientes), explica brevemente por qué y termina tu respuesta con la marca exacta [[HUMANO]] para derivarlo a una persona.",
      "",
      SOPORTE_GUIA
    ].join("\n");

    const r = iaLlamar_(sistema, contents);
    if (r.error) return iaJson_({ ok: true, texto: "Ahora mismo no puedo responder. Escríbenos por WhatsApp y te ayudamos.", escalar: true });
    let texto = String(r.ok || "");
    const escalar = /\[\[HUMANO\]\]/.test(texto);
    texto = texto.replace(/\[\[HUMANO\]\]/g, "").trim();
    soporteRegistrar_(tipo, email, pregunta, texto, escalar);
    return iaJson_({ ok: true, texto: texto, escalar: escalar });
  } catch (err) {
    return iaJson_({ ok: false, error: "Error de soporte: " + err });
  }
}

/** Convierte una fila en "Encabezado: valor", sin columnas sensibles ni vacías. */
function soporteFila_(encabezados, fila) {
  const out = [];
  encabezados.forEach(function (h, i) {
    const nombre = String(h || "").trim();
    if (!nombre || SOPORTE_SENSIBLE.test(nombre)) return;
    let v = fila[i];
    if (v === "" || v === null || v === undefined) return;
    if (v instanceof Date) v = Utilities.formatDate(v, "America/Lima", "dd/MM/yyyy");
    v = String(v);
    if (/^https?:/i.test(v)) v = "(tiene archivo/foto)"; // no enviar enlaces
    out.push(nombre + ": " + v.slice(0, 160));
  });
  return out.join(" | ");
}

function soporteHoja_(nombre) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre); }

function soporteIdx_(encabezados, regex) {
  for (let i = 0; i < encabezados.length; i++) if (regex.test(String(encabezados[i] || ""))) return i;
  return -1;
}

function soporteContextoGerente_(email, login) {
  const L = [];
  const rest = String(login.restaurante || "");
  L.push("Tipo: gerente. Restaurante: " + rest + ".");
  L.push("Plan: " + (login.planNombre || "sin plan asignado") + (login.planVence ? ", vence " + Utilities.formatDate(new Date(login.planVence), "America/Lima", "dd/MM/yyyy") : "") +
    ", " + (login.planActivo === false ? "VENCIDO" : "activo") + ", máximo de meseros: " + (login.maxMeseros || "n/d") + ", asistente IA: " + (login.funciones && login.funciones.ia ? "sí" : "no") + ".");

  // Su fila en Administradores
  const hA = soporteHoja_(HOJA_GERENTES);
  if (hA) {
    const dA = hA.getDataRange().getValues(), ie = soporteIdx_(dA[0], /correo|e-?mail/i);
    for (let i = 1; i < dA.length; i++) if (ie >= 0 && String(dA[i][ie]).trim().toLowerCase() === email) { L.push("Su cuenta -> " + soporteFila_(dA[0], dA[i])); break; }
  }
  // Meseros del restaurante
  const hM = soporteHoja_(HOJA_TRABAJADORES);
  if (hM) {
    const dM = hM.getDataRange().getValues(), ir = soporteIdx_(dM[0], /restaurante/i), ic = soporteIdx_(dM[0], /contrase|password|hash/i);
    const eq = [];
    for (let i = 1; i < dM.length; i++) if (ir >= 0 && String(dM[i][ir]).trim().toLowerCase() === rest.toLowerCase()) {
      eq.push(soporteFila_(dM[0], dM[i]) + (ic >= 0 ? " | Contraseña creada: " + (dM[i][ic] ? "sí" : "NO") : ""));
    }
    L.push("Meseros registrados en su restaurante (" + eq.length + "):\n" + eq.slice(0, 25).map(function (x) { return "  * " + x; }).join("\n"));
  }
  // Config del restaurante
  const hC = soporteHoja_(HOJA_CONFIG);
  if (hC) {
    const dC = hC.getDataRange().getValues(), ir2 = soporteIdx_(dC[0], /restaurante|negocio|nombre/i);
    for (let i = 1; i < dC.length; i++) if (ir2 >= 0 && String(dC[i][ir2]).trim().toLowerCase() === rest.toLowerCase()) { L.push("Configuración del negocio -> " + soporteFila_(dC[0], dC[i])); break; }
  }
  // Resumen de reseñas
  const hR = soporteHoja_(HOJA_RESENAS);
  if (hR) {
    const dR = hR.getDataRange().getValues(), ir3 = soporteIdx_(dR[0], /restaurante/i), ie3 = soporteIdx_(dR[0], /estrella|calific|rating/i);
    let n = 0, s = 0;
    for (let i = 1; i < dR.length; i++) if (ir3 >= 0 && String(dR[i][ir3]).trim().toLowerCase() === rest.toLowerCase()) { n++; s += Number(dR[i][ie3]) || 0; }
    L.push("Reseñas del restaurante: " + n + (n ? ", promedio " + (s / n).toFixed(2) : "") + ".");
  }
  return L.join("\n");
}

function soporteContextoMesero_(email, login) {
  const L = ["Tipo: mesero."];
  const hM = soporteHoja_(HOJA_TRABAJADORES);
  let rest = "", nombre = "";
  if (hM) {
    const dM = hM.getDataRange().getValues(), ie = soporteIdx_(dM[0], /correo|e-?mail/i), ir = soporteIdx_(dM[0], /restaurante/i), ic = soporteIdx_(dM[0], /contrase|password|hash/i), inm = soporteIdx_(dM[0], /nombre/i);
    for (let i = 1; i < dM.length; i++) if (ie >= 0 && String(dM[i][ie]).trim().toLowerCase() === email) {
      L.push("Su cuenta -> " + soporteFila_(dM[0], dM[i]) + (ic >= 0 ? " | Contraseña creada: " + (dM[i][ic] ? "sí" : "NO") : ""));
      rest = ir >= 0 ? String(dM[i][ir]) : ""; nombre = inm >= 0 ? String(dM[i][inm]) : "";
      break;
    }
  }
  const hR = soporteHoja_(HOJA_RESENAS);
  if (hR && nombre) {
    const dR = hR.getDataRange().getValues(), im = soporteIdx_(dR[0], /mesero|trabajador|nombre/i), ie3 = soporteIdx_(dR[0], /estrella|calific|rating/i);
    let n = 0, s = 0;
    for (let i = 1; i < dR.length; i++) if (im >= 0 && String(dR[i][im]).trim().toLowerCase() === nombre.trim().toLowerCase()) { n++; s += Number(dR[i][ie3]) || 0; }
    L.push("Sus reseñas: " + n + (n ? ", promedio " + (s / n).toFixed(2) : "") + ".");
  }
  if (rest) {
    try { const p = infoPlan(rest); L.push("Plan de su restaurante: " + (p.planNombre || "sin plan") + (p.planActivo === false ? " (VENCIDO)" : "") + "."); } catch (e) { }
  }
  return L.join("\n");
}

function soporteRegistrar_(tipo, email, pregunta, respuesta, escalado) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let h = ss.getSheetByName(SOPORTE_HOJA);
    if (!h) {
      h = ss.insertSheet(SOPORTE_HOJA);
      h.appendRow(["Fecha", "Tipo", "Correo", "Consulta", "Respuesta IA", "Derivado a humano"]);
      h.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#0A1530").setFontColor("#FFFFFF");
      h.setFrozenRows(1);
    }
    h.appendRow([new Date(), tipo, email, pregunta, respuesta, escalado ? "SÍ" : "no"]);
  } catch (e) { }
}

/** Ejecuta UNA vez para ver qué datos recibiría la IA de un gerente (cambia el correo). */
function soporteProbar() {
  const email = "TU_CORREO_DE_GERENTE@gmail.com";
  const ctx = soporteContextoGerente_(email, { restaurante: "TU RESTAURANTE", planNombre: "Growth", maxMeseros: 10, funciones: { ia: true } });
  Logger.log(ctx);
}
