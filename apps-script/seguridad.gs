/**
 * Archivo NUEVO: crea "Seguridad.gs" en Apps Script y pega esto. No modifica Code.gs por sí solo.
 * Trae las piezas que faltaban; cada una indica dónde enchufarla.
 * (Los límites de login/código están en LimiteIntentos.gs.)
 */

/* ───────── 1) Token de sesión con vencimiento (en vez de reenviar la contraseña) ─────────
 * En loginMesero/loginGerente, al validar OK:   const token = tokenCrear_(email, "mesero"|"gerente");
 *   y devuelve { ok:true, token, ... } SIN devolver passwordHash/salt.
 * En el front guarda solo `token` (no la contraseña) y mándalo en cada llamada.
 * En cada acción protegida (reseñas del gerente, publicar vacante, editar perfil, soporte_ia, ia_asistente…):
 *   const ses = tokenValidar_(body.token, "gerente");  if (!ses) return respuesta({ok:false,error:"Sesión vencida"});
 *   y usa ses.email (NUNCA el email que llegue en el body). Al cerrar sesión: tokenBorrar_(token).
 */
const TOK_VIDA_SEG = 6 * 3600; // máximo de CacheService
function tokenCrear_(email, tipo) {
  const t = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, "");
  CacheService.getScriptCache().put("tok_" + t, JSON.stringify({ email: String(email).toLowerCase(), tipo: tipo }), TOK_VIDA_SEG);
  return t;
}
function tokenValidar_(token, tipo) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get("tok_" + token);
  if (!raw) return null;
  const s = JSON.parse(raw);
  return (!tipo || s.tipo === tipo) ? s : null;
}
function tokenBorrar_(token) { if (token) CacheService.getScriptCache().remove("tok_" + token); }

/* ───────── 2) Endpoints públicos: solo campos permitidos + auditoría ─────────
 * Al devolver datos públicos usa:  respuesta(soloCampos_(obj, CAMPOS_PUBLICOS_MESERO))
 * Corre auditarEndpointsPublicos() desde el editor: te dice si algún endpoint filtra DNI/teléfono/correo/hash.
 */
const CAMPOS_PUBLICOS_MESERO = ["codigo", "nombre", "restaurante", "fotoUrl", "promedio", "totalReviews"];
function soloCampos_(obj, permitidos) {
  const o = {}; permitidos.forEach(k => { if (obj && obj[k] !== undefined) o[k] = obj[k]; }); return o;
}
function auditarEndpointsPublicos() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Meseros");
  const codigo = h && h.getLastRow() > 1 ? String(h.getRange(2, 1).getValue()) : "";
  const pruebas = [
    { accion: "buscar_meseros", minRating: "0" },
    { accion: "restaurantes" },
    { accion: "vacantes" }
  ];
  if (codigo) { pruebas.push({ accion: "perfil_publico", code: codigo }); pruebas.push({ accion: "buscar_mesero", code: codigo }); }
  const malo = /"(dni|telefono|teléfono|email|correo|password|passwordHash|salt|edad)"\s*:/i;
  pruebas.forEach(p => {
    let txt = "";
    try { txt = doGet({ parameter: p }).getContent(); } catch (e) { Logger.log("⚠️ " + p.accion + " falló: " + e); return; }
    const m = txt.match(malo);
    Logger.log((m ? "❌ FILTRA " + m[1] : "✅ ok") + "  ←  " + JSON.stringify(p));
  });
}

/* ───────── 3) Límite de reseñas (anti-spam) ─────────
 * En guardarResena, al inicio:  const lim = limiteResena_(body.code, body.dispositivo);
 *   if (lim) return respuesta({ ok:false, error: lim });
 * El front debe mandar `dispositivo` (un id aleatorio guardado en localStorage).
 * Máx. 1 reseña por dispositivo y pulsera cada 6 h, y 30 por pulsera cada hora.
 */
function limiteResena_(codigo, dispositivo) {
  const c = CacheService.getScriptCache(), cod = String(codigo || "").slice(0, 40);
  const dev = String(dispositivo || "").slice(0, 64);
  if (dev) {
    const kd = "rs_" + cod + "_" + dev;
    if (c.get(kd)) return "Ya calificaste a este mesero hace poco. Gracias por tu opinión.";
  }
  const kh = "rh_" + cod, n = Number(c.get(kh) || 0);
  if (n >= 30) return "Este mesero recibió demasiadas reseñas en poco tiempo. Inténtalo más tarde.";
  if (dev) c.put("rs_" + cod + "_" + dev, "1", 6 * 3600);
  c.put(kh, String(n + 1), 3600);
  return "";
}

/* ───────── 4) Texto seguro para el Sheet (inyección de fórmulas) ─────────
 * Envuelve TODO texto escrito por usuarios antes de setValue/appendRow: comentario, nombre, mensaje, puesto, etc.
 *   hoja.appendRow([fecha, codigo, textoSeguro_(restaurante), textoSeguro_(nombre), estrellas, textoSeguro_(tags), textoSeguro_(comentario)]);
 */
function textoSeguro_(v, max) {
  let s = String(v == null ? "" : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "").slice(0, max || 500);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

/* ───────── 5) Validar fotos en el servidor ─────────
 * En registrarMesero / editarPerfil / actualizarNegocio, ANTES de guardar en Drive:
 *   const f = validarFotoBase64_(body.fotoBase64);  if (!f.ok) return respuesta({ok:false,error:f.error});
 *   const blob = Utilities.newBlob(f.bytes, f.mime, "foto." + f.ext);
 */
const FOTO_MAX_BYTES = 2 * 1024 * 1024;
function validarFotoBase64_(dataUrl) {
  const m = String(dataUrl || "").match(/^data:(image\/(?:png|jpeg|webp));base64,([A-Za-z0-9+\/=\s]+)$/);
  if (!m) return { ok: false, error: "La foto debe ser PNG, JPG o WEBP." };
  if (m[2].length > FOTO_MAX_BYTES * 1.4) return { ok: false, error: "La foto pesa demasiado (máx. 2 MB)." };
  let bytes; try { bytes = Utilities.base64Decode(m[2].replace(/\s/g, "")); } catch (e) { return { ok: false, error: "Foto inválida." }; }
  if (bytes.length > FOTO_MAX_BYTES) return { ok: false, error: "La foto pesa demasiado (máx. 2 MB)." };
  const b = bytes.map(x => x & 255), mime = m[1];
  const okMagic = (mime === "image/png" && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47) ||
                  (mime === "image/jpeg" && b[0] === 0xFF && b[1] === 0xD8) ||
                  (mime === "image/webp" && b[0] === 0x52 && b[1] === 0x49 && b[8] === 0x57 && b[9] === 0x45);
  if (!okMagic) return { ok: false, error: "El archivo no es una imagen válida." };
  return { ok: true, bytes: bytes, mime: mime, ext: mime === "image/png" ? "png" : mime === "image/webp" ? "webp" : "jpg" };
}

/* ───────── 6) Copia de seguridad semanal del Sheet ─────────
 * Corre UNA vez instalarBackupSemanal() desde el editor (te pedirá permiso de Drive).
 * Crea una copia cada domingo ~3 am en la carpeta "Tapport Backups" y conserva las últimas 8.
 */
function crearBackupSemanal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const it = DriveApp.getFoldersByName("Tapport Backups");
  const carpeta = it.hasNext() ? it.next() : DriveApp.createFolder("Tapport Backups");
  DriveApp.getFileById(ss.getId()).makeCopy("Backup " + ss.getName() + " " + Utilities.formatDate(new Date(), "America/Lima", "yyyy-MM-dd"), carpeta);
  const archivos = []; const fi = carpeta.getFiles();
  while (fi.hasNext()) { const f = fi.next(); archivos.push({ f: f, t: f.getDateCreated().getTime() }); }
  archivos.sort((a, b) => b.t - a.t).slice(8).forEach(x => x.f.setTrashed(true));
}
function instalarBackupSemanal() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "crearBackupSemanal").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("crearBackupSemanal").timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create();
  crearBackupSemanal();
  Logger.log("Backup semanal instalado y primera copia creada ✅");
}
