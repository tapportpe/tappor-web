/**
 * Tapport — Reparar y ordenar la hoja "Meseros".
 * Archivo NUEVO: crea "Reparar.gs" en Apps Script y pega esto. No cambia Code.gs.
 *
 * Qué pasó: en la hoja Meseros la columna J (PasswordHash) quedó con un enlace en vez
 * de la contraseña cifrada, y K (Salt) sin encabezado. Con eso el login nunca coincide.
 * La contraseña original no se puede recuperar (solo se guarda cifrada), así que se
 * establece una nueva.
 *
 * ORDEN DE USO:
 *  1) Cambia EMAIL_A_RESTABLECER y PASSWORD_NUEVA (abajo) y ejecuta  restablecerPasswordMesero
 *  2) Ejecuta  repararMeseros   (revisa todas las filas y arregla encabezados/enlaces/hash)
 *  3) Ejecuta  ordenarHojaMeseros   (formato, orden y columnas técnicas ocultas)
 *  4) BORRA la contraseña de la constante PASSWORD_NUEVA y guarda.
 *
 * Columnas de Meseros: A Codigo | B Restaurante | C Nombre | D Telefono | E Edad | F DNI |
 *                      G Email | H FotoUrl | I Link | J PasswordHash | K Salt
 */
const EMAIL_A_RESTABLECER = "santiagocajoar@gmail.com";
const PASSWORD_NUEVA = "";   // <- escribe aquí tu nueva contraseña (mín. 6 caracteres), ejecuta y luego bórrala

const REP_ENCABEZADOS = ["Codigo", "Restaurante", "Nombre", "Telefono", "Edad", "DNI", "Email", "FotoUrl", "Link", "PasswordHash", "Salt"];

function repHash_(h) { return /^[0-9a-f]{64}$/.test(String(h || "")); }

/** 1) Pone una contraseña nueva a un mesero (hash + salt correctos). */
function restablecerPasswordMesero() {
  if (String(PASSWORD_NUEVA).length < 6) { Logger.log("Escribe PASSWORD_NUEVA (mínimo 6 caracteres) arriba en el código y vuelve a ejecutar."); return; }
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TRABAJADORES);
  const datos = hoja.getDataRange().getValues();
  const email = EMAIL_A_RESTABLECER.trim().toLowerCase();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][6] || "").trim().toLowerCase() === email) {
      const salt = generarSalt();
      hoja.getRange(i + 1, 10).setValue(hashPassword(PASSWORD_NUEVA, salt));
      hoja.getRange(i + 1, 11).setValue(salt);
      Logger.log("Listo: contraseña nueva guardada para " + email + " (fila " + (i + 1) + "). Ya puedes iniciar sesión. Ahora borra PASSWORD_NUEVA del código.");
      return;
    }
  }
  Logger.log("No encontré a " + email + " en la columna G (Email) de la hoja Meseros.");
}

/** 2) Revisa todas las filas: encabezados, enlace y hash. No borra datos útiles. */
function repararMeseros() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TRABAJADORES);
  hoja.getRange(1, 1, 1, REP_ENCABEZADOS.length).setValues([REP_ENCABEZADOS]);
  const datos = hoja.getDataRange().getValues();
  let arreglados = 0;
  for (let i = 1; i < datos.length; i++) {
    const f = datos[i], fila = i + 1;
    if (!f[0]) continue;
    // Link (I): si no es un enlace, se reconstruye con el código
    if (!/^https?:\/\//i.test(String(f[8] || ""))) {
      hoja.getRange(fila, 9).setValue(LANDING_BASE_URL + "?code=" + f[0]);
      Logger.log("Fila " + fila + ": enlace (I) reconstruido.");
      arreglados++;
    }
    // PasswordHash (J): debe ser 64 caracteres hexadecimales. Si no lo es, se vacía y
    // el mesero crea su contraseña de nuevo con 'Crear mi cuenta'.
    const hash = String(f[9] || "");
    if (hash && !repHash_(hash)) {
      hoja.getRange(fila, 10, 1, 2).clearContent();
      Logger.log("Fila " + fila + " (" + f[2] + "): J no era una contraseña cifrada; se vació. Debe crear su contraseña (usa restablecerPasswordMesero o 'Crear mi cuenta').");
      arreglados++;
    } else if (hash && !String(f[10] || "")) {
      hoja.getRange(fila, 10).clearContent();
      Logger.log("Fila " + fila + " (" + f[2] + "): tenía hash sin salt; se vació para que cree su contraseña.");
      arreglados++;
    }
  }
  // Celdas sueltas a la derecha de K (por ejemplo un salt duplicado en L): la hoja solo usa A–K
  if (hoja.getLastColumn() > REP_ENCABEZADOS.length && hoja.getLastRow() > 0) {
    hoja.getRange(1, REP_ENCABEZADOS.length + 1, hoja.getLastRow(), hoja.getLastColumn() - REP_ENCABEZADOS.length).clearContent();
    Logger.log("Se limpiaron celdas sueltas a la derecha de la columna K.");
  }
  Logger.log("Revisión terminada. Correcciones: " + arreglados);
}

/** 3) Deja la hoja Meseros con el mismo estilo que Administradores y ordenada por restaurante. */
function ordenarHojaMeseros() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TRABAJADORES);
  const n = REP_ENCABEZADOS.length, filas = hoja.getLastRow();

  // Quita ajustes de la versión anterior (columnas ocultas / congeladas)
  hoja.showColumns(1, n);
  hoja.setFrozenColumns(0);

  // Encabezado igual al de Administradores: fondo azul oscuro, texto blanco centrado
  hoja.getRange(1, 1, 1, n).setValues([REP_ENCABEZADOS])
    .setBackground("#0A1530").setFontColor("#FFFFFF").setFontWeight("normal")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  hoja.setFrozenRows(1);

  if (filas > 1) {
    const cuerpo = hoja.getRange(2, 1, filas - 1, n);
    cuerpo.setHorizontalAlignment("left").setVerticalAlignment("middle").setFontColor("#000000").setFontSize(10).setFontWeight("normal");
    hoja.getRange(2, 4, filas - 1, 3).setNumberFormat("@"); // Telefono, Edad, DNI como texto
    hoja.getRange(2, 8, filas - 1, 4).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // FotoUrl, Link, Hash, Salt
    cuerpo.sort([{ column: 2, ascending: true }, { column: 1, ascending: true }]);
  }

  const anchos = [95, 150, 170, 100, 55, 95, 230, 150, 150, 170, 150];
  anchos.forEach(function (w, i) { hoja.setColumnWidth(i + 1, w); });

  try { colorearFilasExistentes(); } catch (e) { Logger.log("No se pudo recolorear: " + e); }
  Logger.log("Hoja Meseros ordenada con el estilo de Administradores.");
}
