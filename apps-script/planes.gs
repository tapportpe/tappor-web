/**
 * Tapport — PLANES (Starter / Growth / Pro)
 * Archivo NUEVO: en Apps Script crea un archivo "Planes.gs" y pega TODO esto.
 * Luego haz los 5 cambios pequeños que están al final de este archivo.
 *
 * Dónde se guarda el plan de cada restaurante:
 *   Hoja "Config" → columna J = Plan (starter | growth | pro)
 *                   columna K = PlanVence (fecha, ej. 2026-11-30; vacío = sin vencimiento)
 * Tú los llenas a mano cuando el restaurante te paga (Yape/transferencia).
 * Si J está vacía, el restaurante cuenta como "starter".
 */

// Qué incluye cada plan. Si cambias algo, cámbialo solo aquí.
const PLANES = {
  starter: { nombre: "Starter", maxMeseros: 5,   ia: false },
  growth:  { nombre: "Growth",  maxMeseros: 10,  ia: true  },
  pro:     { nombre: "Pro",     maxMeseros: 999, ia: true  }
};

const COL_PLAN = 10;        // J
const COL_PLAN_VENCE = 11;  // K

/** Devuelve el plan vigente de un restaurante: { plan, nombre, vence, activo, maxMeseros, ia } */
function planDeRestaurante(restaurante) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONFIG);
  let clave = "starter", vence = null;
  if (hoja) {
    const datos = hoja.getDataRange().getValues();
    const norm = String(restaurante || "").trim().toLowerCase();
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0] || "").trim().toLowerCase() === norm) {
        const p = String(datos[i][COL_PLAN - 1] || "").trim().toLowerCase();
        if (PLANES[p]) clave = p;
        vence = parsearFecha_(datos[i][COL_PLAN_VENCE - 1]);
        break;
      }
    }
  }
  const def = PLANES[clave];
  const activo = !vence || vence.getTime() >= Date.now() - 24 * 3600 * 1000; // un día de gracia
  return {
    plan: clave,
    nombre: def.nombre,
    vence: vence ? vence.toISOString() : "",
    activo: activo,
    maxMeseros: def.maxMeseros,
    // Si el plan venció, se bloquean las funciones de pago hasta que renueve.
    ia: def.ia && activo
  };
}

/** Acepta una fecha real de la hoja o texto como "18/11/26", "18/11/2026" o "2026-11-18". Devuelve Date o null. */
function parsearFecha_(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  const s = String(v || "").trim();
  if (!s) return null;
  const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m) {
    let y = Number(m[3]); if (y < 100) y += 2000;
    const d = new Date(y, Number(m[2]) - 1, Number(m[1]), 23, 59, 59); // vence al final de ese día
    return isNaN(d) ? null : d;
  }
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

/** Cuántos meseros tiene hoy un restaurante. */
function contarMeseros(restaurante) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TRABAJADORES);
  const datos = hoja.getDataRange().getValues();
  const norm = String(restaurante || "").trim().toLowerCase();
  let n = 0;
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][1] || "").trim().toLowerCase() === norm) n++;
  }
  return n;
}

/** Lo que el panel necesita saber (se añade a login_gerente y panel_gerente). */
function infoPlan(restaurante) {
  const p = planDeRestaurante(restaurante);
  return {
    plan: p.plan,
    planNombre: p.nombre,
    planVence: p.vence,
    planActivo: p.activo,
    maxMeseros: p.maxMeseros,
    funciones: { ia: p.ia }
  };
}

/** Lista desplegable con los 3 planes (se usa en toda la columna Plan). */
function reglaPlanes_() {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(["starter", "growth", "pro"], true)
    .setAllowInvalid(false)
    .setHelpText("Elige el plan del restaurante: starter, growth o pro.")
    .build();
}

/**
 * Deja lista la fila de un restaurante en "Config": lista de planes en J, formato de fecha en K,
 * cupo (L), meseros usados (M) y total de gente (E). Si todavía no tiene plan, pinta J de amarillo
 * para recordarte que falta elegirlo (se quita solo cuando eliges uno).
 * Se llama sola cuando creas un restaurante nuevo (desde onEdit y obtenerOAsignarPrefijo).
 */
function prepararFilaPlan(hoja, fila) {
  const celda = hoja.getRange(fila, COL_PLAN);
  celda.setDataValidation(reglaPlanes_());
  hoja.getRange(fila, COL_PLAN_VENCE).setNumberFormat("dd/MM/yyyy");
  hoja.getRange(fila, 5).setFormula("=C" + fila + "+D" + fila);
  hoja.getRange(fila, 12).setFormula('=IF(J' + fila + '="pro","Sin límite",IF(J' + fila + '="growth",10,5))');
  hoja.getRange(fila, 13).setFormula('=C' + fila + '&" de "&L' + fila);
  if (!String(celda.getValue() || "").trim()) celda.setBackground("#FFF4CC");
}

/** Diagnóstico: ejecútala desde el editor y mira el Registro de ejecución. Cambia "Tapport" si quieres probar otro restaurante. */
function probarPlan() {
  Logger.log("1) Plan que lee de la hoja Config: " + JSON.stringify(infoPlan("Tapport")));
  Logger.log("2) ¿Code.gs ya manda el plan en el login? " + (loginGerente.toString().indexOf("infoPlan") > -1 ? "SÍ" : "NO → falta pegar el Code.gs nuevo"));
  Logger.log("3) ¿registrarMesero ya tiene el límite? " + (registrarMesero.toString().indexOf("planDeRestaurante") > -1 ? "SÍ" : "NO → falta pegar el Code.gs nuevo"));
}

/** Ejecuta UNA vez desde el editor: crea los encabezados J y K en "Config". */
function prepararColumnasPlan() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONFIG);
  hoja.getRange(1, COL_PLAN).setValue("Plan");
  hoja.getRange(1, COL_PLAN_VENCE).setValue("PlanVence");
  hoja.getRange(1, 12).setValue("Cupo del plan");   // L
  hoja.getRange(1, 13).setValue("Meseros usados");  // M  (ej. "3 de 5")
  const ultima = hoja.getLastRow();
  for (let f = 2; f <= ultima; f++) {
    if (!String(hoja.getRange(f, 1).getValue() || "").trim()) continue;
    hoja.getRange(f, 12).setFormula('=IF(J' + f + '="pro","Sin límite",IF(J' + f + '="growth",10,5))');
    hoja.getRange(f, 13).setFormula('=C' + f + '&" de "&L' + f);
  }
  Logger.log("Listo. Escribe starter / growth / pro en la columna J de cada restaurante. C cuenta los meseros solos; L y M muestran el cupo.");
}

/**
 * Ordena la hoja "Config" (ejecútala UNA vez desde el editor; es segura repetirla):
 *  - Pone los encabezados que faltan en F:M (Direccion, LogoUrl, PortadaUrl, Color, Plan,
 *    PlanVence, Cupo del plan, Meseros usados) con el mismo estilo que la fila 1.
 *  - Borra colores (#RRGGBB) que quedaron sueltos en columnas que no son "Color"
 *    (restos de migraciones viejas que rompían logo/portada/plan).
 *  - Pone la lista de planes (starter/growth/pro) en TODA la columna Plan y pinta de amarillo los que aún no tienen plan.
 *  - Formatea la fecha de vencimiento, congela la fila 1 y ajusta los anchos.
 * No toca A:E ni los links de logo/portada/dirección.
 */
function organizarConfig() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONFIG);
  const ult = Math.max(hoja.getLastRow(), 2);

  prepararColumnasPlan(); // encabezados Plan/PlanVence/Cupo/Meseros usados + fórmulas L y M

  hoja.getRange(1, 6, 1, 4).setValues([["Direccion", "LogoUrl", "PortadaUrl", "Color"]]);
  hoja.getRange(1, 1).copyTo(hoja.getRange(1, 6, 1, 8), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  const esColor = function (v) { return /^#[0-9a-f]{6}$/i.test(String(v || "").trim()); };
  let limpiadas = 0;
  [6, 7, 8, 10, 11].forEach(function (col) {   // F, G, H, J, K nunca deben tener un color
    for (let f = 2; f <= ult; f++) {
      const celda = hoja.getRange(f, col);
      if (esColor(celda.getValue())) { celda.clearContent(); limpiadas++; }
    }
  });

  hoja.getRange(2, 10, ult - 1, 4).setBackground(null); // J:M sin el relleno de color suelto

  // La lista de planes queda en TODA la columna J, así cualquier restaurante nuevo ya la trae.
  const filasTotal = hoja.getMaxRows() - 1;
  hoja.getRange(2, COL_PLAN, filasTotal, 1).setDataValidation(reglaPlanes_());
  let sinPlan = 0;
  for (let f = 2; f <= ult; f++) {
    if (!String(hoja.getRange(f, 1).getValue() || "").trim()) continue;
    const c = hoja.getRange(f, COL_PLAN);
    if (!String(c.getValue() || "").trim()) { c.setBackground("#FFF4CC"); sinPlan++; }
  }

  hoja.getRange(2, COL_PLAN_VENCE, filasTotal, 1).setNumberFormat("dd/MM/yyyy");

  // E = C + D (total de gente), y cada fila con el color de su restaurante (columna I)
  for (let f = 2; f <= ult; f++) {
    if (!String(hoja.getRange(f, 1).getValue() || "").trim()) continue;
    hoja.getRange(f, 5).setFormula("=C" + f + "+D" + f);
    const color = String(hoja.getRange(f, 9).getValue() || "").trim();
    if (esColor(color)) hoja.getRange(f, 1, 1, 9).setBackground(color);
  }

  // Quita las columnas vacías que sobran a la derecha (todo lo que pase de la M)
  if (hoja.getMaxColumns() > 13) hoja.deleteColumns(14, hoja.getMaxColumns() - 13);

  // Notas en los encabezados: para qué sirve cada columna
  const notas = [
    "Nombre del restaurante (lo escribes tú; el prefijo se crea solo).",
    "Prefijo de los códigos de sus meseros (se crea solo).",
    "Meseros registrados en este restaurante (se cuenta solo).",
    "Cuentas de gerente de este restaurante (se cuenta solo).",
    "Meseros + gerentes (se calcula solo).",
    "Dirección del local. La edita el gerente desde su panel.",
    "Link del logo. Lo sube el gerente desde su panel.",
    "Link de la portada. La sube el gerente desde su panel.",
    "Color del restaurante en las hojas (se asigna solo).",
    "Plan contratado: starter, growth o pro. Elígelo de la lista.",
    "Fecha hasta la que está pagado el plan. Vacío = sin vencimiento.",
    "Meseros que permite el plan (se calcula solo).",
    "Meseros registrados / cupo del plan (se calcula solo)."
  ];
  hoja.getRange(1, 1, 1, 13).setNotes([notas]);
  hoja.getRange(1, 1, 1, 13).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  hoja.getRange(2, 3, ult - 1, 3).setHorizontalAlignment("center");   // C:E
  hoja.getRange(2, 9, ult - 1, 5).setHorizontalAlignment("center");   // I:M
  hoja.setRowHeight(1, 42);

  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, 13);
  hoja.setColumnWidth(3, 130); hoja.setColumnWidth(4, 140); hoja.setColumnWidth(5, 130);
  hoja.setColumnWidth(6, 180); hoja.setColumnWidth(7, 140); hoja.setColumnWidth(8, 140);

  // Repinta los meseros y las reseñas con el color de su restaurante (funciones que ya tienes)
  try { colorearFilasExistentes(); reordenarYColorearResenas(); } catch (e) { Logger.log("Colores de Meseros/Reseñas: " + e); }

  Logger.log("Config ordenada. Colores sueltos borrados: " + limpiadas + ". Restaurantes sin plan elegido (en amarillo): " + sinPlan + ". Elige el plan en la lista de la columna J.");
}

/* ==================================================================
 * CAMBIOS QUE DEBES HACER EN TU Code.gs (son 5, muy cortos)
 * ==================================================================
 *
 * 1) En loginGerente(), dentro del respuesta({...}) de éxito, agrega al final
 *    (después de portadaUrl: negocio.portadaUrl) :
 *
 *        portadaUrl: negocio.portadaUrl,
 *        ...infoPlan(restaurante)
 *
 * 2) En panelGerente(), igual: después de portadaUrl: negocio.portadaUrl agrega
 *
 *        portadaUrl: negocio.portadaUrl,
 *        ...infoPlan(restaurante)
 *
 * 3) En registrarMesero(), justo DESPUÉS de la línea
 *        const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TRABAJADORES);
 *    agrega:
 *
 *        const planR = planDeRestaurante(restaurante);
 *        if (contarMeseros(restaurante) >= planR.maxMeseros) {
 *          return respuesta({ ok: false, error: "Este restaurante alcanzó el límite de meseros de su plan (" + planR.maxMeseros + "). Pídele a tu gerente que suba de plan." });
 *        }
 *
 * 4) En ia-asistente.gs (IA.gs), dentro de iaAsistente_(), justo DESPUÉS de
 *        if (!login || !login.ok) return iaJson_({ ok: false, error: "Sesión no válida." });
 *    agrega:
 *
 *        if (!login.funciones || !login.funciones.ia) {
 *          return iaJson_({ ok: false, error: "Tu plan no incluye el asistente IA. Disponible desde el plan Growth." });
 *        }
 *
 * 5) Ejecuta una vez prepararColumnasPlan() y escribe el plan en la columna J
 *    de cada restaurante (growth para tu piloto, por ejemplo).
 *
 * Después: Implementar > Administrar implementaciones > Editar > Nueva versión.
 */
