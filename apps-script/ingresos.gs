/**
 * Tapport — INGRESOS POR PLANES (Starter / Growth / Pro)
 * Archivo NUEVO: crea "Ingresos.gs" en Apps Script y pega TODO esto. Luego:
 *  1) En Code.gs, al INICIO de la función onEdit(e), agrega esta línea:
 *        try { if (ingresosOnEdit_(e)) return; } catch (x) {}
 *  2) Guarda, recarga el Sheet y usa el menú nuevo "💰 Ingresos" (arriba, junto a "Ayuda").
 *     Primera vez: Crear / actualizar hoja Ingresos → Sincronizar clientes → Activar registro automático.
 *     (Google pedirá autorizar permisos una sola vez.)
 *
 * QUÉ CREA
 *  Hoja "Ingresos":  precios editables de cada plan, resumen (ingreso mensual, ganancia anual
 *                    proyectada, total cobrado...) y la tabla de clientes con lista desplegable
 *                    Plan y Estado (Activo / Inactivo). Todo se calcula solo con fórmulas.
 *  Hoja "Registro mensual": una fila por mes con clientes activos, ingreso por plan, ingreso total,
 *                    variación vs el mes anterior, acumulado del año y un gráfico.
 *
 * CONEXIÓN CON LA APP: la hoja "Config" sigue siendo la que usa Tapport. Si en "Ingresos" cambias
 *  el Plan, el Estado o la fecha Vence de un cliente, se copia solo a Config (columnas J y K):
 *   - Inactivo  → PlanVence = ayer (el panel del restaurante muestra el plan como vencido)
 *   - Activo    → si estaba vencido, PlanVence = hoy + 30 días
 */
const ING_HOJA = "Ingresos";
const ING_REG = "Registro mensual";
const ING_FILA_TITULOS = 12;       // fila de encabezados de la tabla de clientes
const ING_PRIMERA = 13;            // primera fila de clientes
const ING_ULTIMA = 1000;
const ING_PLANES = [["starter", 125], ["growth", 175], ["pro", 225]];
const ING_ENC = ["Restaurante", "Plan", "Precio mensual (S/)", "Inicio del plan", "Vence", "Estado", "Ingreso mensual (S/)", "Meses cobrados", "Total cobrado (S/)", "¿Pagó este mes?", "Fecha de pago", "Notas"];

function onOpen() {
  SpreadsheetApp.getUi().createMenu("💰 Ingresos")
    .addItem("1. Crear / actualizar hoja Ingresos", "crearHojaIngresos")
    .addItem("2. Sincronizar clientes desde Config", "sincronizarIngresos")
    .addItem("3. Registrar ingreso de este mes ahora", "registrarIngresoMes")
    .addItem("4. Activar registro automático diario", "activarRegistroAutomatico")
    .addItem("Reiniciar pagos del mes (todos a Pendiente)", "reiniciarPagosMes")
    .addToUi();
}

/** Crea (o repara) la hoja Ingresos con precios, resumen, tabla y formato. No borra tus clientes. */
function crearHojaIngresos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(ING_HOJA);
  const nueva = !h;
  if (!h) h = ss.insertSheet(ING_HOJA, 0);
  h.setHiddenGridlines(true);
  const AZUL = "#0A1530", ACENTO = "#2E6CF6", SUAVE = "#E8F0FE";

  // Título
  h.getRange("A1").setValue("INGRESOS POR PLANES — Tapport").setFontSize(16).setFontWeight("bold").setFontColor(AZUL);
  h.getRange("A2").setValue("Cambia el Plan o el Estado (Activo/Inactivo) de cada cliente y todo se recalcula. Los precios se editan en la tabla de planes.")
    .setFontColor("#5B6270").setFontSize(10);

  // Resumen (filas 3-4)
  const rotulos = ["Ingreso mensual actual", "Clientes activos", "Clientes inactivos", "Ganancia anual proyectada", "Total cobrado histórico", "Ticket promedio"];
  h.getRange(3, 1, 1, 6).setValues([rotulos]).setBackground(AZUL).setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(10);
  const f = ING_PRIMERA, l = ING_ULTIMA;
  h.getRange(4, 1, 1, 6).setFormulas([[
    `=SUM(G${f}:G${l})`,
    `=COUNTIF(F${f}:F${l},"Activo")`,
    `=COUNTIF(F${f}:F${l},"Inactivo")`,
    `=A4*12`,
    `=SUM(I${f}:I${l})`,
    `=IFERROR(A4/B4,0)`
  ]]).setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center").setBackground(SUAVE).setFontColor(AZUL);
  h.getRange(3, 7, 1, 2).setValues([["Cobrado este mes", "Pendiente de cobro"]]).setBackground("#12875B").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(10);
  h.getRange(4, 7, 1, 2).setFormulas([[`=SUMIFS(G${f}:G${l},J${f}:J${l},"Pagado")`, `=A4-G4`]])
    .setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#E7F8F0").setFontColor("#12875B").setNumberFormat('"S/ "#,##0.00');
  h.getRange("A4").setNumberFormat('"S/ "#,##0.00');
  h.getRange("D4:F4").setNumberFormat('"S/ "#,##0.00');
  h.getRange("B4:C4").setNumberFormat("0");
  h.setRowHeight(4, 38);

  // Precios por plan (filas 6-9)
  h.getRange(6, 1, 1, 5).setValues([["Plan", "Precio mensual (S/)", "Clientes activos", "Ingreso mensual (S/)", "% del ingreso"]])
    .setBackground(AZUL).setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(10);
  const actuales = h.getRange(7, 2, 3, 1).getValues();
  ING_PLANES.forEach((p, i) => {
    const r = 7 + i;
    h.getRange(r, 1).setValue(p[0]).setFontWeight("bold");
    h.getRange(r, 2).setValue(nueva || !actuales[i][0] ? p[1] : actuales[i][0]).setBackground("#FFF8E1").setNumberFormat('"S/ "#,##0.00');
    h.getRange(r, 3).setFormula(`=COUNTIFS($B$${f}:$B$${l},A${r},$F$${f}:$F$${l},"Activo")`);
    h.getRange(r, 4).setFormula(`=B${r}*C${r}`).setNumberFormat('"S/ "#,##0.00');
    h.getRange(r, 5).setFormula(`=IFERROR(D${r}/$A$4,0)`).setNumberFormat("0%");
  });
  h.getRange("A10").setValue("Total").setFontWeight("bold");
  h.getRange("C10").setFormula("=SUM(C7:C9)").setFontWeight("bold");
  h.getRange("D10").setFormula("=SUM(D7:D9)").setFontWeight("bold").setNumberFormat('"S/ "#,##0.00');
  h.getRange("A10:E10").setBorder(true, null, null, null, null, null, AZUL, SpreadsheetApp.BorderStyle.SOLID);
  h.getRange("G6").setValue("💡 Amarillo = precios editables. Marca \"Pagado\" cuando el cliente pague: renueva el plan 30 días.").setFontColor("#B7791F").setFontSize(10);
  h.getRange("A7:E10").setHorizontalAlignment("center");

  // Tabla de clientes
  h.getRange(ING_FILA_TITULOS, 1, 1, ING_ENC.length).setValues([ING_ENC])
    .setBackground(AZUL).setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  h.setRowHeight(ING_FILA_TITULOS, 36);
  h.setFrozenRows(ING_FILA_TITULOS);

  const n = l - f + 1;
  const formulas = [];
  for (let r = f; r <= l; r++) {
    formulas.push([
      `=IF(B${r}="","",IFERROR(VLOOKUP(LOWER(B${r}),$A$7:$B$9,2,FALSE),0))`,
    ]);
  }
  h.getRange(f, 3, n, 1).setFormulas(formulas).setNumberFormat('"S/ "#,##0.00');
  const g = [], hh = [], ii = [];
  for (let r = f; r <= l; r++) {
    g.push([`=IF(A${r}="","",IF(F${r}="Activo",C${r},0))`]);
    hh.push([`=IF(OR(A${r}="",D${r}=""),"",IFERROR(IF(F${r}="Activo",DATEDIF(D${r},TODAY(),"M")+1,IF(E${r}<>"",MAX(1,DATEDIF(D${r},E${r},"M")),1)),1))`]);
    ii.push([`=IF(H${r}="","",H${r}*C${r})`]);
  }
  h.getRange(f, 7, n, 1).setFormulas(g).setNumberFormat('"S/ "#,##0.00');
  h.getRange(f, 8, n, 1).setFormulas(hh).setNumberFormat("0");
  h.getRange(f, 9, n, 1).setFormulas(ii).setNumberFormat('"S/ "#,##0.00');
  h.getRange(f, 4, n, 2).setNumberFormat("dd/mm/yyyy");
  h.getRange(f, 2, n, 8).setHorizontalAlignment("center");
  h.getRange(f, 1, n, 1).setHorizontalAlignment("left");

  // Listas desplegables
  h.getRange(f, 2, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["starter", "growth", "pro"], true).setAllowInvalid(false).build());
  h.getRange(f, 6, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["Activo", "Inactivo"], true).setAllowInvalid(false).build());

  h.getRange(f, 10, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["Pagado", "Pendiente"], true).setAllowInvalid(false).build());
  h.getRange(f, 11, n, 1).setNumberFormat("dd/mm/yyyy");
  h.getRange(f, 10, n, 2).setHorizontalAlignment("center");
  // Colores por estado y plan
  const rango = h.getRange(f, 1, n, ING_ENC.length);
  h.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$F${f}="Activo"`).setBackground("#E7F8F0").setRanges([h.getRange(f, 6, n, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$J${f}="Pagado"`).setBackground("#D7F5E6").setFontColor("#12875B").setRanges([h.getRange(f, 10, n, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND($A${f}<>"",$F${f}="Activo",$J${f}<>"Pagado")`).setBackground("#FFF3D6").setFontColor("#B7791F").setRanges([h.getRange(f, 10, n, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$F${f}="Inactivo"`).setBackground("#FDE8E6").setFontColor("#B4432E").setRanges([rango]).build()
  ]);

  const anchos = [190, 100, 130, 120, 120, 100, 150, 130, 140, 130, 120, 220];
  anchos.forEach((w, i) => h.setColumnWidth(i + 1, w));
  if (nueva) { try { sincronizarIngresos(); } catch (e) { } }
  SpreadsheetApp.getActive().toast("Hoja Ingresos lista ✅", "Tapport", 4);
}

function ingNorm_(s) { return String(s || "").trim().toLowerCase(); }
function ingFecha_(v) { if (v instanceof Date) return v; const d = v ? new Date(v) : null; return d && !isNaN(d) ? d : null; }

/** Trae a "Ingresos" los restaurantes de Config que todavía no estén (con su plan y vencimiento). */
function sincronizarIngresos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(ING_HOJA); if (!h) { crearHojaIngresos(); return; }
  const cfg = ss.getSheetByName(HOJA_CONFIG); if (!cfg) return;
  const datos = cfg.getDataRange().getValues();
  const existentes = h.getRange(ING_PRIMERA, 1, ING_ULTIMA - ING_PRIMERA + 1, 1).getValues().map(r => ingNorm_(r[0]));
  let sig = existentes.findIndex(x => !x); sig = sig < 0 ? existentes.length : sig;
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  let nuevos = 0;
  for (let i = 1; i < datos.length; i++) {
    const nombre = String(datos[i][0] || "").trim(); if (!nombre) continue;
    if (existentes.indexOf(ingNorm_(nombre)) >= 0) continue;
    const plan = ingNorm_(datos[i][COL_PLAN - 1]); const venceD = ingFecha_(datos[i][COL_PLAN_VENCE - 1]);
    const activo = !venceD || venceD.getTime() >= Date.now() - 24 * 3600 * 1000;
    const r = ING_PRIMERA + sig + nuevos;
    h.getRange(r, 1).setValue(nombre);
    h.getRange(r, 2).setValue(["starter", "growth", "pro"].indexOf(plan) >= 0 ? plan : "starter");
    h.getRange(r, 4).setValue(hoy);
    if (venceD) h.getRange(r, 5).setValue(venceD);
    h.getRange(r, 6).setValue(activo ? "Activo" : "Inactivo");
    h.getRange(r, 10).setValue("Pendiente");
    existentes.push(ingNorm_(nombre)); nuevos++;
  }
  SpreadsheetApp.getActive().toast(nuevos ? `Se agregaron ${nuevos} cliente(s). Revisa Plan e "Inicio del plan".` : "Todos los clientes de Config ya están en Ingresos.", "Tapport", 5);
}

/** Si cambias Plan / Vence / Estado en "Ingresos", se copia a Config. Devuelve true si fue en esa hoja. */
function ingresosOnEdit_(e) {
  const h = e.range.getSheet(); if (h.getName() !== ING_HOJA) return false;
  const fila = e.range.getRow(), col = e.range.getColumn();
  if (fila < ING_PRIMERA || (col !== 2 && col !== 5 && col !== 6 && col !== 10)) return true;
  const nombre = ingNorm_(h.getRange(fila, 1).getValue()); if (!nombre) return true;
  const cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONFIG); if (!cfg) return true;
  const datos = cfg.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (ingNorm_(datos[i][0]) !== nombre) continue;
    const rf = i + 1;
    if (col === 2) cfg.getRange(rf, COL_PLAN).setValue(ingNorm_(h.getRange(fila, 2).getValue())).setBackground(null);
    if (col === 5) { const v = h.getRange(fila, 5).getValue(); cfg.getRange(rf, COL_PLAN_VENCE).setValue(v || ""); }
    if (col === 10 && h.getRange(fila, 10).getValue() === "Pagado" && !h.getRange(fila, 11).getValue()) {
      // Pagó: queda Activo, se registra la fecha y el plan se renueva 30 días (desde hoy o desde su vencimiento si aún está vigente).
      const ahora = new Date(); ahora.setHours(0, 0, 0, 0);
      const v = ingFecha_(cfg.getRange(rf, COL_PLAN_VENCE).getValue());
      const base = v && v.getTime() > ahora.getTime() ? v : ahora;
      const nuevo = new Date(base.getTime() + 30 * 24 * 3600 * 1000);
      cfg.getRange(rf, COL_PLAN_VENCE).setValue(nuevo);
      h.getRange(fila, 5).setValue(nuevo); h.getRange(fila, 6).setValue("Activo"); h.getRange(fila, 11).setValue(ahora);
    }
    if (col === 6) {
      const est = h.getRange(fila, 6).getValue(), ahora = new Date();
      if (est === "Inactivo") {
        const ayer = new Date(ahora.getTime() - 2 * 24 * 3600 * 1000); ayer.setHours(0, 0, 0, 0);
        cfg.getRange(rf, COL_PLAN_VENCE).setValue(ayer); h.getRange(fila, 5).setValue(ayer);
      } else if (est === "Activo") {
        const v = ingFecha_(cfg.getRange(rf, COL_PLAN_VENCE).getValue());
        if (v && v.getTime() < Date.now() - 24 * 3600 * 1000) {
          const nuevo = new Date(ahora.getTime() + 30 * 24 * 3600 * 1000); nuevo.setHours(0, 0, 0, 0);
          cfg.getRange(rf, COL_PLAN_VENCE).setValue(nuevo); h.getRange(fila, 5).setValue(nuevo);
        }
      }
    }
    break;
  }
  return true;
}

/** Guarda (o actualiza) la fila del mes actual en "Registro mensual". */
function registrarIngresoMes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(ING_HOJA); if (!h) { crearHojaIngresos(); return; }
  const filas = h.getRange(ING_PRIMERA, 1, ING_ULTIMA - ING_PRIMERA + 1, 6).getValues();
  const precios = {}; h.getRange("A7:B9").getValues().forEach(r => precios[ingNorm_(r[0])] = Number(r[1]) || 0);
  const c = { starter: 0, growth: 0, pro: 0 };
  filas.forEach(r => { if (r[0] && r[5] === "Activo" && c[ingNorm_(r[1])] !== undefined) c[ingNorm_(r[1])]++; });
  let cobrado = 0;
  const pagos = h.getRange(ING_PRIMERA, 10, ING_ULTIMA - ING_PRIMERA + 1, 1).getValues();
  filas.forEach((r, i) => { if (r[0] && r[5] === "Activo" && pagos[i][0] === "Pagado") cobrado += Number(precios[ingNorm_(r[1])]) || 0; });
  const activos = c.starter + c.growth + c.pro;
  const ingreso = c.starter * precios.starter + c.growth * precios.growth + c.pro * precios.pro;
  const mes = Utilities.formatDate(new Date(), "America/Lima", "yyyy-MM");

  let reg = ss.getSheetByName(ING_REG);
  if (!reg) {
    reg = ss.insertSheet(ING_REG);
    reg.getRange(1, 1, 1, 9).setValues([["Mes", "Clientes activos", "Starter", "Growth", "Pro", "Ingreso del mes (S/)", "Variación vs mes anterior", "Acumulado del año (S/)", "Última actualización"]])
      .setBackground("#0A1530").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setWrap(true);
    reg.setFrozenRows(1); reg.setRowHeight(1, 36);
    reg.getRange("A2:A200").setNumberFormat("@");
    [90, 110, 80, 80, 80, 150, 150, 160, 150].forEach((w, i) => reg.setColumnWidth(i + 1, w));
  }
  const meses = reg.getRange(2, 1, Math.max(1, reg.getLastRow() - 1), 1).getValues().map(r => String(r[0]));
  let idx = meses.indexOf(mes); const fila = idx >= 0 ? idx + 2 : (reg.getLastRow() < 2 ? 2 : reg.getLastRow() + 1);
  reg.getRange(fila, 1).setNumberFormat("@").setValue(mes);
  reg.getRange(fila, 2, 1, 5).setValues([[activos, c.starter, c.growth, c.pro, ingreso]]);
  reg.getRange(1, 10, 1, 2).setValues([["Cobrado (S/)", "Pendiente (S/)"]]).setBackground("#0A1530").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  reg.getRange(fila, 10, 1, 2).setValues([[cobrado, ingreso - cobrado]]).setNumberFormat('"S/ "#,##0.00').setHorizontalAlignment("center");
  reg.setColumnWidth(10, 120); reg.setColumnWidth(11, 120);
  reg.getRange(fila, 7).setFormula(fila === 2 ? '=""' : `=IFERROR(F${fila}/F${fila - 1}-1,"")`).setNumberFormat("+0.0%;-0.0%;0.0%");
  reg.getRange(fila, 8).setFormula(`=SUMIF($A$2:$A$200,LEFT(A${fila},4)&"*",$F$2:$F$200)`);
  reg.getRange(fila, 9).setValue(new Date()).setNumberFormat("dd/mm/yyyy hh:mm");
  reg.getRange(fila, 6).setNumberFormat('"S/ "#,##0.00'); reg.getRange(fila, 8).setNumberFormat('"S/ "#,##0.00');
  reg.getRange(fila, 1, 1, 9).setHorizontalAlignment("center");
  reg.getRange(2, 1, Math.max(1, reg.getLastRow() - 1), 11).sort({ column: 1, ascending: true });
  // Gráfico (se crea una sola vez)
  if (reg.getCharts().length === 0) {
    reg.insertChart(reg.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(reg.getRange("A1:A200")).addRange(reg.getRange("F1:F200"))
      .setPosition(2, 11, 0, 0).setNumHeaders(1)
      .setOption("title", "Ingreso mensual por planes (S/)").setOption("legend", { position: "none" })
      .setOption("colors", ["#2E6CF6"]).setOption("width", 520).setOption("height", 300).build());
  }
  SpreadsheetApp.getActive().toast(`Registrado ${mes}: S/ ${ingreso} con ${activos} cliente(s) activo(s).`, "Tapport", 5);
}

/** Corre registrarIngresoMes todos los días a las 11 pm: el último valor del mes queda como cierre. */
function activarRegistroAutomatico() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "registrarIngresoMes") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("registrarIngresoMes").timeBased().everyDays(1).atHour(23).create();
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "reiniciarPagosMes") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("reiniciarPagosMes").timeBased().onMonthDay(1).atHour(1).create();
  registrarIngresoMes();
  SpreadsheetApp.getActive().toast("Registro automático activado (diario 11 pm) y pagos se reinician el día 1 de cada mes.", "Tapport", 6);
}

/** Día 1 de cada mes: todos los clientes pasan a "Pendiente" y se borra la fecha de pago (el mes anterior ya quedó guardado). */
function reiniciarPagosMes() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ING_HOJA); if (!h) return;
  const n = ING_ULTIMA - ING_PRIMERA + 1;
  const nombres = h.getRange(ING_PRIMERA, 1, n, 1).getValues();
  const pago = nombres.map(r => [r[0] ? "Pendiente" : ""]), fecha = nombres.map(() => [""]);
  h.getRange(ING_PRIMERA, 10, n, 1).setValues(pago);
  h.getRange(ING_PRIMERA, 11, n, 1).setValues(fecha);
}
