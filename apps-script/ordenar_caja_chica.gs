/**
 * Configura el sheet Caja_chica para que el saldo de caja se calcule solo.
 *
 * Cómo usarlo (una sola vez):
 *   1. En el sheet: Extensiones → Apps Script.
 *   2. Botón "+" → Secuencia de comandos → nombre: ordenar_caja_chica
 *   3. Pegar este archivo completo y guardar (Ctrl+S).
 *   4. Elegir la función configurarCajaChica en el desplegable y tocar "Ejecutar".
 *      La primera vez Google pide autorizar el script: es tu propio script.
 *
 * Qué hace:
 *   - Movimientos: encabezado fijo, formatos uniformes, Tipo con desplegable
 *     (Ingreso / Egreso), Fecha solo fechas válidas, Monto solo números > 0.
 *   - Toma el saldo actual de Saldo!A2 y lo carga en Movimientos como
 *     "Saldo inicial" (Ingreso). Si ya había movimientos, el saldo inicial se
 *     calcula para que el resultado siga siendo el saldo actual.
 *   - Reemplaza Saldo!A2 por la fórmula Ingresos − Egresos, C2 por la fecha del
 *     último movimiento, y protege A2:C2 para que solo el dueño pueda editarlas.
 *
 * Se puede volver a ejecutar sin problema: si ya existe el "Saldo inicial" no
 * lo duplica, y las validaciones se vuelven a aplicar igual.
 * No borra ni cambia montos, fechas ni descripciones existentes.
 */

const HOJA_MOV = 'Movimientos';
const HOJA_SALDO = 'Saldo';
const FILAS = 1000; // filas preparadas para carga

function configurarCajaChica() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mov = ss.getSheetByName(HOJA_MOV);
  const saldo = ss.getSheetByName(HOJA_SALDO);
  if (!mov || !saldo) throw new Error('Faltan las hojas ' + HOJA_MOV + ' y/o ' + HOJA_SALDO);

  const corregidos = ordenarMovimientos_(mov);
  const inicial = cargarSaldoInicial_(mov, saldo);
  activarFormulaSaldo_(saldo);

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Caja chica configurada.\n\n' +
    'Tipos corregidos: ' + corregidos + '\n' +
    inicial + '\n' +
    'Saldo de caja ahora: ' + saldo.getRange('A2').getDisplayValue() + ' (se calcula solo).');
}

/** Formatos y validaciones de Movimientos. Devuelve cuántos Tipo corrigió. */
function ordenarMovimientos_(sh) {
  if (sh.getMaxRows() < FILAS + 1) sh.insertRowsAfter(sh.getMaxRows(), FILAS + 1 - sh.getMaxRows());

  sh.getRange('A1:D1').setValues([['Fecha', 'Tipo', 'Descripcion', 'Monto']]);
  sh.setFrozenRows(1);

  let corregidos = 0;
  const ultima = sh.getLastRow();
  if (ultima >= 2) {
    const rangoTipo = sh.getRange(2, 2, ultima - 1, 1);
    rangoTipo.setValues(rangoTipo.getValues().map(([t]) => {
      const n = String(t).trim().toLowerCase();
      const nuevo = n === 'egreso' ? 'Egreso' : n === 'ingreso' ? 'Ingreso' : t;
      if (nuevo !== t) corregidos++;
      return [nuevo];
    }));
  }

  sh.getRange(2, 1, FILAS, 4)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setFontFamily('Arial').setFontSize(10);
  sh.getRange(2, 1, FILAS, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(2, 4, FILAS, 1).setNumberFormat('"$"#,##0.00');

  sh.getRange(2, 1, FILAS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .setHelpText('Ingresá una fecha válida (dd/mm/aaaa). Doble clic abre el calendario.')
      .build());

  // Desplegable: muestra la flecha en la celda y solo acepta estas dos opciones
  sh.getRange(2, 2, FILAS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Ingreso', 'Egreso'], true)
      .setAllowInvalid(false)
      .setHelpText('Elegí Ingreso (entra plata a la caja) o Egreso (sale plata).')
      .build());

  sh.getRange(2, 4, FILAS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThan(0)
      .setAllowInvalid(false)
      .setHelpText('Solo el número, mayor a 0, sin signo. El Tipo indica si entra o sale.')
      .build());

  return corregidos;
}

/** Inserta la fila "Saldo inicial" si no existe. Devuelve un texto para el resumen. */
function cargarSaldoInicial_(mov, saldo) {
  const ultima = mov.getLastRow();
  const filas = ultima >= 2 ? mov.getRange(2, 1, ultima - 1, 4).getValues() : [];

  if (filas.some(([, tipo, desc]) => String(tipo).toLowerCase() === 'ingreso'
                                  && String(desc).toLowerCase().includes('saldo inicial'))) {
    return 'Saldo inicial: ya existía, no se modificó.';
  }

  const actual = leerMonto_(saldo.getRange('A2').getValue());
  if (isNaN(actual)) throw new Error('No se pudo leer el saldo actual de Saldo!A2: ' + saldo.getRange('A2').getDisplayValue());

  // Si ya hay movimientos cargados, el saldo inicial es lo que había antes de ellos
  let neto = 0;
  let primera = null;
  filas.forEach(([fecha, tipo, , monto]) => {
    const m = leerMonto_(monto) || 0;
    const t = String(tipo).toLowerCase();
    if (t === 'ingreso') neto += m; else if (t === 'egreso') neto -= m;
    if (fecha instanceof Date && (!primera || fecha < primera)) primera = fecha;
  });
  const montoInicial = actual - neto;
  if (montoInicial <= 0) {
    throw new Error('El saldo inicial calculado da ' + montoInicial +
      '. Revisá Saldo!A2 y los movimientos antes de continuar.');
  }

  const hoy = new Date();
  const fecha = primera || new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
  mov.insertRowBefore(2);
  mov.getRange(2, 1, 1, 4).setValues([[fecha, 'Ingreso', 'Saldo inicial', montoInicial]]);

  return 'Saldo inicial cargado: $' + montoInicial.toLocaleString('es-AR') +
         ' con fecha ' + Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy') + '.';
}

/** Fórmula de saldo en A2, fecha del último movimiento en C2, y protección. */
function activarFormulaSaldo_(sh) {
  // Apps Script recibe las fórmulas en sintaxis inglesa (comas), sin importar el idioma del sheet
  sh.getRange('A2')
    .setFormula('=SUMIF(Movimientos!B2:B,"Ingreso",Movimientos!D2:D)-SUMIF(Movimientos!B2:B,"Egreso",Movimientos!D2:D)')
    .setNumberFormat('"$"#,##0.00');
  sh.getRange('C2').setFormula('=MAX(Movimientos!A2:A)').setNumberFormat('dd/mm/yyyy');

  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter(p => p.getDescription() === 'Saldo de caja automático')
    .forEach(p => p.remove());
  const prot = sh.getRange('A2:C2').protect().setDescription('Saldo de caja automático');
  const yo = Session.getEffectiveUser();
  prot.addEditor(yo);
  prot.removeEditors(prot.getEditors().filter(e => e.getEmail() !== yo.getEmail()));
  if (prot.canDomainEdit()) prot.setDomainEdit(false);
}

/** Convierte un valor de celda en número: 1565000, "$1,565,000.00", "$ 1.565.000,00", "-". */
function leerMonto_(v) {
  if (typeof v === 'number') return v;
  let s = String(v || '').trim();
  let negativo = false;
  if (/^\(.*\)$/.test(s)) { negativo = true; s = s.slice(1, -1); }
  s = s.replace(/[$\s ]|ARS/gi, '');
  if (s.startsWith('-')) { negativo = !negativo; s = s.slice(1); }
  if (s === '' || s === '-') return 0;
  const c = s.lastIndexOf(','), d = s.lastIndexOf('.');
  if (c > -1 && d > -1) s = c > d ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  else if (c > -1) s = /,\d{1,2}$/.test(s) && s.split(',').length === 2 ? s.replace(',', '.') : s.replace(/,/g, '');
  else if (d > -1 && /^\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  if (!/^\d+(\.\d+)?$/.test(s)) return NaN;
  return negativo ? -parseFloat(s) : parseFloat(s);
}
