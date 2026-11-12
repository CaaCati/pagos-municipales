/**
 * Ordena y blinda el sheet Caja_chica.
 *
 * Cómo usarlo (una sola vez):
 *   1. En el sheet: Extensiones → Apps Script.
 *   2. Botón "+" → Secuencia de comandos → nombre: ordenar_caja_chica
 *   3. Pegar este archivo completo y guardar (Ctrl+S).
 *   4. Elegir la función ordenarMovimientos en el desplegable y tocar "Ejecutar".
 *      La primera vez Google pide autorizar el script: es tu propio script.
 *   5. Cuando esté definido el saldo inicial, cargarlo en Movimientos y ejecutar
 *      activarSaldoAutomatico.
 *
 * No borra ni cambia montos, fechas ni descripciones. Solo:
 *   - corrige el texto de Tipo ("egreso", "EGRESO ", etc. → "Egreso"/"Ingreso")
 *   - unifica formatos (fecha dd/mm/yyyy, monto en moneda, alineación)
 *   - agrega validaciones para que no se pueda cargar mal
 */

const HOJA_MOV = 'Movimientos';
const HOJA_SALDO = 'Saldo';
const FILAS = 1000; // filas preparadas para carga (se pueden agregar más)

function ordenarMovimientos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(HOJA_MOV);
  if (!sh) throw new Error('No existe la hoja ' + HOJA_MOV);

  // Asegurar que existan las filas a validar
  if (sh.getMaxRows() < FILAS + 1) sh.insertRowsAfter(sh.getMaxRows(), FILAS + 1 - sh.getMaxRows());

  // 1. Encabezado fijo
  sh.getRange('A1:D1').setValues([['Fecha', 'Tipo', 'Descripcion', 'Monto']]);
  sh.setFrozenRows(1);

  // 2. Normalizar Tipo existente
  const ultima = sh.getLastRow();
  let corregidos = 0;
  if (ultima >= 2) {
    const rangoTipo = sh.getRange(2, 2, ultima - 1, 1);
    const tipos = rangoTipo.getValues().map(([t]) => {
      const n = String(t).trim().toLowerCase();
      const nuevo = n === 'egreso' ? 'Egreso' : n === 'ingreso' ? 'Ingreso' : t;
      if (nuevo !== t) corregidos++;
      return [nuevo];
    });
    rangoTipo.setValues(tipos);
  }

  // 3. Formatos uniformes
  const datos = sh.getRange(2, 1, FILAS, 4);
  datos.setHorizontalAlignment('center').setVerticalAlignment('middle')
       .setFontFamily('Arial').setFontSize(10);
  sh.getRange(2, 1, FILAS, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(2, 4, FILAS, 1).setNumberFormat('"$"#,##0.00');

  // 4. Validaciones (rechazan lo que no cumpla)
  sh.getRange(2, 1, FILAS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .setHelpText('Ingresá una fecha válida (dd/mm/aaaa). Doble clic abre el calendario.')
      .build());

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

  SpreadsheetApp.getUi().alert(
    'Movimientos ordenado.\n' +
    'Tipos corregidos: ' + corregidos + '\n' +
    'Validaciones aplicadas a las filas 2 a ' + (FILAS + 1) + '.');
}

/**
 * Ejecutar recién cuando Movimientos tenga la fila "Saldo inicial" (Ingreso).
 * Pone en Saldo!A2 el cálculo Ingresos − Egresos, en C2 la fecha del último
 * movimiento, y protege A2:C2 para que solo el dueño pueda modificarlas.
 */
function activarSaldoAutomatico() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mov = ss.getSheetByName(HOJA_MOV);
  const sh = ss.getSheetByName(HOJA_SALDO);

  const tieneInicial = mov.getRange('B2:C').getValues()
    .some(([tipo, desc]) => String(tipo).toLowerCase() === 'ingreso'
                         && String(desc).toLowerCase().includes('saldo inicial'));
  if (!tieneInicial) {
    SpreadsheetApp.getUi().alert(
      'Falta el saldo inicial: cargá en Movimientos una fila Tipo "Ingreso" con ' +
      'Descripcion "Saldo inicial" antes de activar el cálculo.');
    return;
  }

  sh.getRange('A2').setFormula(
    // Apps Script recibe las fórmulas en sintaxis inglesa (comas), sin importar el idioma del sheet
    '=SUMIF(Movimientos!B2:B,"Ingreso",Movimientos!D2:D)-SUMIF(Movimientos!B2:B,"Egreso",Movimientos!D2:D)')
    .setNumberFormat('"$"#,##0.00');
  sh.getRange('C2').setFormula('=MAX(Movimientos!A2:A)').setNumberFormat('dd/mm/yyyy');

  // Proteger A2:C2: solo el dueño edita
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter(p => p.getDescription() === 'Saldo de caja automático')
    .forEach(p => p.remove());
  const prot = sh.getRange('A2:C2').protect().setDescription('Saldo de caja automático');
  const yo = Session.getEffectiveUser();
  prot.addEditor(yo);
  prot.removeEditors(prot.getEditors().filter(e => e.getEmail() !== yo.getEmail()));
  if (prot.canDomainEdit()) prot.setDomainEdit(false);

  SpreadsheetApp.getUi().alert('Saldo automático activado y protegido.');
}
