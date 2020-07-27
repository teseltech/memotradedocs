Automatizaciones hojas de saldos
================================

Las automatizaciones al 2020/07/23 consisten de cuatro partes:  
- El recuadro de fórmulas de saldos
- El menú `Fijar Fechas`
- La función `RENDIMIENTOS`
- La función `FACTURACIONTOTAL`

## El recuadro de fórmulas de saldos

![El recuadro de fórmulas][recuadro]

El recuadro de fórmulas esta listo para ser copiado y pegado en el resto de las hojas de los libros. Este recuadro:
1. Calcula cuál fue la última semana completa a partir de la fecha más reciente en la columna `A`
2. Toma la _Inversión Activa_ a partir de la última celda de la columna `F`
3. Calcula la _Utilidad Semanal_ dentro del rango determinado por la semana calculada sumando y filtrando la columna `D`
4. Calcula el _Rendimiento Semanal_ dividiendo la _Utilidad Semanal_ por el último saldo registrado para la semana anterior

Está colocado en la celda H2, al igual que en resto de las hojas. Al momento de copiar, se deberá hacer con las celdas de "ÚLTIMA SEMANA" incluidas, de lo contrario las fórmulas no funcionarán.

**NOTA: No es necesario editar ninguna de estas fórmulas para ninguna de las funcionalidades descritas anteriormente.**

## El menú `Fijar Fechas`

![Menú Fijar Fechas][menu]

Este menú carga unos segundos después de abrir el spreadsheet.
Una vez que se desee convertir **todas** las fechas del spreadsheet a valores fijos se puede dar click a este botón.

Es importante recordar que si se están utilizando fórmulas como `TODAY()`, estas fórmulas se actualizarán al volver abrir el documento, de modo que es importante fijarlas antes de cerrar el spreadsheet.

## El menú `Llenar Hojas RT`

![El menú Llenar hojas RT][hojasrt]

Este menú carga unos segundos después de abrir el spreadsheet.
Antes de _publicar_ las actualizaciones presione este botón para que se copien los valores de las **Hojas Princiales** a las **Hojas RT** esta es una copia de valores y formatos, no de fórmulas. Esto garantiza que se muestren correctamente los valores en Wordpress.

## La función `RENDIMIENTOS`

![La función rendimientos][rendimientos1]

Esta función _capitaliza_ una o más cantidades. Para usarla se requiere una tasa de interés, la fecha actual, la fecha del monto a capitalizar, el monto a capitalizar.

En la fórmula el orden es de esta forma:

```
=RENDIMIENTOS(0.075,A34,C33,A33,F32,A32)
```

- `0.075`: La tasa de interés
- `A34`: La fecha del día actual, el día del corte
- `C33`: El primer monto a capitalizar
- `A33`: La fecha en que se registró el primer monto

Estos dos últimos parámetros se pueden  repetir tantas veces como sea necesario, para dos, tres o cualquier cantidad de montos y fechas. Pero siempre es necesario un monto y una fecha

![La ayuda contextual de la función rendimientos][rendimientos2]

## La función `FACTURACIONTOTAL`

![La función FACTURACIONTOTAL][facturacion]

La función `FACTURACIONTOTAL()` toma un solo argumento, esta es la celda que _buscará_ en el resto del spreadsheet. Esta función **no** suma la hoja en la que es utilizada.

Por ejemplo, si en un libro con las hojas: `Hoja1`, `Hoja2` y `Hoja3` la función se utiliza en la `Hoja1`, el resultado de `=FACTURACIONTOTAL(H2)` será el mismo que `=Hoja2!H2 + Hoja3!H3`

---

## Copiar y Pegar al Scripts Editor

Para que las funciones sean utilizables en los spreadsheets es necesario copiar y pegar este código en el Scripts Editor

![Script Editor][script]

```
/**
 * Capitaliza un valor númerico desde la fecha de aportación hasta la fecha requerida.
 * Puede tener todas las aportaciones necesarias
 *
 * @param {0.05} tasa El valor de la tasa de interés
 * @param {"25/7/2020"} fecha_actual La fecha hasta la que se van a capitalizar las aportaciones
 *
 * @param {1000} k1 La aportación a capitalizar
 * @param {"18/7/2020"} d1 La fecha de aportación
 *
 * @param {"21/7/2020"} d2 La fecha hasta la que se van a capitalizar las aportación 2 – Opcional
 * @param {200} k2 La segunda aportación a capitalizar – Opcional
 *
 * @return Regresa la suma de las aportaciones capitalizadas.
 * @customfunction
*/
function RENDIMIENTOS(tasa, fecha_actual,k1,d1,k2,d2) {

  rendimiento = 0
  for(var arg = 2; arg < arguments.length; arg += 2)
  {
    var delta = Math.abs(fecha_actual - arguments[arg + 1]) / 1000;
    var days = Math.floor(delta / 86400);

    rendimiento += arguments[arg] * tasa * days/7

  }  

  return rendimiento
}

function onOpen() {
  var UI= SpreadsheetApp.getUi();
  UI.createMenu('Automatización')
  .addItem('Fijar Fechas', 'FijarFechas')   // Agregar menú de Fijar Fechas
  .addItem('Llenar hojas RT', 'LlenarRT')   // Agregar menú de copiar valores
  .addToUi();
}

function FijarFechas() {

  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length ; i++ ) {
    var spreadsheet = sheets[i];
    spreadsheet.getRange('A:A').copyTo(spreadsheet.getRange('A:A'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRange('A:A').setNumberFormat('d/M/yyyy');

    var currency = spreadsheet.getRange('C2').getNumberFormat();
    spreadsheet.getRange('C:F').copyTo(spreadsheet.getRange('C:F'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRange('C:F').setNumberFormat(currency);
  }
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Éxito','Se han fijado las fechas', ui.ButtonSet.OK);

}

function LlenarRT() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length ; i++ ) {

    var spreadsheet = sheets[i];
    var spreadsheetName = spreadsheet.getName();

    if (spreadsheetName.substring(0,2) == 'RT') {
      var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(spreadsheetName.substring(3));
      var destinationSheet = spreadsheet;

      sourceSheet.getRange('H1:K2').copyTo(destinationSheet.getRange('A1'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      sourceSheet.getRange('H1:K2').copyTo(destinationSheet.getRange('A1'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

      sourceSheet.getRange('H4:K5').copyTo(destinationSheet.getRange('A4'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      sourceSheet.getRange('H4:K5').copyTo(destinationSheet.getRange('A4'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

      sourceSheet.getRange('H7:K8').copyTo(destinationSheet.getRange('A7'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      sourceSheet.getRange('H7:K8').copyTo(destinationSheet.getRange('A7'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }

  }
}

/**
 * Regresa la suma de todas las celdas que corresponden a 'range' en todas las hojas
 *
 * @param {"H2"} range La celda que se va a sumar de todas las hojas
 *
 * @return La suma de la celda range .
 * @customfunction
*/
function FACTURACIONTOTAL(range) {

  thisSheet = SpreadsheetApp.getActiveSheet();
  var s = 0;
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length ; i++ ) {
    var spreadsheet = sheets[i];
    if (spreadsheet.getSheetName() != thisSheet.getSheetName()) {
      s += parseFloat(spreadsheet.getRange(range).getValue());
    }

  }
  return s
}
```

[recuadro]: recuadro.png
[menu]: menu.png
[hojasrt]: hojasrt.png
[rendimientos1]: rendimientos_1.png
[rendimientos2]: rendimientos_2.png
[facturacion]: facturacion.png
[script]: script.png
