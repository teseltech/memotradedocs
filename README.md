Automatizaciónes hojas de saldos
================================

Las automatizaciones al 2020/07/23 consisten de cuatro partes:  
- El recuadro de fórmulas de saldos
- El menú `Fijar Fechas`
- La función `RENDIMIENTOS`
- La función `FACTURACIONTOTAL`

## El recuadro de fórmulas de saldos

![El recuadro de fórmulas][recuadro]

El recuadro de formulas esta listo para ser copiado y pegado en el resto de las hojas de los libros. Este recuadro:
1. Calcula cual fue la última semana completa a partir de la fecha más reciente en la columna `A`
2. Toma la _Inversión Activa_ a partir de la última celda de la columna `F`
3. Calcula la _Utilidad Semanal_ dentro del rango determinado por la semana calculada sumando y filtrando la columna `D`
4. Calcula el _Rendimiento Semanal_ dividiendo la _Utilidad Semanal_ por el último saldo registrado para la semana anterior

## El menú `Fijar Fechas`

![Menú de fórmulas][menu]

Este menú carga unos segundos después de abrir el spreadsheet.
Una vez que se desee convertir **todas** las fechas del spreadsheet a valores fijos es puede dar click a este botón.

Es importante recordar que si se estan utilizando fórmulas como `TODAY()` estas fórmulas se actualizarán al volver abrir el documento, de modo que es importante fijarlas antes de cerrar el spreadsheet.

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

Por ejemplo. En un libro con las hojas: `Hoja1`, `Hoja2`, `Hoja3` y la función se utiliza en la `Hoja1` el resultado de `=FACTURACIONTOTAL(H2)` será el mismo que `=Hoja2!H2 + Hoja3!H3`


[recuadro]: recuadro.png
[menu]: menu.png
[rendimientos1]: rendimientos_1.png
[rendimientos2]: rendimientos_2.png
[facturacion]: facturacion.png
