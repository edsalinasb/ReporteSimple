# Formato 5.2 – Libro Diario Simplificado: preparación y exportación a Excel

## Resumen

- `PreparaDetalledelDiarioSimplificado` (en `ReporteSimple.frm`) prepara a la vez las tablas temporales de Crystal (`cfFile10`, `cfFile11`, `cfFile7`) y un modelo en memoria (`RSimple_Diario52.bas`).
- El botón **Excel**, visible solo en Formato 5.2, usa ese modelo para abrir Excel con el diario. No guarda archivos, no genera CSV y no imprime: el libro queda abierto y visible para el usuario.

## Por qué se reemplazó el `RSimple_Excel.bas` anterior

El módulo anterior no era compatible con el sistema real:

- Consultaba `cfFile7`/`cfFile11` con campos inventados (`Fecha`, `Correlativo`, `Glosa`, `Seccion`, `Cuenta`, `Debe`, `Haber`) que esas tablas no tienen. `cfFile11` es una tabla de cabeceras posicionales, no de importes.
- Los totales salían mal: el Debe total iba a la primera columna de cuenta y el Haber total a la segunda.
- La empresa ("TEXTILES TBM") estaba fija en el código.
- Agregaba una columna Cuenta y subtotales por cuenta que el formato no tiene.
- Una función `Public` recibía un tipo `Private` (no compila en VB6).
- No estaba en el `.vbp` y tenía codificación UTF-8 con saltos LF.

## Reglas del reporte

| Orden | Agrupación | Cuentas (2 primeros dígitos) |
|---|---|---|
| 1 | Activo Corriente | 10–29 |
| 2 | Activo No Corriente | 30–39 |
| 3 | Pasivo Corriente | 40–46 |
| 4 | Pasivo No Corriente | 47–49 |
| 5 | Patrimonio | 50–59 |
| 6 | Gestión Egresos | 60–69 |
| 7 | Gestión Ingresos | 70–79 |
| 8 | Saldos Intermedios | 80–89 |
| 9 | Analitica de Explotacion | 90–99 |
| 10 | Cuentas de Orden | clase 0 (`01`…`09`) |

- **Clasificación:** usa estos rangos, no `Pl.Tipo` del plan. Si alguna cuenta está vacía o no empieza con dos dígitos, la preparación se detiene con un mensaje que lista esas cuentas. Ya no se descartan ni se clasifican mal en silencio.
- **Orden de los datos:** agrupación, luego Cuenta, luego saldo inicial y después `Fecha_Documento` (cronológica, como Fecha/Hora). Desempates: `TipoNumero`, `Fecha_Operacion`. El recordset se abre con `ORDER BY` explícito; no depende del orden físico de `SELECT INTO`.
- **Bloques:** cada agrupación se divide en bloques de hasta 14 cuentas (`Cabecera`, columnas S1..S14), en orden ascendente. Cuando una cuenta vuelve a aparecer, se usa la columna que ya tenía asignada (antes se usaba la última columna, lo que desplazaba importes).
- **Importes:** `DebeMN - HaberMN`, de modo que los créditos son negativos. Los ceros quedan en blanco.
- **Subtotales:** solo `TOTAL <agrupación>` por columna de cuenta. No hay subtotales por cuenta o voucher ni total general.
- **CAR:**
  - Ventas (7*): RUC de 11 dígitos + documento.
  - Compras (6*): auxiliar del voucher + documento. El auxiliar es el menor no vacío del voucher, así el resultado no depende del orden de lectura.
  - Resto: `TipoNumero`.
  - El CAR y las cuentas se escriben como texto, por lo que se conservan los ceros a la izquierda.

## Cambios que afectan a Crystal (`FORMATO5_2.rpt`) – verificar

- En `cfFile10`/`cfFile7`, `Tipo` y `Nombre` ahora contienen el número y el nombre de la agrupación de la tabla anterior, en lugar del tipo de balance del plan. `Nombre` se recorta al tamaño del campo.
- `cfFile7` agrega `Fecha_Documento` al final, después de `S14a`, y la incluye en el `GROUP BY`. No se probó si el `.rpt` acepta el campo adicional.
- **Corrección:** la separación debe/haber (criterio "Milton Flores") usaba `S1 > 0 OR …` y `S1 <= 0 OR …`. Si las columnas no usadas valían 0, la misma línea entraba en ambas consultas y se sumaba dos veces. Ahora se separa por la suma de S1..S14 (> 0 y <= 0), que es exacta porque cada línea tiene un solo importe.
- `cfFile14` usa `LEFT JOIN` con el plan y con `xCtbBalance`, para no perder movimientos de cuentas que no figuren en ellos.

## Modo Resumido (compras/ventas)

Las líneas con Tipo 600–799 se resumen por Tipo y Cuenta en tres grupos excluyentes:

1. Debe > 0
2. Debe <= 0 y Haber > 0
3. resto (por ejemplo, importes negativos)

Antes, una línea con Debe y Haber > 0 se sumaba dos veces y las líneas con importes negativos se perdían. Ahora el saldo se conserva. El tratamiento de centros de costo (contrapartida en 7911101, Orden 1/3/4, marca `MesPro='XX'`) mantiene las mismas sentencias y el mismo orden en un único procedimiento compartido (`PreparaTransaccionesConCostos`). `cfTra` nunca se modifica.

## Errores y transacciones

- Cada sentencia se ejecuta con `dbFailOnError` dentro de una transacción propia. Si falla, se revierte solo esa transacción y el error indica la sentencia que falló.
- Los recordsets se cierran en todas las salidas y el puntero del mouse se restaura.
- Si la preparación falla, Crystal no se ejecuta (`PrepararReporte = False`) y Excel no se abre.
- `ActualizaTransacciones`:
  - Ante bloqueos (3260/3197) reintenta como máximo 5 veces y luego falla.
  - Ya no ignora errores (antes había ciclos infinitos y `vbOKCancel` se comparaba con `vbYes`).
- En Excel, si algo falla se cierra sin guardar el libro y se cierra la instancia que se creó. Si todo sale bien, Excel queda abierto y visible.

## Incertidumbre sobre el período (`cMesPro`)

El saldo inicial conserva exactamente el criterio anterior:

- Fecha: último día del mes `cMesPro - 2`.
- Saldo: `ViewSaldo(..., 1, cMesPro - 1)`.
- El encabezado usa `Periodo(cMesPro - 1)`.

Esto sugiere que `cMesPro = 01` corresponde a Apertura, pero el repositorio no lo confirma (depende de `ES26STRIN`/`ES26VERSA`). Por eso no se cambió. La rama interna "enero" del código anterior era inalcanzable y se eliminó sin cambiar el resultado.

## Integración

1. `ES26RSMPL.vbp` ya incluye `RSimple_Diario52.bas` y `RSimple_Excel.bas`. Excel se usa con `CreateObject`, así que no hace falta agregar referencias.
2. El formulario incluye `cmdExcel` (Left 6720, Top 1560), oculto salvo en Formato 5.2.
3. Rutinas sin llamadas (`VerificarArreglo`, `PonerRegistro`, `aCeros`) se dejaron sin cambios.

## Pruebas

```
python3 tests/run_tests.py
```

- Requiere LibreOffice (`libreoffice-calc`).
- Ejecuta 66 pruebas de `tests/D52_Pruebas.bas` sobre `RSimple_Diario52.bas` en LibreOffice Basic (modo VBA). Cubren:
  - límites de agrupación, cuentas de orden y ceros a la izquierda;
  - más de 14 cuentas y cuentas que reaparecen;
  - cuentas inválidas, orden incorrecto e importes Null;
  - CAR;
  - totales con signo y saldo;
  - orden cronológico entre años y saldo inicial primero;
  - reinicio para exportaciones repetidas.
- También compila `RSimple_Excel.bas`, solo como verificación de sintaxis.

**Limitaciones:** no se compiló en VB6. Tampoco se ejecutó contra DAO/Jet real, Excel ni Crystal. Antes de usar en producción, verificar con una base real:

1. Compilar `ES26RSMPL.vbp`.
2. Imprimir el Formato 5.2 en modo Analítico y Resumido, y comparar los totales con el reporte anterior.
3. Revisar que el `.rpt` acepte el campo adicional y las nuevas agrupaciones.
4. Exportar a Excel dos veces seguidas y revisar:
   - ceros a la izquierda en cuentas y CAR;
   - créditos negativos y ceros en blanco;
   - totales por agrupación;
   - continuación de bloques;
   - títulos repetidos al imprimir.
5. Provocar un error (por ejemplo, una cuenta inválida) y confirmar que no queda Excel abierto ni un reporte parcial.
