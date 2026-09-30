# Invoices Archive — Maqueta con "Cruzar información"

Maqueta funcional (HTML/CSS/JS, sin compilación) de la vista **Invoices Archive** con la nueva funcionalidad de cruce contra un reporte externo.

## Cómo abrirla

Abra `index.html` con doble clic (Chrome o Edge). No requiere servidor ni instalación.

Para probar el flujo completo:
1. (Opcional) Cambie el **Date Range**, por ejemplo `2026-09-01 ~ 2026-09-20`.
2. Clic en **Reconcile**.
3. Cargue el reporte de **PL Colab** (solo .xlsx). **Debe tener una columna llamada `Nro. Documento`**; si no la tiene, no se puede cruzar.
4. En **Select period to reconcile** elija el rango de **Issue Date** de Project Agenda (se sugiere automáticamente con las fechas de emisión del archivo; máximo 3 meses) y haga clic en **Continue**: el cruce se ejecuta y se abre el reporte.
5. Se abre **PL Colab & Project Agenda Invoice Reconciliation**: total de facturas de cada fuente y la tabla (Consecutive, Supplier, Issue Date) con las facturas de PL Colab no encontradas en Project Agenda. **Export to Excel** descarga exactamente esa vista con todas las filas.

> PL Colab = reporte externo cargado (se analiza completo). Project Agenda = facturas del software con Issue Date en el período elegido en el cruce. Los registros duplicados se listan cada vez que aparecen.

## Estructura

```
maqueta/
├── index.html              Vista + modales
├── css/styles.css          Estilos (Material, igual a la vista actual)
├── js/data-software.js     1.386 registros del reporte del software (tomados del Excel, sin modificar)
├── js/crossmatch.js        Lógica del cruce (normalización, llave, coincidencias, duplicados)
├── js/app.js               Tabla principal, filtros, ordenamiento, rango de fechas, REPORT
├── js/cross-ui.js          Modales: carga, análisis, proceso y resultado (Reconciliation + Export to Excel)
├── lib/xlsx.full.min.js    SheetJS 0.20.3 (lectura/escritura de Excel en el navegador, local)
└── data/…xlsx              Excel de referencia (usar como reporte externo de ejemplo)
```

## Análisis de los archivos

### Columnas

| Excel            | Pantallazo      | Contenido                                                            |
|------------------|-----------------|----------------------------------------------------------------------|
| Type             | TYPE            | `material` (450), `No order` (418), `comex services` (249), `service` (232), `asset` (37) |
| Supplier         | SUPPLIER        | Razón social del proveedor                                           |
| Invoice ID       | INVOICE ID      | Nº de factura; en el Excel es un hipervínculo al PDF                 |
| Issue Date       | ISSUE DATE      | `YYYY-MM-DD`, 2026-09-01 a 2026-09-29                                |
| Due Date         | DUE DATE        | `YYYY-MM-DD`; 195 vacíos                                             |
| Invoice Amount   | INVOICE AMOUNT  | Texto con formatos mixtos (`86,388,016.10`, `19.782.023`, `99.999,99`); 13 vacíos |
| OC               | OC              | Orden de compra; vacía en los 418 registros `No order`               |
| SAP              | SAP (checkbox + filtro None) | `checked` / `unchecked`                                 |
| ACCOUNTANT       | ACCOUNTANT (checkbox + filtro None) | `checked` / `unchecked`                         |
| —                | ACTIONS         | Solo en la interfaz (menú ⋮)                                         |

**Diferencias pantallazo vs. Excel:** solo *ACTIONS*, que es un control de la interfaz. Los valores se muestran exactamente como vienen en el Excel (por eso *Invoice Amount* conserva sus formatos mixtos, igual que en el pantallazo).

### Identificador de la factura

- **Regla fija del cruce:** `Invoice ID` (software / Project Agenda) contra la columna **`Nro. Documento`** del reporte de PL Colab. No se usa el proveedor.
- Ambos valores se comparan normalizados: mayúsculas, sin espacios, guiones ni puntos (`fe-677` = `FE677`).
- El reporte (**PL Colab & Project Agenda Invoice Reconciliation**) muestra solo las facturas de PL Colab cuyo `Nro. Documento` **no está** en el software para el período seleccionado.

### Hallazgos en los datos

- 1.386 registros → 1.109 facturas únicas.
- 132 facturas aparecen más de una vez (276 filas repetidas). Ejemplo: `TRANSCONT SAS / TRA284442` aparece 16 veces. **No se eliminan**: se reportan como duplicados, indicando si las filas son idénticas o en qué campos varían.
- 1 registro sin Invoice ID (`ALLOY FASTENERS INC (RI)`, 2026-09-02) → se reporta como **Error** (no se puede cruzar).

## Reglas del cruce

- **Software** = registros de la vista filtrados por el *Date Range* (fecha de emisión).
- **Reporte externo** = se analiza completo, sin filtrar por fecha.
- Se compara registro contra registro por llave:
  - *Encontradas*: la llave existe en ambos.
  - *No encontradas en el software*: solo en el externo. Si la factura existe en el software fuera del período, se indica en *Observación*.
  - *No encontradas en el externo*: solo en el software.
  - *Duplicados*: llave repetida dentro de un mismo reporte (por fuente).
  - *Con diferencias*: coinciden, pero ningún registro comparte valor (±1), fecha u OC.
  - *Errores*: registros sin Nº de factura.
- El mapeo de columnas del externo se detecta por nombre (inglés/español) y se puede ajustar.

## Supuestos de la maqueta

- El enlace al PDF de *Invoice ID* no descarga nada (las URLs originales contienen tokens de acceso y no se incluyeron).
- Los checkboxes SAP/ACCOUNTANT no se pueden marcar: sus valores son aleatorios, solo para presentación (si ACCOUNTANT está marcado, SAP también).
- **REPORT** exporta a Excel las filas visibles del período.
