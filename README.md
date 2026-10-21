# CALERA — Control de órdenes de servicio y automatización SAP MIRO

CALERA consolida expedientes de servicios por semana y provincia. Su entregable
principal es un maestro de control Excel compatible con la hoja `BD` del modelo
operativo. El reporte histórico `Masivo Reporte.xlsx` continúa disponible para
los bots de SAP MIRO.

Todo el procesamiento documental ocurre en el navegador o en Python local. La
interfaz incluye su propia copia de la dependencia XLSX y no envía los PDF a
servicios externos.

## Flujo de trabajo

1. Cargar la carpeta raíz con los expedientes PDF. El control Excel es opcional
   y se usa únicamente como referencia de estructura del archivo de salida.
2. La interfaz crea un registro por cada HES y completa sus campos con el HES,
   la OC, el comprobante y la SOLPED relacionados por provincia y documento.
3. Descargar el maestro generado y el masivo SAP.
4. Completar manualmente MIRO en la columna U y, si corresponde, observaciones
   en las columnas T y V.
5. Ejecutar la herramienta local de renombrado. La herramienta muestra una
   vista previa y solicita confirmación antes de cambiar nombres.

## Interfaz web

Abre `index.html` en Chrome o Edge. También puede servirse localmente:

```powershell
python -m http.server 8000
```

Visita `http://localhost:8000`.

La interfaz permite:

- Seleccionar una carpeta raíz con subcarpetas, archivos PDF y el control Excel.
- Agregar sedes en varias selecciones sin perder las anteriores.
- Arrastrar directorios de forma recursiva cuando el navegador lo admite.
- Quitar una selección o limpiar todo el estado.
- Importar por separado un maestro editado para recuperar MIRO y observaciones.
- Leer automáticamente la carpeta y procesar todas las semanas, sin selector de
  modo ni de semana.
- Filtrar la vista previa por provincia, tipo, estatus o identificador.
- Descargar el nuevo maestro y el antiguo masivo SAP como entregables distintos.
- El Excel de control no filtra, valida ni aporta filas al nuevo maestro. Los
  datos se extraen de los PDF; un campo que no aparezca en ellos queda vacío y
  se registra como incidencia.

Con la carpeta de ejemplo actual se generan 24 registros: 14 de Arequipa, 3 de
Chimbote, 4 de Chiclayo y 3 de ICA.

## Nuevo maestro A–W

El libro generado contiene:

- `BD`: las 23 columnas de control, de `ID` a `ESTATUS`.
- `RESUMEN`: totales generales y por provincia, tipos documentales, importes y
  estados.
- `INCIDENCIAS`: semana, provincia, documento, fuente, campo, gravedad y motivo.

Las fechas son fechas reales de Excel, los importes son numéricos y RUC/SOLPED/
OC/HES/MIRO se tratan como identificadores para evitar pérdida de dígitos.

La columna U del maestro es `MIRO`. Se deja vacía en una generación nueva. No se
copian los MIRO históricos del archivo de referencia. La columna W usa una
fórmula que devuelve:

- `OBSERVADO` si falta un campo esencial o existe una incidencia bloqueante.
- `PENDIENTE MIRO` si el expediente está validado y U está vacía.
- `CERRADO` si existe un MIRO válido y no hay bloqueos.

La columna U del antiguo `Masivo Reporte.xlsx` continúa siendo `N° OC`. Los dos
formatos no son intercambiables.

`N° Doc` se exporta como referencia SAP: `01-0E001-0000622` para facturas y
`02-0E001-0000047` para recibos por honorarios. Al reimportar el maestro, el
número se normaliza para mantener las coincidencias documentales. Las columnas
Q (`¿Solped Liberada`) y S (`¿OS Liberada?`) conservan las fórmulas del control:
Q convierte `Lib Solped = 2` en 2 y los demás valores en 0; S convierte `G` en
2, `B` en 0 y los demás valores en 1.

## Origen y validación de campos

| Campo | Origen y validación |
|---|---|
| FECHA, SEM, SOLICITANTE, RUC, proveedor, tipo y documento | PDF del comprobante, HES y OC relacionados. Si la semana no está escrita en los PDF, se calcula desde la fecha y se informa como advertencia. |
| VALORIZADO | Importe Total de la factura o Total Neto Recibido/Total por Honorarios del RXH. Para un RXH sin comprobante legible puede usarse el valor de su HES. |
| Provincia | Carpeta `CODIGO_PROVINCIA`. |
| MOTIVO | Descripción del servicio contenida en la HES. |
| SOLPED | Nombre del PDF SOLPED relacionado con el documento. Si no existe, queda vacío; nunca se copia del control ni se inventa. |
| N° OC | Número del PDF de OC y relación indicada en la HES. |
| N° HES | Una fila nueva por cada PDF HES. |
| Liberaciones | Quedan vacías si los PDF no contienen esos indicadores. |
| MIRO y observaciones manuales | Solo un maestro editado que el usuario importa expresamente. |

El masivo SAP utiliza el mismo `VALORIZADO` leído desde el PDF y no aplica un
incremento automático de IGV.

## Generación por Python

Instala dependencias:

```powershell
pip install -r scripts/requirements.txt
```

Generar Semana 41 con expedientes físicos:

```powershell
python scripts/control_ordenes.py expedientes/Semana_41 --modo expedientes --semana 41
```

Generar el maestro completo:

```powershell
python scripts/control_ordenes.py expedientes/Semana_41 `
  --modo completo `
  --semana 41 `
  --maestro expedientes/Semana_41/Control_Ordenes_Servicio_Semana_41.xlsx
```

Recuperar datos manuales de una versión anterior:

```powershell
python scripts/control_ordenes.py expedientes/Semana_41 `
  --semana 41 `
  --maestro-editado "C:\ruta\Control_Ordenes_Servicio_Semana_41_Generado.xlsx"
```

## Renombrado de PDF por MIRO

Ejecuta:

```powershell
python scripts/renombrar_documentos_miro.py
```

La ventana incluye los botones `Seleccionar Excel`, `Seleccionar expedientes`,
`Analizar MIRO`, `Vista previa`, `Renombrar documentos` y `Deshacer último
renombrado`.

El alcance predeterminado renombra solo PDF de HES:

```text
HES 1000094356 - RXH E001-41 - PROVEEDOR MODELO S.A.C.pdf
4300038610 - HES 1000094356 - RXH E001-41 - PROVEEDOR MODELO S.A.C.pdf
```

El alcance opcional incluye comprobante, OC y SOLPED cuando las relaciones son
inequívocas. La herramienta bloquea MIRO inválidos, colisiones, enlaces
simbólicos, archivos con un prefijo diferente y asociaciones ambiguas. Guarda
el historial en `.calera_renombrado/historial_renombrado.json`, dentro de la
carpeta autorizada, y comprueba el hash antes de revertir.

## Reporte SAP existente

La web todavía descarga `Masivo Reporte.xlsx`. También puede generarse con:

```powershell
python scripts/procesar_masivo.py
```

Los bots existentes permanecen separados y aplican política **Fail-Closed**:

```powershell
python scripts/1_sap_miro_scripting.py --simular
python scripts/1_sap_miro_scripting.py
```

* El modo interactivo exige confirmación humana afirmativa explícita (`SI`) para contabilizar.
* El modo `--simular` ejecuta la verificación en SAP y cancela limpiamente (F12) sin persistir datos.
* El modo desatendido `--auto` requiere autorización previa (`CALERA_ALLOW_UNATTENDED=true` o `--forzar-auto-autorizado`).

## Pruebas

```powershell
python -m unittest discover -s tests -v
```

Las pruebas de renombrado crean carpetas temporales y copias de los PDF. Nunca
modifican los expedientes originales.

## Protección de datos y despliegue

`.gitignore` excluye expedientes, libros generados y bitácoras. `.vercelignore`
impide publicar `expedientes/`, `docs/`, `scripts/`, pruebas, PDF y Excel. El
sitio desplegado necesita únicamente `index.html`, `app.js`, `Icono.png` y las
dependencias locales de `vendor/`: XLSX, JSZip y PDF.js con su worker.

Si documentos reales ya se publicaron o entraron en el historial Git, primero
revoca el despliegue afectado y coordina la limpieza del historial con el equipo.
No se reescribe el historial ni se eliminan fuentes originales automáticamente.

Consulta `docs/Guia_Maestro_Control_y_Renombrado.md` para la arquitectura y los
controles técnicos.
