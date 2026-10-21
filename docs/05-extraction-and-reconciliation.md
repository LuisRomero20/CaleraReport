# CaleraReport Enterprise — Extracción y Conciliación Documental (M1)

> **Documento:** `docs/05-extraction-and-reconciliation.md`  
> **Versión:** 1.0.0 (Enterprise Specification)  
> **Fecha:** 21 de octubre de 2026  
> **Estado:** Aprobado / Especificación Oficial  
> **Módulo:** Motores de Parseo Regex, Conciliación Multivía y Calidad de Datos

---

## 1. Arquitectura de Ingesta y Procesamiento Local

El motor de extracción documental opera bajo el principio de **Privacidad por Diseño (Privacy by Design)**:
* **Cero APIs Externas:** Ningún archivo PDF, comprobante fiscal ni dato de proveedor es transmitido a la nube, servidores de terceros ni modelos LLM externos.
* **Procesamiento Dual Idéntico:**
  1. **Motor Web en Navegador:** Implementado en `app.js` mediante la biblioteca local `vendor/pdf.min.mjs` y su worker compilado.
  2. **Motor CLI en Servidor/Local:** Implementado en `scripts/control_ordenes.py` mediante `pypdf`.
* **Detección sin OCR Automático:** Se prioriza la extracción de capas de texto embebidas en los flujos PDF de SUNAT y SAP. Si un documento está rasterizado o carece de texto, se detecta de forma determinista (`readable == false`), generando una incidencia informativa y utilizando los metadatos de nomenclatura normalizada sin forzar OCRs pesados o inseguros.

```mermaid
flowchart TD
    RawFiles["Entrada: Lote de Carpetas / Provincias"] --> TypeFilter{"Clasificación por Nombre & Extensión"}
    TypeFilter -->|PDF| PdfParser["Parser PDF (Local PDF.js / pypdf)"]
    TypeFilter -->|Excel Histórico| ExcelParser["Parser Maestro (BD / ME5A / ME2N / ML84)"]
    TypeFilter -->|No soportado| Discard["Ignorado con Registro"]

    PdfParser --> TextExtract{"¿Contiene capa de texto?"}
    TextExtract -->|Sí| RegexEngine["Motor Regex de Campos Semánticos"]
    TextExtract -->|No (Rasterizado)| MetaFallback["Extracción por Nomenclatura + Incidencia ADVERTENCIA"]

    RegexEngine --> NormalizedEntities["Entidades Normalizadas (HES, OC, Comprobante, SOLPED)"]
    MetaFallback --> NormalizedEntities

    NormalizedEntities --> Reconciler["Motor de Conciliación Multivía"]
    ExcelParser --> Reconciler

    Reconciler --> ValidationGates{"Puertas de Calidad & Consistencia"}
    ValidationGates -->|Válido| MasterRecord["Fila Apta (PENDIENTE MIRO / CERRADO)"]
    ValidationGates -->|Discrepancia| BlockedRecord["Fila Bloqueada (OBSERVADO + Incidencias)"]
```

---

## 2. Motor de Extracción Semántica por Expresiones Regulares

### 2.1 Facturas Electrónicas
* **Identificación:** Cabecera con `"FACTURA ELECTRONICA"` o nombre con patrón `FACTURA [A-Z]\d{3}-\d+`.
* **Regla Crítica de Importes:** **Importe Total Consignado** (incluye IGV 18%).
  * *Expresión Regular:*
    ```regex
    \b(?:IMPORTE\s+TOTAL|TOTAL)\s*:\s*(?:S\/\.?\s*)?([0-9,]+\.\d{2})
    ```
  * *Control Financiero:* Se prohíbe terminantemente tomar el `"Valor Venta"` o multiplicar importes netos por 1.18. Se extrae el monto total liquidado por SUNAT.
* **RUC del Proveedor:** `\bFACTURA\s+ELECTRONICA\s+RUC\s*:\s*(\d{11})\b` o `\bRUC\s*:\s*(\d{11})\b`.

### 2.2 Recibos por Honorarios Electrónicos (RXH)
* **Identificación:** Cabecera con `"RECIBO POR HONORARIOS ELECTRONICO"` o patrón `RXH [A-Z]\d{3}-\d+`.
* **Regla Crítica de Importes:** **Total Neto Recibido**.
  * En SUNAT, los RXH mayores a S/ 1,500 sufren retención del 8% de 4ta categoría. El desembolso real de tesorería y la provisión en SAP MIRO corresponden al **Neto Recibido**:
  * *Expresión Regular:*
    ```regex
    \bTOTAL\s+NETO\s+RECIBIDO\s*:?[^0-9]*([0-9,]+\.\d{2})\s+SOLES
    ```
  * Si el recibo no está sujeto a retención, el Total Neto coincide con el Total por Honorarios.

### 2.3 Hoja de Entrada de Servicios (HES)
* **Número de HES (10 dígitos):** `\bN[UÚ]MERO\s*/?\s*FECHA\s+(10\d{8,10})\b`.
* **Vínculo a Orden de Compra (OC):** `\bPOSICION\s+PEDIDO\s*/?\s*FECHA\s+(43\d{7,10})\b`.
* **Importe Valorizado HES:** `\bVALOR\s+TOTAL\s+SIN\s+IVA\s+PEN\s+([0-9,]+\.\d{2})\b`.
* **Descripción del Servicio:** `\b\d{7}\s+(.{3,120}?)(?:L.{0,8}PLANIFICACI|LINEA\s+PLANIFICACION)`.

### 2.4 Orden de Compra / Servicio (OC)
* **Número de OC (10 dígitos):** `\bORDEN\s+DE\s+SERVICIO\s+N\s*(43\d{7,10})\b`.
* **Vínculo a SOLPED:** `\bSOLPED\s*:?\s*(\d{7,10})\b`.
* **Importe Contractual:** `\bNETO\s+DESCUENTO\s+SUB-TOTAL\s+I\.G\.V\.\s+TOTAL\s+([0-9,]+\.\d{2})\b`.

---

## 3. Algoritmo de Conciliación Multivía (Cross-Document Matching)

La conciliación sigue un orden determinista de encadenamiento con aislamiento provincial estricto:

### Paso 1: Pivote en la Hoja de Entrada de Servicios (HES)
Cada archivo PDF identificado como `HES` genera un registro candidato. Una HES sin comprobante o sin OC no puede cerrarse, pero sí debe registrarse con su incidencia correspondiente.

### Paso 2: Asociación de Orden de Compra (OC)
1. Si la HES consigna explícitamente el número de OC en su texto, se busca en la misma provincia un PDF de OC cuyo número coincida.
2. Si la HES no consigna la OC en texto, se busca por coincidencia cruzada con el número de comprobante consignado en el nombre del archivo.
3. **Control de Ambigüedad:** Si existen 2 o más PDFs de OC distintos que reclaman la misma HES o comprobante, el sistema **no elige arbitrariamente**: registra una incidencia `BLOQUEANTE` (`"Se encontraron varios PDF de OC candidatos para el mismo expediente"`).

### Paso 3: Vinculación con Comprobante Fiscal (Factura / RXH)
1. Se asocia por el identificador de serie y número normalizado (ej. `E001-41`).
2. Se verifica que el RUC del comprobante coincida con el RUC consignado en la HES y en la OC.
3. Se compara el importe valorizado: si el comprobante no es legible pero la HES tiene sustento físico, se emite advertencia de trazabilidad.

### Paso 4: Asociación de SOLPED
1. Se extrae de los archivos `SOLPED` relacionados o del campo SOLPED de la OC.
2. La SOLPED es opcional para ciertos servicios directos. Si no existe, el campo queda vacío y no bloquea el expediente.

---

## 4. Métricas Reproducibles de Rendimiento y Calidad por Lote

El motor emite un resumen auditable por ejecución, calculado mediante la función `to_summary()`:

$$\text{Tasa de Conciliación (\%)} = \left( \frac{\text{Registros con Soporte Físico Completo}}{\text{Total de Registros Extraídos}} \right) \times 100$$

### Indicadores Oficiales del Contrato de Auditoría:
| Métrica | Propósito de Auditoría | Criterio de Éxito |
|---|---|---|
| `selected_files` | Total de archivos procesados en las carpetas. | Coincide con recuento de sistema de archivos. |
| `records` | Total de líneas de servicio consolidadas en `BD`. | $\ge 1$ por cada HES detectada. |
| `supported_records` | Expedientes con HES, OC y Comprobante 100% coincidentes. | Objetivo: $100\%$ en lotes completos. |
| `blocking_incidents` | Desviaciones críticas que impiden el pago o registro. | $0$ para lotes listos para facturación. |
| `warning_incidents` | Advertencias operativas conservadas con trazabilidad. | Auditables en la hoja `INCIDENCIAS`. |
| `closed_records` | Registros con MIRO válido cerrados. | Condicionado a MIRO numérico de 10 dígitos. |
| `pending_miro` | Registros validados que esperan gestión en SAP MIRO. | Insumo directo para el robot SAP. |
| `total_valorizado` | Suma aritmética exacta de los importes a pagar (PEN). | Cuadre exacto con la sumatoria de comprobantes. |
| `reconciliation_rate` | Porcentaje de salud documental del lote. | Medido matemáticamente, sin estimaciones. |

---

## 5. Prevención de Duplicados y Manejo de Conflictos

1. **Idempotencia de Selección:** Si el usuario carga la misma carpeta dos veces o arrastra subdirectorios superpuestos, el motor compara el hash y la clave `(GroupKey, motivo, valorizado)`. El duplicado es ignorado silenciosamente y se registra una advertencia informativa.
2. **Conflictos en Regeneración:** Si un usuario importa un maestro editado que contiene dos filas con la misma clave documental pero diferentes códigos MIRO (ej. `5105773695` vs `5105773699`), el sistema **rechaza la asignación automática**, emite una incidencia `BLOQUEANTE` en la hoja `INCIDENCIAS` y mantiene el registro en estado `OBSERVADO` hasta aclaración contable.
