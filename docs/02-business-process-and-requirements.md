# CaleraReport Enterprise — Proceso de Negocio y Requerimientos

> **Documento:** `docs/02-business-process-and-requirements.md`  
> **Versión:** 1.0.0 (Enterprise Specification)  
> **Fecha:** 21 de octubre de 2026  
> **Estado:** Aprobado / Especificación Oficial  
> **Módulo:** Contexto Empresarial, Proceso As-Is vs To-Be y Requerimientos

---

## 1. Contexto Empresarial

**Sociedad Agrícola La Calera** (Sociedad SAP `1001`) es una de las empresas agroindustriales y avícolas líderes del Perú, con operaciones centrales en Chincha/Ica y centros de distribución y acopio en múltiples provincias del país (**Arequipa**, **Chiclayo**, **Chimbote**, **Ica**).

Para mantener su continuidad operativa semanal, la organización contrata cientos de servicios locales a terceros (saneo ambiental, mantenimiento de maquinaria, estiba portuaria, vigilancia, transporte y suministros eléctricos).

---

## 2. Diagnóstico del Proceso Inicial (Situación As-Is)

Antes de la implementación de CaleraReport, el flujo de pagos a contratistas presentaba serias ineficiencias operativas y riesgos financieros:

```mermaid
flowchart TD
    A["Recepción Descentralizada:\nProveedores envían PDFs por email a cada sede"] --> B["Compilación Manual:\nAsistentes guardan archivos en carpetas desordenadas"]
    B --> C["Cruce Manual en Excel:\nComparación visual PDF vs reportes SAP (ME5A, ME2N, ML84)"]
    C --> D["Digitación Manual en SAP MIRO:\nOperador digita campo por campo en 5 pantallas"]
    D --> E["Renombrado Manual:\nSe añade a mano el N° MIRO al nombre del PDF"]
    E --> F["Riesgos Críticos:\n• Error en importes (tomar valor venta en vez de total)\n• Duplicación de pagos\n• PDFs extraviados sin sustento\n• 4 a 6 horas por lote semanal"]
```

### Principales Puntos de Dolor Identificados:
1. **Dispersión Documental:** Archivos con nomenclaturas heterogéneas, espacios dobles, tildes y nombres incompletos de proveedores.
2. **Confusión en Importes Fiscales:** En facturas, operadores tomaban erróneamente el *Valor Venta* o multiplicaban por 1.18 en comprobantes inafectos. En Recibos por Honorarios, no se discriminaba la retención del 8% de 4ta categoría frente al neto recibido.
3. **Pérdida de Tiempo en SAP MIRO:** La digitación manual de 24 a 50 comprobantes por sede tomaba entre 4 y 6 horas hombre por semana, con alta propensión a errores tipográficos en el número de Orden de Compra o Referencia.
4. **Falta de Trazabilidad:** Si un documento cambiaba de nombre, no existía un registro histórico inmutable para revertir la operación en caso de auditoría interna.

---

## 3. Modelo Operativo Objetivo (Situación To-Be)

CaleraReport transforma este flujo en un ecosistema automatizado, seguro y auditable:

```mermaid
flowchart TD
    subgraph IngestaLocal["1. Ingesta Local Segura (Cero Nube)"]
        A1["Carpetas Provinciales:\nAREQUIPA, CHICLAYO, CHIMBOTE, ICA"] --> A2["Lector Cliente PDF.js / Python pypdf"]
    end

    subgraph ConciliacionAuto["2. Conciliación y Validación Multivía"]
        A2 --> B1["Extracción Semántica Regex (HES, OC, Factura, RXH)"]
        B1 --> B2["Cruce de Identidad Documental & Detección de Faltantes"]
        B2 --> B3["Generación de Incidencias (BLOQUEANTE / ADVERTENCIA)"]
    end

    subgraph SalidasDual["3. Generación Dual de Contratos"]
        B3 --> C1["Maestro de Control A–W\n(Semáforos condicionales & fórmulas nativas)"]
        B3 --> C2["Masivo SAP A–U\n(Estructura compatible /nMIRO)"]
    end

    subgraph AutomatizacionSegura["4. Ejecución Controlada en SAP"]
        C2 --> D1["Robot Local Fail-Closed\n(Simulación F7 + Aprobación Explícita 'SI')"]
        D1 --> D2["Extracción de N° Factura MIRO y Asiento FI"]
    end

    subgraph RenombradoReversible["5. Renombrado con Auditoría"]
        D2 --> E1["Renombrador MIRO con Hash SHA-256"]
        E1 --> E2["Reversión Garantizada en caso de contingencia"]
    end
```

---

## 4. Requerimientos Funcionales y No Funcionales

### 4.1 Requerimientos Funcionales (RF)
* **RF-01 (Ingesta Multicarpetas):** El sistema debe permitir cargar lotes completos de carpetas departamentales en una sola operación o de forma acumulativa mediante Drag & Drop.
* **RF-02 (Parseo Fiscal Exacto):**
  - Facturas: Extraer obligatoriamente el Importe Total consignado (con IGV).
  - Recibos por Honorarios: Extraer el Total Neto Recibido.
* **RF-03 (Generación Maestro A–W):** Generar libro Excel con 3 hojas (`BD`, `RESUMEN`, `INCIDENCIAS`), inmovilización de paneles en `A2`, fórmulas nativas en columnas Q, S, W y formato condicional de 3 semáforos.
* **RF-04 (Generación Masivo SAP A–U):** Generar archivo de carga en batch para `/nMIRO` con 21 columnas y fila de totales sumatoria en columna Q.
* **RF-05 (Gestión Dinámica de MIRO):** Permitir la edición interactiva del código MIRO (10 dígitos) en tabla web con recálculo instantáneo de estados.
* **RF-06 (Renombrado Reversible):** Añadir el prefijo `{MIRO} - ` a los archivos PDF conservando bitácora inmutable con hash SHA-256 para permitir reversión total (*undo*).
* **RF-07 (Automatización SAP Fail-Closed):** Los bots de SAP deben simular primero (F7), verificar saldo 0.00 PEN y exigir confirmación afirmativa explícita (`SI`) antes de emitir la orden de contabilización (`Ctrl + S`).

### 4.2 Requerimientos No Funcionales (RNF)
* **RNF-01 (Privacidad Estricta / Zero-Cloud Leak):** Todo el procesamiento de texto y archivos debe realizarse estrictamente en memoria del cliente o en local. Queda prohibido enviar datos a servidores externos.
* **RNF-02 (Arquitectura Offline / Zero-CDN):** La aplicación web debe incluir todas sus bibliotecas (`pdf.js`, `xlsx`, `jszip`) de forma local en la carpeta `vendor/` para operar sin conexión a Internet.
* **RNF-03 (Reproducibilidad de Pruebas):** El repositorio debe incluir fixtures sintéticos 100% anonimizados que permitan ejecutar la suite de pruebas unitarias inmediatamente tras un clon limpio en Linux o Windows.
* **RNF-04 (Rendimiento):** El procesamiento de un lote completo de 4 provincias (70+ archivos PDF) debe completarse en menos de 5 segundos.
