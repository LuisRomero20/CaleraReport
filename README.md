# Consolidador Masivo de Reportes de Pago a Proveedores (SAP) 📊

Herramienta diseñada para el procesamiento, validación y consolidación masiva de expedientes de pago a proveedores por departamentos (RxH, Facturas, OCs y Consolidados), generando el archivo oficial **`Masivo Reporte.xlsx`** listo para carga y contabilización en **SAP**.

---

## 🚀 Características Principales

- **🌐 Herramienta Web 100% Local (Sin Instalación ni Servidores)**:
  - Funciona directamente abriendo `index.html` o `Masivo Reporte Tool.html` en Chrome, Edge o cualquier navegador moderno.
  - Diseñada especialmente para entornos corporativos con restricciones de instalación de software o permisos de administrador.
  - Procesamiento 100% en el cliente: los datos nunca salen de la computadora.

- **📁 Soporte Multi-Departamento en Lote**:
  - Permite arrastrar una carpeta raíz que contenga múltiples subcarpetas de departamentos (`1110 AREQUIPA`, `1120 TRUJILLO`, etc.).
  - Genera automáticamente un único libro Excel con **una pestaña independiente por cada departamento** (ej. `Arequipa Semana 34`).

- **📄 Extracción Automática de Fechas desde PDFs de OC**:
  - Escanea internamente cada PDF de Orden de Compra (`...OC...PDF`) presente en las carpetas.
  - Extrae de forma automática la **Fecha de Emisión** oficial (ej. `26.08.2026`) vinculándola a su respectiva OC / RxH.
  - Manejo seguro de zonas horarias (UTC) para evitar cualquier desfase de días.

- **🗓️ Filtrado Inteligente por Semana Reciente**:
  - Detecta el número máximo de semana en los archivos de consolidado histórico (ej. Semana 34) y filtra automáticamente solo los registros que corresponden al período actual.

- **🎨 Formato y Estructura Oficial SAP (Columnas A – U)**:
  - **Columna L**: Texto de Posición estandarizado (ej. `SERVICIO DE LIMPIEZA SEMANA 34`, `SERVICIO DE ESTIBA SEMANA 34`).
  - **Columna Q**: Importe en formato de moneda peruana (`S/ #,##0.00`) con resaltado contable.
  - **Columna T**: Copia exacta y sincronizada de la columna L.
  - **Columna U**: Número de Orden de Compra (`N° OC`) correspondiente a cada comprobante.
  - **Fila de TOTAL**: Calculada automáticamente con fórmula dinámica `=SUM(...)`.

- **✏️ Barra de Edición Rápida de Fecha**:
  - En la vista previa web se muestra el origen de la fecha (`✅ Extraída de Orden de Compra`) y permite modificar o unificar la fecha de toda la hoja con un solo clic.

---

## 📂 Estructura del Repositorio

```text
├── index.html                  # Aplicación Web principal (landing GitHub Pages)
├── Masivo Reporte Tool.html    # Aplicación Web standalone (copia de trabajo directo)
├── Masivo Reporte.xlsx         # Reporte final consolidado generado
├── requirements.txt            # Dependencias opcionales para Python
├── .gitignore                  # Reglas de exclusión de archivos temporales
├── README.md                   # Documentación oficial del proyecto
│
├── scripts/
│   └── procesar_masivo.py      # Motor CLI en Python para procesamiento por terminal
│
└── 1110 AREQUIPA/              # Carpeta de ejemplo / datos de departamento
    ├── SEM 34 CONSOLIDADO PAGO PROV aqp.xlsx
    ├── AREQUIPA OC 430060343 RXH E001-160 AGRIPINA VALERIANO.PDF
    ├── ... (demás OCs, HES y Solpeds en PDF)
```

---

## 💻 Instrucciones de Uso

### Opción 1: Interfaz Web (Recomendada para Usuarios Finales)

1. Haz doble clic en [`index.html`](index.html) o [`Masivo Reporte Tool.html`](Masivo%20Reporte%20Tool.html) para abrirlo en tu navegador (Chrome / Edge).
2. Haz clic en **"Seleccionar carpeta de departamentos"** o arrastra la carpeta que contiene los departamentos (ej. `1110 AREQUIPA`).
3. El sistema procesará en segundos cada departamento:
   - Detecta el archivo de consolidado.
   - Lee las Órdenes de Compra en PDF y extrae sus fechas de emisión.
   - Construye la tabla con las columnas A–U.
4. Revisa la **Vista Previa** interactiva (pestañas por departamento, fecha detectada, importes en S/ y OCs).
5. Haz clic en **📥 Descargar Masivo Reporte.xlsx**.

### Opción 2: Script en Python (Automatización por Terminal)

Si prefieres ejecutar el proceso por línea de comandos:

```bash
# 1. Instalar dependencias (solo primera vez)
pip install -r scripts/requirements.txt

# 2. Ejecutar el procesador
python scripts/procesar_masivo.py
```

El script escaneará automáticamente las carpetas del proyecto y generará el archivo `Masivo Reporte.xlsx` con todas las pestañas y estilos listos.

---

## 📋 Mapeo de Columnas Oficiales (SAP)

| Columna | Encabezado | Origen / Regla |
|:---:|:---|:---|
| **A** | `ID CABC(1)` | Contador correlativo secuencial (1, 2, 3...) |
| **B** | `Fecha Doc` | Fecha de emisión extraída del PDF de la OC (ej. `26.08.2026`) |
| **C** | `Texto Cab Doc` | Vacío / Opcional según SAP |
| **D** | `REFERENCIA` | Recibo por Honorarios formateado (`02-0E001-0000160`) |
| **E** | `ID POSC(2)Acreedor` | Constante contable `2` |
| **F** | `RUC` | RUC del proveedor / prestador |
| **G** | `Cond Pago` | Condición de pago estándar `CP02` |
| **H** | `Ind Impto` | Indicador de impuesto `C0` |
| **I** | `Tipo Detr` | Vacío |
| **J** | `Ind Detra` | Vacío |
| **K** | `Asignación` | Vacío |
| **L** | `Texto Posición` | Motivo normalizado + `SEMANA {N}` |
| **M** | `ID POSC(2) Cta GASTO` | Constante contable `2` |
| **N** | `Cuenta Gasto` | Vacío |
| **O** | `DETALLE` | Vacío |
| **P** | `Centro Costo` | Vacío |
| **Q** | `Importe MD` | Valor monetario formateado en `S/ #,##0.00` |
| **R** | `Ind Impto` | Indicador `C0` |
| **S** | `Asignación` | Vacío |
| **T** | `Texto Posición` | **Copia exacta e idéntica de la Columna L** |
| **U** | `N° OC` | Número de Orden de Compra extraído del consolidado/PDF |

---

## 🛠️ Tecnologías Empleadas

- **Frontend Web**: HTML5, Vanilla CSS moderno (paleta institucional dark/forest green, glassmorphism, responsive), Vanilla JavaScript (ES6+).
- **Procesamiento de Libros Excel**: `xlsx-js-style` (generación y formateo completo de estilos XML en el navegador).
- **Lectura Binaria de PDFs**: Parser interno de streams en JavaScript puro y `pypdf` para Python.
- **Motor CLI**: Python 3.10+ con `openpyxl`.

---

## 🔒 Privacidad y Seguridad

Esta solución está diseñada respetando los más altos estándares de privacidad:
- **Zero-Cloud**: Ningún archivo, factura o dato personal se sube a internet.
- Funciona 100% desconectado tras la primera carga de caché de la librería de estilos.
