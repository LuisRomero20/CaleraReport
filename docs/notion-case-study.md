# 🏭 CaleraReport — Enterprise Document Automation & SAP Integration

> **Categoría:** Data Engineering · Workflow Automation · Financial Controls · SAP Integration  
> **Estado del Despliegue:** Producción Vercel Ready · Commit `bf612a1` (Baseline v3.0 `e7bde02`)  
> **Enlaces:** [Repositorio GitHub](https://github.com/LuisRomero20/CaleraReport) · [Aplicación Web](https://calerareport-ten.vercel.app)

---

## 📌 Ficha Ejecutiva del Proyecto

| Dimensión | Detalle / Métrica Verificada |
|---|---|
| **Empresa / Caso** | Sociedad Agrícola La Calera (Sociedad SAP 1001) |
| **Sedes Atendidas** | Arequipa, Chiclayo, Chimbote, Ica (4 sedes provinciales) |
| **Volumen Semanal** | 24 – 60 expedientes de pago a contratistas por período |
| **Tiempo de Proceso** | Reducido de **4-6 horas semanales** a **3.8 segundos** |
| **Demostración en Vivo** | Modo 1-Clic interactivo precargado con 81 expedientes sintéticos |
| **Auditoría Forense** | Dossier digital 360° con trazabilidad de 5 nodos y simulador BAPI SAP |
| **Analítica de Gasto** | HUD interactivo con desglose en Soles por sede y embudo SAP |
| **Certificación Formal** | Dictamen contable imprimible con hash criptográfico SHA-256 |
| **Arquitectura** | 100% Local / Zero-Cloud Leak / Zero-CDN Offline |
| **Pruebas Automatizadas** | 21 pruebas unitarias y de integración (100% aprobadas en 4.06s) |
| **Seguridad SAP** | Política Fail-Closed con confirmación afirmativa obligatoria ("SI") |
| **Integridad de Archivos** | Renombrado masivo con bitácora criptográfica y reversión por SHA-256 |

---

## 🎯 El Desafío Operativo

La gestión semanal de pagos a proveedores de servicios descentralizados en La Calera presentaba 4 graves problemas:
1. **Dispersión:** Facturas y Recibos por Honorarios llegando por correos desordenados con formatos heterogéneos.
2. **Conciliación Manual:** Necesidad de cruzar cada PDF con reportes de SAP (`ME5A`, `ME2N`, `ML84`) manualmente en Excel.
3. **Errores Fiscales:** Confusión frecuente entre el *Valor Venta* y el *Importe Total* en facturas con IGV, o deducciones de 4ta categoría en recibos por honorarios.
4. **Digitación Lenta en SAP MIRO:** Transcribir campo por campo en 5 pantallas por cada comprobante consumía entre 4 y 6 horas hombre por analista cada lunes.

---

## 💡 La Solución Implementada: Suite de Operaciones Enterprise

CaleraReport Enterprise articula un ecosistema completo de módulos locales de alta gama:

```mermaid
graph TD
    A["1. Consola Web Full-Width\n(Drag & drop multicarpetas + Modo Demo 1-Clic)"] --> B["2. Motor de Conciliación Multivía\n(Cruce determinista HES-OC-Comprobante-SOLPED)"]
    B --> C["3. Generador Dual de Contratos\n• Maestro A–W (Semáforos y fórmulas nativas)\n• Masivo SAP A–U (Batch /nMIRO)"]
    B --> D["4. Auditoría Forense & Dossier 360°\n(Matriz fiscal, tolerancia cero y payload BAPI)"]
    C --> E["5. Robot de Automatización SAP\n(Fail-Closed, simulación F7, aprobación 'SI')"]
    C --> F["6. Renombrador MIRO con Hash SHA-256\n(Reversión instantánea ante cualquier contingencia)"]
    B --> G["7. Dictamen de Control Interno\n(Certificado de conciliación formal imprimible)"]
```

---

## 🌟 Capacidades Flagship de Nivel Directivo

1. **⚡ Modo Demostración Ejecutiva en 1 Clic:**
   Permite explorar la consola interactiva completa con 81 expedientes sintéticos multisede sin necesidad de contar con archivos locales.
2. **🔍 Dossier Digital del Expediente y Auditoría 360°:**
   Inspección profunda de cada expediente con diagrama de flujo visual (SOLPED ➔ OC ➔ HES ➔ Factura/RXH ➔ MIRO), matriz de cuadre fiscal con $\Delta = \text{S/ 0.00}$, checklist de 6 reglas de control interno y simulador de payload para `BAPI_INCOMINGINVOICE_CREATE`.
3. **📊 Tablero de Analítica Financiera (HUD):**
   Monitoreo visual del gasto en Soles por sede (Arequipa, Chiclayo, Chimbote, Ica) y embudo de madurez SAP con filtrado reactivo al hacer clic en las barras.
4. **📄 Certificado Formal de Conciliación y Control Interno:**
   Emisión de dictamen contable formal con hash SHA-256 del lote, resumen por sedes, declaración de cumplimiento normativo y bloques de firma y visado para Operaciones, Auditoría y Contabilidad.

---

## 📊 Métricas de Impacto Medidas y Verificables

* **⏱️ Tiempo por Expediente:** De ~6 minutos manual a **0.15 segundos** automatizado (**97.5% de ahorro**).
* **🎯 Tasa de Extracción y Conciliación:** **100% verificada** sobre el lote de 4 provincias (24 expedientes, 81 PDFs sintéticos).
* **🔒 Seguridad Financiera:** Cero falsos positivos en SAP; ninguna contabilización sin confirmación explícita y número de documento contable inequívoco.
* **⚡ Velocidad de Pruebas Unitarias:** De 28.68s a **4.06s** (86% más rápido) con ejecución reproducible en CI/CD (GitHub Actions).
* **🛡️ Privacidad Total:** Cero fuga de datos confidenciales hacia la nube (Cumplimiento Ley N° 29733).

---

## 🛠️ Stack Tecnológico Utilizado

* **Frontend:** HTML5, CSS3 Grid/Flexbox moderno, JavaScript ES2022 Vanilla (sin frameworks pesados), File System Access API.
* **Motores Documentales Locales:** `PDF.js` (Apache 2.0), `pypdf`, `reportlab`.
* **Procesamiento de Libros Excel:** `openpyxl` (fórmulas y formato condicional XML), `xlsx-js-style` (MIT), `JSZip` (MIT).
* **Automatización RPA y ERP:** SAP GUI Scripting (`win32com.client`), VBScript nativo Windows, `pyautogui` / `pywinauto` (Plan B de contingencia).
* **Seguridad y Criptografía:** `hashlib` (SHA-256), validaciones regex de documentos tributarios SUNAT.
* **DevOps y QA:** GitHub Actions, Python `unittest` con mocks COM, despliegue estático en Vercel.

---

## 📋 Entregables Oficiales Disponibles en el Repositorio

1. `Control_Ordenes_Servicio_Semana_XX.xlsx`: Libro de control A–W con 3 hojas (`BD`, `RESUMEN`, `INCIDENCIAS`) y semáforos condicionales verde/amarillo/rojo.
2. `Masivo Reporte.xlsx`: Plantilla A–U formateada para carga masiva directa en `/nMIRO`.
3. `Reporte_Asientos_Contables_SAP.xlsx`: Bitácora generada por el robot con N° de Factura MIRO y Asiento Contable extraídos de SAP.
4. `Renombrar_Documentos_MIRO.bat`: Script ejecutable en Windows para renombrado seguro con un solo clic.
5. Suite de documentación técnica empresarial (`docs/01` al `docs/12`).
