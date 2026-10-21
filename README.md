# 🏭 CALERA — Enterprise Document Automation & SAP Integration Suite

[![CI Test Suite & Security](https://github.com/LuisRomero20/CaleraReport/actions/workflows/ci.yml/badge.svg)](https://github.com/LuisRomero20/CaleraReport/actions/workflows/ci.yml)
[![Vercel Deployment](https://img.shields.io/badge/Vercel-Production%20Ready-000000?logo=vercel)](https://calerareport-ten.vercel.app)
[![Python Version](https://img.shields.io/badge/python-3.11%20%7C%203.12%20%7C%203.13-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Zero Data Leak](https://img.shields.io/badge/Privacy-100%25%20Client--Side-brightgreen.svg)](#seguridad-y-privacidad)
[![SAP Financial Control](https://img.shields.io/badge/SAP%20MIRO-Fail--Closed-red.svg)](#controles-financieros-sap-fail-closed)

Plataforma empresarial de **ingeniería documental, conciliación contable determinista y automatización SAP MIRO** para **Sociedad Agrícola La Calera**. Transforma expedientes desordenados de servicios por semana y provincia en reportes oficiales certificados y automatiza la contabilización en SAP GUI bajo estrictos controles financieros *fail-closed*.

🔗 **Acceso Rápido:**  
* **Aplicación Web en Producción:** [https://calerareport-ten.vercel.app](https://calerareport-ten.vercel.app)  
* **Repositorio de Código:** [https://github.com/LuisRomero20/CaleraReport](https://github.com/LuisRomero20/CaleraReport)  
* **Documentación Ejecutiva en Notion:** [Espacio de Trabajo CaleraReport](https://app.notion.com/b34b8038ebdb48c2b04059f945a78496)

---

## 🌟 Capacidades Flagship de Nivel Directivo

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                       CALERAREPORT ENTERPRISE SUITE                         │
├───────────────────────┬──────────────────────────┬──────────────────────────┤
│  ⚡ Modo Demo 1-Clic  │  🔍 Dossier Digital 360° │  📊 HUD Analítica Gasto  │
│  81 PDFs multisede en │  Auditoría y trazabilidad│  Desglose S/ por provincia│
│  memoria instantánea  │  y simulador SAP BAPI    │  y embudo de madurez SAP │
├───────────────────────┴──────────────────────────┴──────────────────────────┤
│              📄 Certificado Formal de Control Interno (SHA-256)             │
│        Dictamen contable imprimible con firmas y sellos de aprobación       │
└─────────────────────────────────────────────────────────────────────────────┘
```

1. **⚡ Modo Demostración Ejecutiva (1-Clic):**  
   Permite explorar la consola interactiva completa con 81 expedientes sintéticos multisede (**Arequipa, Chiclayo, Chimbote, Ica**) sin necesidad de contar con archivos locales.
2. **🔍 Dossier Digital del Expediente y Auditoría 360°:**  
   Inspección profunda de cada expediente con diagrama visual de 5 nodos:  
   $$\text{SOLPED} \longrightarrow \text{OC} \longrightarrow \text{HES} \longrightarrow \text{COMPROBANTE} \longrightarrow \text{MIRO SAP}$$  
   Incluye matriz fiscal de cuadre ($\Delta = \text{S/ 0.00}$), checklist de 6 reglas de control interno y simulador de payload JSON para `BAPI_INCOMINGINVOICE_CREATE`.
3. **📊 Tablero de Analítica Financiera (HUD):**  
   Monitoreo visual del gasto en Soles por sede y embudo de madurez SAP con filtrado reactivo al hacer clic en las barras.
4. **📄 Certificado Formal de Conciliación y Control Interno:**  
   Emisión de dictamen contable formal imprimible con hash SHA-256 del lote, resumen por sedes, declaración de cumplimiento normativo y bloques de firma y visado.

---

## 🏗️ Arquitectura y Principio de Privacidad Zero-Leak

Todo el procesamiento documental ocurre **exclusivamente en el navegador del cliente o en scripts locales de Python**.
* **Cero Fuga a la Nube:** Ningún documento PDF, razón social ni importe financiero se transmite a APIs externas ni servicios de terceros.
* **100% Offline:** Todas las bibliotecas dependientes (`PDF.js`, `SheetJS / xlsx`, `JSZip`) están empaquetadas localmente en `/vendor`.
* **Single Source of Truth:** Los documentos PDF son la única fuente de la verdad; nunca se utilizan archivos Excel previos como filtros.

---

## 📊 Entregables Oficiales Generados

| Entregable | Columnas | Hojas / Formato | Propósito |
|---|---|---|---|
| **Maestro de Control A–W** | 23 (`ID` a `ESTATUS`) | `BD`, `RESUMEN`, `INCIDENCIAS` | Consolidado general de expedientes con fórmulas nativas vivas y semáforos condicionales (`3TrafficLights1`). |
| **Masivo SAP MIRO** | 21 (`A` a `U`) | Pestañas por Sede (`Arequipa`, etc.) | Formato histórico compatible con carga masiva y scripts de `/nMIRO` (Columna U: N° OC). |
| **Certificado de Auditoría** | Imprimible A4 | Hash SHA-256 / Firmas | Dictamen contable para Jefatura de Operaciones, Auditoría y Contabilidad. |
| **ZIP con Renombrado MIRO** | N/A | Archivos `{MIRO} - {nombre}.pdf` | Renombrado masivo con bitácora JSON y reversión criptográfica garantizada. |

---

## 🛡️ Controles Financieros SAP Fail-Closed

Los scripts de automatización SAP (`scripts/1_sap_miro_scripting.py`, `scripts/2_sap_miro_macro.vbs`, `scripts/3_sap_miro_rpa_teclado.py`) implementan una estricta política de tolerancia cero:
* **Confirmación Afirmativa Obligatoria:** Requiere autorización interactiva explícita (`SI` o `CONFIRMAR`). Se bloquea la tecla enter vacía o respuestas ambiguas.
* **Verificación de Saldo 0.00 PEN:** Simulación previa obligatoria en SAP GUI antes de contabilizar.
* **Validación Numérica de 10 Dígitos:** Solo se aceptan números contables válidos (`5\d{9}` y `4\d{9}`).
* **Modo Desatendido Blindado:** El flag `--auto` se rechaza salvo autorización expresa por variable de entorno (`CALERA_ALLOW_UNATTENDED=true`).

---

## 🚀 Guía de Inicio Rápido

### Opción A: Interfaz Web

1. Abre `index.html` en Chrome, Edge o Firefox. También puede servirse localmente:
   ```powershell
   python -m http.server 8000
   ```
2. Visita `http://localhost:8000` o ingresa a [calerareport-ten.vercel.app](https://calerareport-ten.vercel.app).
3. Haz clic en **`⚡ Cargar Demo Interactivo (81 PDFs)`** para probar el sistema inmediatamente, o arrastra carpetas reales de expedientes.

### Opción B: Motor CLI en Python

1. Instala dependencias:
   ```powershell
   pip install -r scripts/requirements.txt
   ```
2. Genera el reporte consolidado desde expedientes:
   ```powershell
   python scripts/control_ordenes.py fixtures/synthetic_expedientes --salida Maestro_Consolidado.xlsx
   ```
3. Ejecuta la simulación de carga SAP:
   ```powershell
   python scripts/1_sap_miro_scripting.py --simular
   ```

---

## 🧪 Pruebas Automatizadas y CI

La suite contiene 21 pruebas automatizadas que cubren extracción, lógica de negocio, semáforos, validación de tolerancias y controles de seguridad financiera:

```powershell
python -m unittest discover tests -v
```

* **Tiempo de Ejecución:** ~4.0 segundos.
* **Matriz CI:** Validado en GitHub Actions sobre **Ubuntu Latest** y **Windows Latest** para Python 3.11, 3.12 y 3.13.

---

## 📚 Índice de Especificaciones Técnicas

| Documento | Título y Enlace |
|---|---|
| `docs/01` | [Enterprise Baseline & Technical Debt Audit](docs/01-enterprise-baseline.md) |
| `docs/02` | [Business Process & Functional Requirements](docs/02-business-process-and-requirements.md) |
| `docs/03` | [System Architecture & Privacy Design](docs/03-system-architecture.md) |
| `docs/04` | [Document Data Model & State Transitions](docs/04-document-data-model.md) |
| `docs/05` | [Extraction, Matching & Reconciliation Rules](docs/05-extraction-and-reconciliation.md) |
| `docs/06` | [Excel Output Contracts & Formula Specifications](docs/06-excel-output-contracts.md) |
| `docs/07` | [SAP MIRO Integration & Fail-Closed Financial Controls](docs/07-sap-integration-and-controls.md) |
| `docs/08` | [Document Rename Engine & Rollback Verification](docs/08-document-rename-and-rollback.md) |
| `docs/09` | [Security, Privacy & Data Protection Compliance](docs/09-security-and-privacy.md) |
| `docs/10` | [Test Strategy, Evidence & QA Automation](docs/10-test-results-and-evidence.md) |
| `docs/11` | [Deployment Architecture, Operations & Runbook](docs/11-deployment-and-operations.md) |
| `docs/12` | [Enterprise Transformation Case Study](docs/12-enterprise-case-study.md) |

---

## ⚖️ Licencia y Confidencialidad

Este proyecto se distribuye bajo licencia MIT. Todas las razones sociales, nombres y números en las pruebas públicas y fixtures son 100% ficticios y sintéticos para garantizar estricto cumplimiento de la Ley N° 29733 de Protección de Datos Personales.
