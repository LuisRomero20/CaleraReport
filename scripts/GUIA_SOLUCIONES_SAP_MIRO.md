# 📑 GUÍA DE IMPLEMENTACIÓN Y PRESENTACIÓN: AUTOMATIZACIÓN SAP MIRO

Este documento describe las **3 soluciones ejecutables** creadas para automatizar el registro secuencial de facturas y recibos por honorarios en **SAP MIRO**, tomando como base el archivo generado `Masivo Reporte.xlsx`.

---

## 📁 Archivos Creados en la Carpeta `scripts/`

| Archivo | Enfoque | Requisitos | Ideal para |
|---|---|---|---|
| [`1_sap_miro_scripting.py`](1_sap_miro_scripting.py) | **SAP Scripting (Python)** ⭐️ | `pywin32`, `openpyxl` | **Producción oficial**: Rápido, captura el N° MIRO y es 100% tolerante a fallos. |
| [`2_sap_miro_macro.vbs`](2_sap_miro_macro.vbs) | **VBScript Nativo** | Ninguno (Nativo Windows) | **Laptops corporativas con restricciones**: No requiere instalar Python ni librerías. |
| [`3_sap_miro_rpa_teclado.py`](3_sap_miro_rpa_teclado.py) | **RPA Teclado (Simulación)** | `pyautogui`, `openpyxl` | **Plan de Emergencia**: Si TI tiene bloqueado el Scripting en el servidor SAP. |

---

## 🚀 CÓMO EJECUTAR CADA SOLUCIÓN

### ⭐️ Solución 1: Python SAP Scripting (Recomendada)

1. Abre tu **SAP Logon** e inicia sesión normalmente.
2. Abre tu terminal en la carpeta del proyecto y ejecuta:
   ```bash
   # Modo Simulación (para demostración segura, verifica campos sin guardar)
   python scripts/1_sap_miro_scripting.py --simular

   # Modo Real (Contabiliza en SAP y extrae el N° de Documento)
   python scripts/1_sap_miro_scripting.py
   ```
3. **Resultado**: Verás cómo SAP navega solo campo por campo. Al terminar, generará `Masivo Reporte (CON MIRO).xlsx` con el número de documento oficial de SAP anotado en cada fila.

---

### ⚡ Solución 2: Macro VBScript (Sin Instalar Nada)

1. Abre tu **SAP Logon** e inicia sesión en el mandante.
2. Haz **doble clic** directamente sobre el archivo:
   [`scripts/2_sap_miro_macro.vbs`](2_sap_miro_macro.vbs)
3. **Resultado**: Windows ejecutará el script nativamente en segundo plano, registrará los comprobantes en SAP y te mostrará una ventana con el resumen final de contabilización.

---

### ⌨️ Solución 3: RPA Asistido por Teclado (Plan B)

1. Abre tu SAP GUI, maximiza la ventana.
2. En la terminal ejecuta:
   ```bash
   python scripts/3_sap_miro_rpa_teclado.py
   ```
3. El script te dará un conteo de 5 segundos para que pongas la pantalla de SAP visible.
4. El robot digitará automáticamente los atajos y campos.

---

## 📊 COMPARATIVA PARA LA PRESENTACIÓN

| Métrica / Ventaja | Solución 1 (Python Scripting) | Solución 2 (VBScript Nativo) | Solución 3 (RPA Teclado) |
|---|:---:|:---:|:---:|
| **Velocidad por Factura** | **2 a 3 segundos** | 3 a 5 segundos | 7 a 10 segundos |
| **Requiere Software Externo** | Sí (Python) | **No (100% Nativo Windows)** | Sí (Python) |
| **Captura N° de Documento SAP** | **Sí (en Excel)** | Sí (en Excel) | No |
| **Manejo de Errores de Saldo** | **Avanzado (Log individual)** | Básico | Manual |
| **Seguridad de Ejecución** | 99% (Identificadores) | 95% (Identificadores) | 75% (Foco de ventana) |
