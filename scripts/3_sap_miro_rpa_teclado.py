#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SOLUCIÓN 3: RPA ASISTIDO POR TECLADO Y PANTALLA (PLAN B)
=========================================================
Para escenarios donde el área de TI de la empresa tiene bloqueado
o restringido el Scripting interno de SAP GUI.

Simula a un operador humano digitando en el teclado de forma ultra precisa:
1. Enfoca la ventana de SAP GUI.
2. Ingresa a la transacción /nMIRO.
3. Navega con teclas TAB y atajos de teclado oficiales de SAP.
4. Digita Fecha, Referencia, Importe y OC.
5. Contabiliza con Ctrl + S.

Requisitos:
- pip install pyautogui pywinauto openpyxl
"""

import os
import sys
import time
import openpyxl

# Compatibilidad UTF-8 en consola de Windows
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

try:
    import pyautogui
    # Failsafe: Si mueves el mouse a la esquina superior izquierda (0,0), el script se detiene de emergencia
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.3
except ImportError:
    print("❌ Error: Requiere la librería 'pyautogui'. Instálala ejecutando:")
    print("   pip install pyautogui openpyxl")
    sys.exit(1)


def enfocar_sap():
    """Busca y enfoca la ventana activa de SAP GUI."""
    try:
        import pywinauto
        app = pywinauto.Desktop(backend="win32")
        # Buscar ventana con título que contenga SAP
        sap_window = None
        for w in app.windows():
            title = w.window_text()
            if "SAP Easy Access" in title or "SAP" in title or "MIRO" in title:
                sap_window = w
                break
        
        if sap_window:
            sap_window.set_focus()
            time.sleep(0.5)
            print(f"✅ Ventana de SAP enfocada: '{sap_window.window_text()}'")
            return True
        else:
            print("⚠️ No se encontró la ventana de SAP automáticamente.")
            print("   Por favor haz clic en la ventana de SAP para ponerla en primer plano.")
            return False
    except Exception:
        # Fallback sin pywinauto
        print("ℹ️ Enfocando pantalla. Asegúrate de tener SAP en primer plano.")
        return True


def digitar_factura_miro(fecha_doc, referencia, ruc, importe, oc):
    """Ejecuta la secuencia de teclas para una factura en MIRO."""
    print(f"  ⌨️  Digitando: Fecha={fecha_doc} | Ref={referencia} | S/ {importe} | OC={oc}")

    # 1. Abrir MIRO desde la barra de comandos
    pyautogui.hotkey('ctrl', '/')  # Atajo SAP para ir a la barra de comandos
    time.sleep(0.3)
    pyautogui.write('/nMIRO', interval=0.05)
    pyautogui.press('enter')
    time.sleep(1.8)  # Esperar que cargue la transacción

    # 2. Llenar Fecha de Documento
    pyautogui.write(str(fecha_doc), interval=0.05)
    pyautogui.press('tab')

    # 3. Fecha de contabilización
    pyautogui.write(str(fecha_doc), interval=0.05)
    pyautogui.press('tab')

    # 4. Referencia (Factura / RxH)
    pyautogui.write(str(referencia), interval=0.05)
    pyautogui.press('tab')

    # 5. Importe
    importe_str = f"{float(importe):.2f}"
    pyautogui.write(importe_str, interval=0.05)
    pyautogui.press('tab')

    # 6. Moneda PEN
    pyautogui.write('PEN', interval=0.05)
    time.sleep(0.5)

    # 7. Asignar Orden de Compra (OC)
    if oc:
        # En SAP MIRO, presionar Tab hasta el campo N° OC o atajo F6
        pyautogui.press('f6')  # o seleccionar pestaña de referencias
        time.sleep(0.5)
        pyautogui.write(str(oc).strip(), interval=0.05)
        pyautogui.press('enter')
        time.sleep(1.5)

    # 8. Contabilizar
    print("  💾 Contabilizando comprobante (Ctrl + S)...")
    pyautogui.hotkey('ctrl', 's')
    time.sleep(2.0)  # Esperar respuesta de SAP


def procesar_rpa(ruta_excel):
    if not os.path.exists(ruta_excel):
        print(f"❌ No se encontró el archivo: {ruta_excel}")
        return

    wb = openpyxl.load_workbook(ruta_excel, data_only=True)
    ws = wb.active
    rows = []
    for r in range(2, ws.max_row + 1):
        id_cab = ws.cell(r, 1).value
        if id_cab and str(id_cab).upper() != 'TOTAL':
            rows.append({
                'r': r,
                'fecha': ws.cell(r, 2).value,
                'ref': ws.cell(r, 4).value,
                'ruc': ws.cell(r, 6).value,
                'imp': ws.cell(r, 17).value,
                'oc': ws.cell(r, 21).value,
            })

    print("=" * 70)
    print(f"   SOLUCIÓN 3: RPA TECLADO PARA SAP MIRO (PLAN B)")
    print(f"   Total registros a procesar: {len(rows)}")
    print("=" * 70)
    print("\n⚠️ AVISO IMPORTANTE:")
    print("1. Deja la ventana de SAP abierta y maximizada.")
    print("2. No toques el teclado ni el mouse mientras el robot digite.")
    print("3. Si necesitas detener el robot de emergencia, mueve el mouse a la")
    print("   esquina superior izquierda de tu pantalla.\n")

    for i in range(5, 0, -1):
        print(f"⏳ El robot iniciará en {i} segundos... (coloca SAP en pantalla)", end='\r')
        time.sleep(1.0)
    print("\n🚀 INICIANDO AUTOMATIZACIÓN...\n")

    enfocar_sap()

    for idx, item in enumerate(rows, start=1):
        print(f"[{idx}/{len(rows)}] Procesando Fila {item['r']}...")
        digitar_factura_miro(item['fecha'], item['ref'], item['ruc'], item['imp'], item['oc'])
        time.sleep(1.0)

    print("\n" + "=" * 70)
    print("🏁 AUTOMATIZACIÓN RPA COMPLETADA EXITOSAMENTE")
    print("=" * 70)


if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    excel_path = os.path.join(base_dir, "Masivo Reporte.xlsx")
    procesar_rpa(excel_path)
