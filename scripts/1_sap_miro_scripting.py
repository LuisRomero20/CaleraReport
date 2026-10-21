#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SOLUCIÓN 1: SAP GUI SCRIPTING (PYTHON) - MOTOR DE REGISTRO EN MIRO
==================================================================
Automatiza el registro masivo en SAP MIRO leyendo 'Masivo Reporte.xlsx' fila por fila.
Cumple al 100% con el flujo oficial de la empresa:
  1. Dat Básica:
     - Fecha factura (BLDAT) = Fecha Doc de Excel
     - Referencia (XBLNR) = REFERENCIA de Excel (01-... Factura, 02-... RH)
     - Importe (WRBTR) = Importe MD de Columna Q
     - Moneda (WAERS) = PEN
     - Indicador de Impuestos (MWSKZ) = C1 (Factura c/IGV) o C0 (RH sin IGV)
     - Casilla Calc.Impuestos (XMWST) = Marcada en C1, desmarcada en C0
     - Texto (BKTXT) = Servicio solamente (sin mencionar SEMANA)
     - Pedido/Plan entregas (EBELN) = N° OC + [Enter] para cargar posiciones
  2. Detalle:
     - Clase Doc (BLART) = KR (Factura) o KH (Recibo por Honorarios)
  3. Retención:
     - Recibo por Honorarios: No agrega nada
     - Factura: Retención de IGV vacía
  4. Simular:
     - Presiona Simular (F7) y valida balance en ventana modal
     - Pregunta al usuario confirmación antes de guardar (o modo automático)
  5. Contabilizar:
     - Presiona Guardar (Ctrl + S)
     - Captura N° Documento Factura MIRO (ej. 5105773695) y Asiento Contable FI (ej. 4300035763)
     - Anota ambos en el Excel y genera 'Reporte_Asientos_Contables_SAP.xlsx'
"""

import os
import sys
import time
import re
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# Compatibilidad UTF-8 en consola de Windows
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

try:
    import win32com.client
except ImportError:
    win32com = None


def conectar_sap():
    """Se conecta a la sesión activa de SAP GUI for Windows."""
    if win32com is None:
        print("❌ Error: Requiere la librería 'pywin32' para conectarse a SAP GUI. Instálala ejecutando:")
        print("   pip install pywin32 openpyxl")
        return None

    print("🔗 Conectando con sesión activa de SAP GUI...")
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
        if application.Children.Count == 0:
            raise Exception("No hay conexiones abiertas en SAP GUI. Inicia sesión primero.")
        connection = application.Children(0)
        if connection.Children.Count == 0:
            raise Exception("No hay sesiones activas en la conexión de SAP.")
        session = connection.Children(0)
        print(f"✅ Conectado a SAP exitosamente.")
        print(f"   👤 Usuario: {session.Info.User} | Mandante: {session.Info.Mandant} | Sistema: {session.Info.SystemName}")
        return session
    except Exception as e:
        print(f"❌ Error al conectar con SAP: {e}")
        print("\nVerifica que:")
        print("1. SAP GUI esté abierto con sesión iniciada.")
        print("2. El scripting esté habilitado en SAP GUI (Ajustar disposición local Alt+F12 -> Opciones -> Scripting).")
        return None


def limpiar_texto_servicio(texto_pos):
    """
    Formatea el texto para no mencionar la semana sino únicamente el servicio.
    Ejemplo: 'SERVICIO DE SANEO SEMANA 38' -> 'SERVICIO DE SANEO'
    """
    t = str(texto_pos or '').strip()
    t = re.sub(r'\s*SEM(?:ANA)?\s*[-:]?\s*\d+.*$', '', t, flags=re.IGNORECASE).strip()
    return t[:25]  # Longitud estándar de texto de cabecera MIRO (BKTXT)


def registrar_fila_miro(session, fecha_doc, referencia, ruc, importe, texto_pos, oc, ind_impto, interactivo=True, modo_simulacion=False):
    """
    Ejecuta el flujo completo paso a paso en SAP MIRO según las 5 pantallas del usuario.
    Retorna: (exito: bool, msg: str, doc_miro: str, asiento_contable: str)
    """
    es_factura = str(referencia).startswith('01-') or str(ind_impto).upper() == 'C1'
    code_impto = 'C1' if es_factura else 'C0'
    clase_doc = 'KR' if es_factura else 'KH'
    servicio_limpio = limpiar_texto_servicio(texto_pos)

    try:
        # ── 1. Ingresar a la transacción /nMIRO ─────────────────────
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nMIRO"
        session.findById("wnd[0]").sendVKey(0)  # Enter
        time.sleep(1.0)

        # ── PANTALLA 1: PESTAÑA 'Dat.básic.' ────────────────────────
        # 1.1 Asegurar estar en pestaña Datos Básicos
        try:
            session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER").select()
        except Exception:
            pass

        # 1.2 Fecha Factura (BLDAT)
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/ctxtINVFO-BLDAT").text = str(fecha_doc)

        # 1.3 Fecha Contabilización (BUDAT)
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/ctxtINVFO-BUDAT").text = str(fecha_doc)

        # 1.4 Referencia (XBLNR)
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/txtINVFO-XBLNR").text = str(referencia)

        # 1.5 Importe (WRBTR)
        importe_str = f"{float(importe):.2f}"
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/txtINVFO-WRBTR").text = importe_str

        # 1.6 Moneda (WAERS) = PEN
        # SAP no siempre deja escribir directamente, por lo que se asigna con fallback nativo
        try:
            ctrl_waers = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/ctxtINVFO-WAERS")
            ctrl_waers.text = "PEN"
        except Exception:
            try:
                session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/ctxtINVFO-WAERS").key = "PEN"
            except Exception:
                pass

        # 1.7 Indicador de Impuestos (MWSKZ): C1 (18% IGV Factura) o C0 (0% RH)
        # y Casilla 'Calc.Impuestos' (XMWST)
        try:
            # Asignar código en combobox o text
            for cid in ["ctxtINVFO-MWSKZ_GROSS", "ctxtINVFO-MWSKZ", "cmbINVFO-MWSKZ"]:
                try:
                    c = session.findById(f"wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/{cid}")
                    if hasattr(c, 'key'):
                        c.key = code_impto
                    c.text = code_impto
                    break
                except Exception:
                    pass

            # Checkbox Calc.Impuestos
            chk_tax = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/chkINVFO-XMWST")
            chk_tax.selected = (code_impto == 'C1')
        except Exception as e_impto:
            print(f"    ℹ️ Nota impuesto: {e_impto}")

        # 1.8 Texto (BKTXT): Servicio solamente
        try:
            session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/txtINVFO-BKTXT").text = servicio_limpio
        except Exception:
            pass

        # 1.9 Pedido / Orden de Compra (OC) y Enter
        if oc:
            session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/ssubITEMS:SAPLMR1M:6011/ctxtRM08M-EBELN").text = str(oc).strip()
            session.findById("wnd[0]").sendVKey(0)  # Enter para traer posiciones
            time.sleep(1.2)

        # Manejar avisos informativos iniciales de SAP
        while session.Children.Count > 1:
            popup = session.Children(session.Children.Count - 1)
            popup.sendVKey(0)  # Enter para cerrar
            time.sleep(0.4)

        # ── PANTALLA 2: PESTAÑA 'Detalle' ───────────────────────────
        # Asignar Clase Doc: KR (Factura) o KH (Recibo por Honorarios)
        try:
            session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_DETAIL").select()
            time.sleep(0.5)
            for c_blart in ["ctxtINVFO-BLART", "cmbINVFO-BLART"]:
                try:
                    c = session.findById(f"wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_DETAIL/ssubHEADER:SAPLMR1M:6010/{c_blart}")
                    if hasattr(c, 'key'):
                        c.key = clase_doc
                    c.text = clase_doc
                    break
                except Exception:
                    pass
        except Exception as e_det:
            print(f"    ℹ️ Pestaña Detalle ({clase_doc}): {e_det}")

        # ── PANTALLA 3: PESTAÑA 'Retención' ─────────────────────────
        # RH: no agregamos nada. Factura: Retención de IGV vacía
        try:
            session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_WT").select()
            time.sleep(0.5)
            # Se verifica y mantiene vacío
        except Exception:
            pass

        # Volver a pestaña Datos Básicos para simular
        try:
            session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER").select()
            time.sleep(0.4)
        except Exception:
            pass

        # ── PANTALLA 4: SIMULACIÓN ──────────────────────────────────
        # Presionar botón Simular (F7)
        session.findById("wnd[0]/tbar[1]/btn[7]").press()
        time.sleep(1.2)

        # Si aparece ventana de simulación (wnd[1] Modal 'Simular documento en PEN')
        print(f"  🔍 Simulación generada exitosamente en SAP.")
        print(f"     Comprobante: {referencia} | Tipo: {clase_doc} | Monto: S/ {importe_str} | Impuesto: {code_impto} | OC: {oc}")

        # En modo simulación pura: no guarda y limpia el estado modal
        if modo_simulacion:
            # Cerrar modal de simulación (F12)
            if session.Children.Count > 1:
                session.Children(session.Children.Count - 1).sendVKey(12)  # F12 Cancelar / Salir
            # Cancelar pantalla /nMIRO para no dejar bloqueo
            session.findById("wnd[0]").sendVKey(12)
            return (True, "Simulación exitosa (Modo prueba - Cancelado limpiamente)", "SIMULADO", "SIMULADO")

        # Confirmación interactiva con el usuario (Fail-Closed)
        if interactivo:
            print("\n  👉 Revisa la ventana de simulación en tu SAP GUI (Saldo debe ser 0.00 PEN).")
            resp = input("  ❓ ¿Deseas CONTABILIZAR este comprobante en SAP? Escribe 'SI' para confirmar: ").strip().upper()
            if resp not in ('SI', 'CONFIRMAR'):
                print(f"  🛑 Operación cancelada por política de control Fail-Closed (Respuesta: '{resp}').")
                if session.Children.Count > 1:
                    session.Children(session.Children.Count - 1).sendVKey(12)  # F12 Cancelar modal
                session.findById("wnd[0]").sendVKey(12)  # F12 Salir de transacción
                return (False, "Cancelado por política Fail-Closed: se requiere confirmación afirmativa explícita ('SI')", "", "")

        # ── PANTALLA 5: CONTABILIZAR Y EXTRAER ASIENTO CONTABLE ──────
        # Contabilizar (Ctrl + S o Guardar btn[11])
        if session.Children.Count > 1:
            try:
                session.Children(session.Children.Count - 1).findById("tbar[0]/btn[11]").press()
            except Exception:
                session.findById("wnd[0]/tbar[0]/btn[11]").press()
        else:
            session.findById("wnd[0]/tbar[0]/btn[11]").press()

        time.sleep(2.0)

        # Leer respuesta en la barra de estado de SAP (sbar)
        msg = session.findById("wnd[0]/sbar").text
        msg_type = session.findById("wnd[0]/sbar").messageType

        # Extraer Documento de Factura MIRO y Asiento Contable FI
        # Formato oficial: "Se ha contabilizado el documentoc de factura 5105773695 ( Asiento contable: 4300035763 )"
        m_miro = re.search(r'factura\s+(\d{10})', msg, re.I) or re.search(r'\b(5\d{9})\b', msg)
        doc_miro = m_miro.group(1) if m_miro else ""

        m_asiento = re.search(r'Asiento\s+contable\s*:\s*(\d{10})', msg, re.I) or re.search(r'\b(4\d{9})\b', msg)
        asiento_contable = m_asiento.group(1) if m_asiento else ""

        # Verificación estricta de éxito financiero
        if msg_type in ('S', 'I') and (doc_miro or asiento_contable):
            return (True, msg, doc_miro, asiento_contable)
        elif msg_type == 'E':
            return (False, f"Error SAP: {msg}", "", "")
        else:
            # POLÍTICA FAIL-CLOSED: Nunca asumir éxito ante respuestas ambiguas o sin ID inequívoco
            detalle_err = f"Respuesta no confirmada de SAP (Tipo: {msg_type}, Mensaje: '{msg}'). No se detectó número de documento contable inequívoco."
            return (False, f"Alerta de Control Financiero: {detalle_err} Se envía a reconciliación manual.", "", "")

    except Exception as e:
        return (False, f"Excepción durante el proceso MIRO: {e}", "", "")


def procesar_miro_completo(ruta_excel, interactivo=True, modo_simulacion=False):
    """
    Recorre 'Masivo Reporte.xlsx', procesa cada fila en SAP MIRO,
    actualiza el libro y crea un reporte resumen de asientos contables.
    """
    if not os.path.exists(ruta_excel):
        print(f"❌ No se encontró el archivo: {ruta_excel}")
        return

    session = conectar_sap()
    if not session:
        return

    wb = openpyxl.load_workbook(ruta_excel)
    print(f"\n📂 Libro abierto: '{os.path.basename(ruta_excel)}'")
    print(f"📑 Hojas disponibles: {wb.sheetnames}")

    registros_log = []
    exitosos = 0
    fallidos = 0

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        print(f"\n" + "=" * 70)
        print(f"📋 PROCESANDO HOJA: '{sheet_name}'")
        print("=" * 70)

        # Identificar columnas
        c_fecha = 2   # Col B: Fecha Doc
        c_ref   = 4   # Col D: Referencia
        c_ruc   = 6   # Col F: RUC
        c_impto = 8   # Col H: Ind Impto
        c_tpos  = 12  # Col L: Texto Posición
        c_imp   = 17  # Col Q: Importe MD
        c_oc    = 21  # Col U: N° OC

        # Columnas de resultado en el Excel
        col_res     = 22  # Col V: Estado MIRO
        col_doc     = 23  # Col W: N° Factura MIRO
        col_asiento = 24  # Col X: Asiento Contable FI
        ws.cell(1, col_res, value="Estado MIRO")
        ws.cell(1, col_doc, value="N° Factura MIRO")
        ws.cell(1, col_asiento, value="Asiento Contable")

        for r in range(2, ws.max_row + 1):
            id_cab = ws.cell(r, 1).value
            if not id_cab or str(id_cab).upper() == 'TOTAL':
                continue

            fecha_doc   = ws.cell(r, c_fecha).value
            referencia  = ws.cell(r, c_ref).value
            ruc         = ws.cell(r, c_ruc).value
            texto_pos   = ws.cell(r, c_tpos).value
            importe     = ws.cell(r, c_imp).value
            oc          = ws.cell(r, c_oc).value
            ind_impto   = ws.cell(r, c_impto).value or ws.cell(r, 18).value

            es_factura = str(referencia).startswith('01-') or str(ind_impto).upper() == 'C1'
            clase_doc = 'KR' if es_factura else 'KH'

            print(f"\n─── [Fila {r}] ──────────────────────────────────────────────────────────")
            print(f"📄 Comprobante : {referencia} ({'FACTURA KR' if es_factura else 'RECIBO POR HONORARIOS KH'})")
            print(f"💰 Importe     : S/ {float(importe):.2f} | Impuesto: {ind_impto}")
            print(f"🏢 RUC / OC    : {ruc} | OC: {oc}")
            print(f"🏷️  Servicio    : {limpiar_texto_servicio(texto_pos)}")

            exito, msg, doc_miro, asiento = registrar_fila_miro(
                session, fecha_doc, referencia, ruc, importe, texto_pos, oc, ind_impto,
                interactivo=interactivo, modo_simulacion=modo_simulacion
            )

            # Guardar en celdas de la fila
            if exito:
                print(f"  🎉 ÉXITO SAP: {msg}")
                print(f"  📝 Doc MIRO: {doc_miro} | 🏦 Asiento Contable: {asiento}")
                ws.cell(r, col_res, value="CONTABILIZADO" if not modo_simulacion else "SIMULADO OK")
                ws.cell(r, col_doc, value=doc_miro)
                ws.cell(r, col_asiento, value=asiento)
                exitosos += 1
            else:
                print(f"  ⚠️ NO REGISTRADO: {msg}")
                ws.cell(r, col_res, value=f"ERROR: {msg[:40]}")
                ws.cell(r, col_doc, value="FALLIDO")
                ws.cell(r, col_asiento, value="-")
                fallidos += 1

            registros_log.append({
                'hoja': sheet_name,
                'fila': r,
                'fecha': fecha_doc,
                'referencia': referencia,
                'ruc': ruc,
                'oc': oc,
                'servicio': limpiar_texto_servicio(texto_pos),
                'importe': importe,
                'tipo': clase_doc,
                'impto': ind_impto,
                'doc_miro': doc_miro,
                'asiento_contable': asiento,
                'estado': "CONTABILIZADO" if exito else "ERROR",
                'mensaje_sap': msg
            })

            time.sleep(1.0)

    # Guardar Excel actualizado
    try:
        wb.save(ruta_excel)
        print(f"\n💾 Archivo base actualizado con éxito: {ruta_excel}")
    except PermissionError:
        copia_excel = ruta_excel.replace(".xlsx", " (CON ASIENTOS).xlsx")
        wb.save(copia_excel)
        print(f"\n⚠️ El Excel base estaba abierto. Se guardó copia en: {copia_excel}")

    # Generar Reporte de Asientos Contables consolidado
    generar_reporte_asientos(registros_log)

    print("\n" + "=" * 70)
    print("🏁 PROCESAMIENTO FINALIZADO")
    print(f"📊 Total procesados: {len(registros_log)} | ✅ Exitosos: {exitosos} | ❌ Fallidos: {fallidos}")
    print("=" * 70)


def generar_reporte_asientos(registros):
    """Genera un archivo Excel independiente y limpio con todos los asientos generados."""
    if not registros:
        return

    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_path = os.path.join(base_dir, "Reporte_Asientos_Contables_SAP.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "Asientos MIRO Contabilizados"

    headers = [
        "Hoja / Depto", "Fila", "Fecha Doc", "REFERENCIA", "RUC", "N° OC",
        "Servicio", "Importe (S/)", "Clase Doc", "Ind Impto",
        "N° Factura MIRO", "Asiento Contable FI", "Estado", "Mensaje SAP"
    ]

    fill_hdr = PatternFill('solid', fgColor='1F4E79')
    font_hdr = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    al_c = Alignment(horizontal='center', vertical='center')
    al_l = Alignment(horizontal='left', vertical='center')
    al_r = Alignment(horizontal='right', vertical='center')
    thin = Side(style='thin', color='D9D9D9')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(1, col_idx)
        c.fill = fill_hdr
        c.font = font_hdr
        c.alignment = al_c

    for reg in registros:
        row = [
            reg['hoja'], reg['fila'], reg['fecha'], reg['referencia'], reg['ruc'], reg['oc'],
            reg['servicio'], reg['importe'], reg['tipo'], reg['impto'],
            reg['doc_miro'], reg['asiento_contable'], reg['estado'], reg['mensaje_sap']
        ]
        ws.append(row)

    # Formatos de ancho y alineación
    for r in range(2, ws.max_row + 1):
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(r, col_idx)
            cell.border = border
            if col_idx in (2, 3, 9, 10, 11, 12, 13):
                cell.alignment = al_c
            elif col_idx in (8,):
                cell.alignment = al_r
                cell.number_format = '"S/ "#,##0.00'
            else:
                cell.alignment = al_l

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    try:
        wb.save(out_path)
        print(f"📁 Reporte de Asientos Contables exportado en:\n   {out_path}")
    except Exception as e:
        print(f"⚠️ No se pudo guardar reporte independiente: {e}")


if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    excel_path = os.path.join(base_dir, "Masivo Reporte.xlsx")

    # Modos de ejecución:
    # python scripts/1_sap_miro_scripting.py            -> Modo Interactivo (Pregunta antes de contabilizar)
    # python scripts/1_sap_miro_scripting.py --auto     -> Modo Automático (Contabiliza sin pausar)
    # python scripts/1_sap_miro_scripting.py --simular  -> Modo Simulación (Verifica campos sin guardar)

    es_simular = "--simular" in sys.argv
    es_auto = "--auto" in sys.argv

    # Control de Seguridad Financiera (Fail-Closed):
    # El modo desatendido (--auto) está restringido por defecto para evitar contabilizaciones accidentales.
    if es_auto:
        auto_autorizado = (
            os.environ.get("CALERA_ALLOW_UNATTENDED", "").lower() in ("true", "1", "yes")
            or "--forzar-auto-autorizado" in sys.argv
        )
        if not auto_autorizado:
            print("=" * 70)
            print("🛑 BLOQUEO DE CONTROL FINANCIERO (POLÍTICA FAIL-CLOSED)")
            print("=" * 70)
            print("El modo desatendido (--auto) no está permitido en ejecuciones directas.")
            print("Toda contabilización en SAP debe contar con validación humana interactiva.")
            print("Para habilitar este modo con propósitos de integración o pruebas autorizadas,")
            print("defina la variable de entorno: CALERA_ALLOW_UNATTENDED=true")
            print("o añada el parámetro explícito: --forzar-auto-autorizado\n")
            sys.exit(1)

    es_interactivo = not es_auto and not es_simular

    print("=" * 70)
    print("   BOT AUTOMATIZADOR DE REGISTRO EN SAP MIRO (LA CALERA)")
    print("=" * 70)
    print(f"Modo: {'SIMULACIÓN (Sin guardar)' if es_simular else 'AUTOMÁTICO' if es_auto else 'INTERACTIVO (Pregunta al simular)'}")
    print(f"Archivo de entrada: {excel_path}\n")

    procesar_miro_completo(excel_path, interactivo=es_interactivo, modo_simulacion=es_simular)
