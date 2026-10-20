#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generador de las 3 Guías PDF Oficiales para la Laptop de La Calera
==================================================================
Crea 3 documentos PDF profesionales, listos para imprimir o compartir:
1. Guia_Opcion_1_Python_SAP_Scripting.pdf
2. Guia_Opcion_2_VBScript_Macro_Nativo.pdf
3. Guia_Opcion_3_RPA_Teclado_Simulacion.pdf
"""

import os
import sys

# Compatibilidad UTF-8 en consola de Windows
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas

# ── Canvas numerado para pie de página "Página X de Y" ─────────
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, total_pages):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#64748b"))

        # Encabezado (páginas > 1)
        if self._pageNumber > 1:
            self.drawString(40, letter[1] - 30, "LA CALERA · Sistema de Automatización SAP MIRO")
            self.setStrokeColor(colors.HexColor("#e2e8f0"))
            self.setLineWidth(0.5)
            self.line(40, letter[1] - 34, letter[0] - 40, letter[1] - 34)

        # Pie de página
        self.setStrokeColor(colors.HexColor("#e2e8f0"))
        self.setLineWidth(0.5)
        self.line(40, 40, letter[0] - 40, 40)
        self.drawString(40, 28, "Confidencial - Uso Interno La Calera S.A.C.")
        page_str = f"Página {self._pageNumber} de {total_pages}"
        self.drawRightString(letter[0] - 40, 28, page_str)
        self.restoreState()


def get_custom_styles():
    styles = getSampleStyleSheet()

    c_primary = colors.HexColor("#1b4d24")   # Verde Corporativo Calera
    c_dark = colors.HexColor("#0f172a")
    c_muted = colors.HexColor("#475569")
    c_code_bg = colors.HexColor("#f1f5f9")

    styles.add(ParagraphStyle(
        'DocTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=20,
        leading=24,
        textColor=c_primary,
        spaceAfter=4
    ))
    styles.add(ParagraphStyle(
        'DocSubtitle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=11,
        leading=15,
        textColor=c_muted,
        spaceAfter=14
    ))
    styles.add(ParagraphStyle(
        'SecHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=13,
        leading=17,
        textColor=c_primary,
        spaceBefore=14,
        spaceAfter=6,
        keepWithNext=True
    ))
    styles.add(ParagraphStyle(
        'SubSecHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10.5,
        leading=14,
        textColor=c_dark,
        spaceBefore=8,
        spaceAfter=4,
        keepWithNext=True
    ))
    styles.add(ParagraphStyle(
        'BodyDark',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=13,
        textColor=c_dark,
        spaceAfter=6
    ))
    styles.add(ParagraphStyle(
        'BodyBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=13,
        textColor=c_dark,
        spaceAfter=4
    ))
    styles.add(ParagraphStyle(
        'BulletItem',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=12,
        textColor=c_dark,
        leftIndent=14,
        firstLineIndent=-10,
        spaceAfter=3
    ))
    styles.add(ParagraphStyle(
        'CodeBlock',
        parent=styles['Normal'],
        fontName='Courier',
        fontSize=8,
        leading=11,
        textColor=colors.HexColor("#0f172a"),
        backColor=c_code_bg,
        borderPadding=6,
        spaceBefore=4,
        spaceAfter=6
    ))
    styles.add(ParagraphStyle(
        'CalloutText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=12,
        textColor=colors.HexColor("#166534")
    ))
    styles.add(ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=11,
        textColor=colors.white,
        alignment=1
    ))
    styles.add(ParagraphStyle(
        'TableCell',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=11,
        textColor=c_dark
    ))
    styles.add(ParagraphStyle(
        'TableCellBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=11,
        textColor=c_dark
    ))
    return styles


def crear_caja_aviso(texto, styles, tipo="info"):
    bg = colors.HexColor("#f0fdf4") if tipo == "info" else colors.HexColor("#fef2f2")
    border_color = colors.HexColor("#86efac") if tipo == "info" else colors.HexColor("#fca5a5")
    p = Paragraph(f"<b>📌 IMPORTANTE:</b> {texto}", styles['CalloutText'])
    t = Table([[p]], colWidths=[letter[0] - 80])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), bg),
        ('BOX', (0,0), (-1,-1), 1, border_color),
        ('PADDING', (0,0), (-1,-1), 8),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    return t


# ==============================================================================
# 1. GUÍA OPCIÓN 1: PYTHON SAP SCRIPTING (RECOMENDADA)
# ==============================================================================
def generar_pdf_opcion_1(ruta_salida):
    styles = get_custom_styles()
    doc = SimpleDocTemplate(
        ruta_salida,
        pagesize=letter,
        leftMargin=40, rightMargin=40,
        topMargin=45, bottomMargin=45
    )
    story = []

    # Encabezado
    story.append(Paragraph("LA CALERA · MANUAL OFICIAL DE AUTOMATIZACIÓN", styles['DocSubtitle']))
    story.append(Paragraph("Guía 1: Bot SAP MIRO (Python Scripting) ⭐️", styles['DocTitle']))
    story.append(Paragraph("<b>Solución Oficial y Recomendada:</b> Automatización directa de alta velocidad con confirmación interactiva, control de impuestos (C1/C0) y extracción del Asiento Contable oficial de SAP.", styles['DocSubtitle']))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1b4d24"), spaceAfter=10))

    # Sección 1
    story.append(Paragraph("1. ¿Por qué es la solución recomendada para La Calera?", styles['SecHeader']))
    story.append(Paragraph(
        "Esta solución interactúa de forma directa mediante la API COM interna de SAP GUI for Windows. "
        "A diferencia de un operador manual o un macro de pantalla, esta opción lee y digita a nivel de objetos internos "
        "(identificadores de control), lo que garantiza:", styles['BodyDark']
    ))
    story.append(Paragraph("• <b>Inmune a interferencias de mouse</b>: Puedes mover el cursor o consultar otras aplicaciones mientras el bot trabaja.", styles['BulletItem']))
    story.append(Paragraph("• <b>Manejo exacto de Comboboxes</b>: Asigna <code>PEN</code> y los códigos <code>C1</code> (Facturas con IGV) o <code>C0</code> (RH sin IGV) sin fallos.", styles['BulletItem']))
    story.append(Paragraph("• <b>Modo Interactivo con Confirmación</b>: Abre la simulación de SAP, valida que el saldo cuadre en <code>0.00</code> y le pregunta al operador antes de contabilizar.", styles['BulletItem']))
    story.append(Paragraph("• <b>Extracción del Asiento Contable</b>: Captura tanto el N° de Documento MIRO (ej. <code>5105773695</code>) como el Asiento Contable FI (ej. <code>4300035763</code>) y los anota en el Excel.", styles['BulletItem']))

    # Sección 2
    story.append(Paragraph("2. Requisitos Previos en la Laptop de La Calera", styles['SecHeader']))
    story.append(Paragraph("1. <b>SAP GUI for Windows</b> instalado y con acceso a la <b>Sociedad 1001 (LA CALERA S.A.C.)</b>.", styles['BulletItem']))
    story.append(Paragraph("2. <b>Python 3.10 o superior</b> instalado en la máquina (marcar <i>'Add Python to PATH'</i> al instalar).", styles['BulletItem']))
    story.append(Paragraph("3. <b>Librerías necesarias</b>: Ejecutar en el símbolo del sistema (CMD o PowerShell):", styles['BulletItem']))
    story.append(Paragraph("<code>pip install pywin32 openpyxl</code>", styles['CodeBlock']))

    # Sección 3
    story.append(Paragraph("3. Configuración Obligatoria: Habilitar Scripting en SAP Logon", styles['SecHeader']))
    story.append(Paragraph("Para que el script pueda comunicarse con SAP GUI, se debe activar la opción en el cliente:", styles['BodyDark']))
    story.append(Paragraph("1. Abre tu ventana de <b>SAP Logon</b>.", styles['BulletItem']))
    story.append(Paragraph("2. Haz clic en el ícono de la esquina superior izquierda y selecciona <b>Opciones...</b> (o presiona <code>Alt + F12</code> dentro de SAP).", styles['BulletItem']))
    story.append(Paragraph("3. En el árbol de opciones ve a: <b>Accesibilidad y Scripting</b> → <b>Scripting</b>.", styles['BulletItem']))
    story.append(Paragraph("4. Marca la casilla: <b>☑ Habilitar scripting</b>.", styles['BulletItem']))
    story.append(Paragraph("5. <i>Recomendado para evitar alertas molestas</i>: Desmarca las casillas <b>☐ Notificar cuando un script se conecte a SAP GUI</b> y <b>☐ Notificar cuando un script abra una conexión</b>.", styles['BulletItem']))
    story.append(Paragraph("6. Haz clic en <b>Aceptar</b>.", styles['BulletItem']))

    story.append(Spacer(1, 4))
    story.append(crear_caja_aviso("Debes iniciar sesión en SAP GUI normalmente y entrar a tu mandante antes de lanzar el bot.", styles))
    story.append(Spacer(1, 6))

    # Sección 4
    story.append(Paragraph("4. Flujo Automatizado de 5 Pantallas en SAP MIRO", styles['SecHeader']))
    
    t_data = [
        [Paragraph("Pantalla / Pestaña", styles['TableHeader']), Paragraph("Campos Llenados", styles['TableHeader']), Paragraph("Regla Aplicada por el Bot", styles['TableHeader'])],
        [
            Paragraph("<b>1. Dat. Básica</b><br/>(Cabecera)", styles['TableCellBold']),
            Paragraph("• Fecha Factura (B)<br/>• Referencia (D)<br/>• Importe Total (Q)<br/>• Moneda (PEN)<br/>• Cal.Impuestos (H)<br/>• Texto (L)<br/>• N° OC (U)", styles['TableCell']),
            Paragraph("Asigna <code>PEN</code> nativamente. Asigna <code>C1</code> y marca <code>Calc.Impuestos</code> si es Factura (con IGV 18%). Asigna <code>C0</code> si es RH. Limpia el texto de posición dejando únicamente el servicio (sin mencionar SEMANA). Escribe la OC y da <code>[Enter]</code> para cargar la posición y encender el semáforo verde (<code>Saldo 0.00</code>).", styles['TableCell'])
        ],
        [
            Paragraph("<b>2. Detalle</b>", styles['TableCellBold']),
            Paragraph("• Clase de Documento", styles['TableCell']),
            Paragraph("Asigna <b><code>KR</code></b> (Factura acreedor) para Facturas o <b><code>KH</code></b> para Recibos por Honorarios.", styles['TableCell'])
        ],
        [
            Paragraph("<b>3. Retención</b>", styles['TableCellBold']),
            Paragraph("• Retención de IGV", styles['TableCell']),
            Paragraph("Verifica que no se agregue retención si es RH y deja vacía la parte de retención de IGV para Facturas.", styles['TableCell'])
        ],
        [
            Paragraph("<b>4. Simular</b>", styles['TableCellBold']),
            Paragraph("• Balance contable", styles['TableCell']),
            Paragraph("Presiona botón Simular (<code>F7</code>). Valida que cuadre en <code>0.00</code>. En modo interactivo, <b>pregunta en consola al usuario si desea contabilizar</b> antes de proceder.", styles['TableCell'])
        ],
        [
            Paragraph("<b>5. Guardar</b>", styles['TableCellBold']),
            Paragraph("• Asiento Contable", styles['TableCell']),
            Paragraph("Presiona <code>Ctrl + S</code>. Captura de la barra de estado el N° Factura MIRO (ej. <code>5105773695</code>) y el Asiento Contable FI (ej. <code>4300035763</code>) y los registra en el Excel.", styles['TableCell'])
        ]
    ]
    t_flujo = Table(t_data, colWidths=[100, 140, 290])
    t_flujo.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1b4d24")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#f8fafc")]),
        ('PADDING', (0,0), (-1,-1), 5),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(t_flujo)

    # Sección 5
    story.append(Paragraph("5. Modos de Ejecución Paso a Paso", styles['SecHeader']))
    story.append(Paragraph("Abre tu terminal en la carpeta del proyecto donde se encuentra <code>Masivo Reporte.xlsx</code>:", styles['BodyDark']))
    
    story.append(Paragraph("<b>Modo 1: Interactivo (Recomendado para revisión y visto bueno):</b>", styles['BodyBold']))
    story.append(Paragraph("Llena cada pantalla, realiza la simulación en SAP y te pregunta en la consola antes de guardar:", styles['BodyDark']))
    story.append(Paragraph("<code>python scripts/1_sap_miro_scripting.py</code>", styles['CodeBlock']))

    story.append(Paragraph("<b>Modo 2: Automático en Lote (Para procesar muchas facturas sin pausas):</b>", styles['BodyBold']))
    story.append(Paragraph("<code>python scripts/1_sap_miro_scripting.py --auto</code>", styles['CodeBlock']))

    story.append(Paragraph("<b>Modo 3: Simulación Pura (Para pruebas / demo, 100% seguro sin guardar en SAP):</b>", styles['BodyBold']))
    story.append(Paragraph("<code>python scripts/1_sap_miro_scripting.py --simular</code>", styles['CodeBlock']))

    # Sección 6
    story.append(Paragraph("6. Archivos de Salida Generados", styles['SecHeader']))
    story.append(Paragraph("• <b><code>Masivo Reporte.xlsx</code></b>: Se completa agregando al final las columnas <code>Estado MIRO</code>, <code>N° Factura MIRO</code> y <code>Asiento Contable FI</code>.", styles['BulletItem']))
    story.append(Paragraph("• <b><code>Reporte_Asientos_Contables_SAP.xlsx</code></b>: Archivo de auditoría nuevo que consolida la lista de comprobantes con su N° OC, fecha, importe y número de asiento contable para entregar a Contabilidad.", styles['BulletItem']))

    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"✅ Generado: {ruta_salida}")


# ==============================================================================
# 2. GUÍA OPCIÓN 2: VBSCRIPT MACRO NATIVO (CERO INSTALACIONES)
# ==============================================================================
def generar_pdf_opcion_2(ruta_salida):
    styles = get_custom_styles()
    doc = SimpleDocTemplate(
        ruta_salida,
        pagesize=letter,
        leftMargin=36, rightMargin=36,
        topMargin=32, bottomMargin=32
    )
    story = []

    # Encabezado
    story.append(Paragraph("LA CALERA · MANUAL OFICIAL DE AUTOMATIZACIÓN", styles['DocSubtitle']))
    story.append(Paragraph("Guía 2: Macro Nativo VBScript (Cero Instalaciones) ⚡", styles['DocTitle']))
    story.append(Paragraph("<b>Solución para Laptops Corporativas con Restricciones:</b> No requiere instalar Python, administradores ni librerías externas. Funciona con doble clic directo en Windows.", styles['DocSubtitle']))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1b4d24"), spaceAfter=8))

    # Sección 1
    story.append(Paragraph("1. ¿Cuándo utilizar la Macro VBScript?", styles['SecHeader']))
    story.append(Paragraph(
        "Si la laptop asignada por La Calera tiene políticas de seguridad estrictas donde TI no permite instalar "
        "software adicional, Python ni librerías externas, esta solución es ideal porque utiliza el motor nativo "
        "<b>Windows Script Host (WScript)</b>, presente por defecto en todas las computadoras con Windows.", styles['BodyDark']
    ))
    story.append(Paragraph("• <b>Cero instalaciones</b>: No necesitas permisos de administrador ni terminal.", styles['BulletItem']))
    story.append(Paragraph("• <b>Ejecución con doble clic</b>: Simplemente haces doble clic sobre el archivo <code>2_sap_miro_macro.vbs</code>.", styles['BulletItem']))
    story.append(Paragraph("• <b>Mismo flujo contable oficial</b>: Aplica Facturas (KR, C1 con IGV, Calc.Impuestos) y Recibos por Honorarios (KH, C0), cargando la OC y extrayendo los asientos.", styles['BulletItem']))

    # Sección 2
    story.append(Paragraph("2. Requisitos Previos en la Laptop", styles['SecHeader']))
    story.append(Paragraph("1. Tener <b>SAP GUI for Windows</b> abierto con sesión iniciada en tu usuario habitual.", styles['BulletItem']))
    story.append(Paragraph("2. Habilitar Scripting en SAP Logon (<code>Alt + F12</code> → Accesibilidad y Scripting → Scripting → ☑ Habilitar scripting).", styles['BulletItem']))
    story.append(Paragraph("3. Asegurar que el archivo <code>Masivo Reporte.xlsx</code> esté en la misma carpeta que el proyecto y que <b>esté cerrado en Microsoft Excel</b> para que el macro pueda guardar los resultados.", styles['BulletItem']))

    # Sección 3
    story.append(Paragraph("3. Instrucciones de Ejecución Paso a Paso", styles['SecHeader']))
    story.append(Paragraph("1. Inicia sesión en tu mandante de SAP habitual.", styles['BulletItem']))
    story.append(Paragraph("2. Entra a la carpeta del proyecto: <code>Proyectos/Calera/scripts/</code>.", styles['BulletItem']))
    story.append(Paragraph("3. Haz <b>doble clic</b> sobre el archivo:", styles['BulletItem']))
    story.append(Paragraph("<b><code>2_sap_miro_macro.vbs</code></b>", styles['CodeBlock']))
    story.append(Paragraph("4. Verás cómo en tu pantalla de SAP se abre la transacción <code>/nMIRO</code> y se procesan automáticamente cada una de las facturas y recibos:", styles['BodyDark']))
    story.append(Paragraph("   - Escribe fecha, referencia, importe total y moneda PEN.", styles['BulletItem']))
    story.append(Paragraph("   - Selecciona indicador de impuesto C1 o C0 y marca Calc.Impuestos.", styles['BulletItem']))
    story.append(Paragraph("   - Coloca la OC, trae las posiciones y valida el saldo a cero.", styles['BulletItem']))
    story.append(Paragraph("   - Cambia a clase de documento KR o KH en la pestaña Detalle.", styles['BulletItem']))
    story.append(Paragraph("   - Simula y contabiliza el comprobante.", styles['BulletItem']))
    story.append(Paragraph("5. Al finalizar todo el lote, Windows mostrará una ventana emergente (MsgBox) informando el total de comprobantes contabilizados con éxito.", styles['BulletItem']))

    story.append(Spacer(1, 4))
    story.append(crear_caja_aviso("Si deseas ejecutarlo por consola de comandos para ver logs en vivo, puedes abrir CMD y escribir: cscript scripts/2_sap_miro_macro.vbs", styles))
    story.append(Spacer(1, 6))

    # Sección 4
    story.append(Paragraph("4. Manejo de Errores Comunes", styles['SecHeader']))
    story.append(Paragraph("• <b>Mensaje 'No se encontro el archivo Masivo Reporte.xlsx'</b>: Verifica que el archivo Excel esté guardado en la carpeta raíz del proyecto con ese nombre exacto.", styles['BulletItem']))
    story.append(Paragraph("• <b>Mensaje 'No se pudo conectar con SAP GUI'</b>: Asegúrate de tener SAP abierto e iniciada la sesión antes del doble clic.", styles['BulletItem']))
    story.append(Paragraph("• <b>Error de permiso en Excel</b>: Cierra Microsoft Excel antes de ejecutar el macro para que no bloquee el archivo.", styles['BulletItem']))

    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"✅ Generado: {ruta_salida}")


# ==============================================================================
# 3. GUÍA OPCIÓN 3: RPA POR TECLADO Y PANTALLA (PLAN B)
# ==============================================================================
def generar_pdf_opcion_3(ruta_salida):
    styles = get_custom_styles()
    doc = SimpleDocTemplate(
        ruta_salida,
        pagesize=letter,
        leftMargin=40, rightMargin=40,
        topMargin=45, bottomMargin=45
    )
    story = []

    # Encabezado
    story.append(Paragraph("LA CALERA · MANUAL OFICIAL DE AUTOMATIZACIÓN", styles['DocSubtitle']))
    story.append(Paragraph("Guía 3: RPA Asistido por Teclado y Pantalla (Plan B) ⌨️", styles['DocTitle']))
    story.append(Paragraph("<b>Plan de Contingencia ante Bloqueos de TI:</b> Solución basada en simulación robótica de teclado y atajos oficiales para cuando el área de TI tiene bloqueado el Scripting de SAP en el servidor.", styles['DocSubtitle']))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1b4d24"), spaceAfter=10))

    # Sección 1
    story.append(Paragraph("1. ¿Cuándo utilizar el RPA por Teclado?", styles['SecHeader']))
    story.append(Paragraph(
        "En algunas organizaciones, las políticas de seguridad de TI deshabilitan el parámetro de servidor "
        "<code>sapgui/user_scripting = TRUE</code>, lo cual impide que cualquier script COM se conecte a SAP GUI. "
        "Para este escenario de emergencia se desarrolló la <b>Opción 3 (RPA de Teclado)</b>.", styles['BodyDark']
    ))
    story.append(Paragraph("• <b>Emula 100% a un operador humano</b>: Digita en el teclado mediante eventos nativos del sistema operativo.", styles['BulletItem']))
    story.append(Paragraph("• <b>No requiere permisos de scripting en SAP</b>: Para el servidor de SAP, es indistinguible de una persona física digitando.", styles['BulletItem']))
    story.append(Paragraph("• <b>Utiliza atajos oficiales de SAP</b>: Navega entre campos con <code>TAB</code>, <code>Ctrl + /</code> para la barra de comandos, <code>F6</code> para referencias y <code>Ctrl + S</code> para contabilizar.", styles['BulletItem']))

    # Sección 2
    story.append(Paragraph("2. Requisitos e Instalación", styles['SecHeader']))
    story.append(Paragraph("1. Python 3.10 o superior instalado.", styles['BulletItem']))
    story.append(Paragraph("2. Instalar las librerías de control de interfaz de usuario:", styles['BulletItem']))
    story.append(Paragraph("<code>pip install pyautogui pywinauto openpyxl</code>", styles['CodeBlock']))

    # Sección 3
    story.append(Paragraph("3. Reglas de Oro durante la Automatización", styles['SecHeader']))
    story.append(Paragraph("Dado que este robot controla el teclado y el foco de la pantalla, es <b>indispensable</b> seguir estas 3 reglas:", styles['BodyDark']))
    story.append(Paragraph("1. <b>Mantener la ventana de SAP maximizada y en primer plano</b>.", styles['BulletItem']))
    story.append(Paragraph("2. <b>NO tocar el mouse ni presionar ninguna tecla</b> mientras el robot esté digitando para evitar alterar el foco de los campos.", styles['BulletItem']))
    story.append(Paragraph("3. <b>Mecanismo de Failsafe (Parada de Emergencia)</b>: Si en cualquier momento necesitas detener el robot inmediatamente, <b>mueve el puntero del mouse con fuerza hacia la esquina superior izquierda de tu pantalla (coordenada 0, 0)</b>. El robot se interrumpirá de inmediato por seguridad.", styles['BulletItem']))

    # Sección 4
    story.append(Paragraph("4. Cómo Ejecutar el Robot", styles['SecHeader']))
    story.append(Paragraph("1. Abre tu SAP GUI y colócalo maximizado en pantalla.", styles['BulletItem']))
    story.append(Paragraph("2. Abre tu terminal y ejecuta:", styles['BulletItem']))
    story.append(Paragraph("<code>python scripts/3_sap_miro_rpa_teclado.py</code>", styles['CodeBlock']))
    story.append(Paragraph("3. El script mostrará una cuenta regresiva de 5 segundos: <i>'⏳ El robot iniciará en 5 segundos... (coloca SAP en pantalla)'</i>.", styles['BulletItem']))
    story.append(Paragraph("4. Haz clic en la ventana de SAP y suelta el mouse y teclado.", styles['BulletItem']))
    story.append(Paragraph("5. El robot navegará campo por campo, colocará la OC, traerá las posiciones y contabilizará con <code>Ctrl + S</code>.", styles['BulletItem']))

    story.append(Spacer(1, 4))
    story.append(crear_caja_aviso("Esta opción es un plan de contingencia. Siempre que el scripting esté habilitado, la Opción 1 (Python Scripting) es preferible por su mayor velocidad y tolerancia a fallos.", styles))

    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"✅ Generado: {ruta_salida}")


def main():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    docs_dir = os.path.join(base_dir, "docs")
    os.makedirs(docs_dir, exist_ok=True)
    p1 = os.path.join(docs_dir, "Manual_1_Bot_Python_SAP_MIRO.pdf")
    p2 = os.path.join(docs_dir, "Manual_2_Macro_VBScript_SAP_MIRO.pdf")
    p3 = os.path.join(docs_dir, "Manual_3_Bot_RPA_Teclado_SAP_MIRO.pdf")

    print("📄 Generando las 3 Guías PDF oficiales para La Calera en docs/...")
    generar_pdf_opcion_1(p1)
    generar_pdf_opcion_2(p2)
    generar_pdf_opcion_3(p3)
    print("🎉 ¡Las 3 guías PDF fueron generadas exitosamente!")

if __name__ == '__main__':
    main()
