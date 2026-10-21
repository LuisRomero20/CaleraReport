#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generador de fixtures sintéticos y anónimos para CaleraReport.

Crea un conjunto completo de documentos PDF y carpetas departamentales sintéticas
(Arequipa, Chiclayo, Chimbote, Ica) que emulan la estructura operativa real pero
con razones sociales ficticias, RUCs válidos sintéticos, importes y documentos
anónimos para pruebas 100% reproducibles y seguras.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas


def create_rxh_pdf(
    path: Path,
    *,
    serie_num: str,
    proveedor: str,
    ruc: str,
    monto_neto: float,
    fecha: str,
    semana: int,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(path), pagesize=letter)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, 750, f"RECIBO POR HONORARIOS ELECTRONICO NRO: {ruc} - {serie_num}")
    c.setFont("Helvetica", 10)
    c.drawString(50, 730, f"PROVEEDOR: {proveedor} RUC: {ruc}")
    c.drawString(50, 715, f"FECHA DE EMISIÓN: {fecha}")
    c.drawString(50, 700, f"Semana: SEM-{semana}")
    c.drawString(50, 680, "Concepto: SERVICIO DE MANTENIMIENTO Y SANEO INTEGRAL")
    c.drawString(50, 660, f"Total Neto Recibido: LA CALERA 13 Octubre 2026 {monto_neto:.2f} SOLES")
    c.save()


def create_factura_pdf(
    path: Path,
    *,
    serie_num: str,
    proveedor: str,
    ruc: str,
    subtotal: float,
    total: float,
    fecha: str,
    semana: int,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(path), pagesize=letter)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, 750, f"FACTURA ELECTRONICA RUC: {ruc} {serie_num}")
    c.setFont("Helvetica", 10)
    c.drawString(50, 730, f"PROVEEDOR: {proveedor} RUC: {ruc}")
    c.drawString(50, 715, f"FECHA DE EMISIÓN: {fecha}")
    c.drawString(50, 700, f"Semana: SEMANA {semana}")
    c.drawString(50, 680, "Descripción: SERVICIOS INDUSTRIALES LOGISTICOS")
    c.drawString(50, 660, f"Valor Venta: S/ {subtotal:.2f} Importe Total: S/ {total:.2f}")
    c.save()


def create_hes_pdf(
    path: Path,
    *,
    hes_num: str,
    oc_num: str,
    proveedor: str,
    ruc: str,
    monto_sin_iva: float,
    fecha: str,
    semana: int,
    servicio_desc: str = "SERV. SANEO",
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(path), pagesize=letter)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, 750, f"HOJA DE ENTRADA DE SERVICIOS - NUMERO / FECHA {hes_num} / {fecha}")
    c.setFont("Helvetica", 10)
    c.drawString(50, 735, f"ALM_LIMA_PRO - ALMACEN GENERAL")
    c.drawString(50, 720, f"PROVEEDOR: {proveedor} RUC: {ruc}")
    c.drawString(50, 705, f"POSICION PEDIDO / FECHA {oc_num} / {fecha}")
    c.drawString(50, 690, f"SEMANA {semana} - {servicio_desc}")
    c.drawString(50, 675, f"VALOR TOTAL SIN IVA PEN {monto_sin_iva:.2f} {proveedor} AV. INDUSTRIAL 123 FUNDO LA CALERA")
    c.drawString(50, 650, f"1000001 {servicio_desc} LINEA PLANIFICACION")
    c.save()


def create_solped_pdf(
    path: Path,
    *,
    solped_num: str,
    proveedor: str,
    ruc: str,
    fecha: str,
    semana: int,
    servicio_desc: str = "SERV. SANEO",
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(path), pagesize=letter)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, 750, f"SOLICITUD DE PEDIDO N {solped_num}")
    c.setFont("Helvetica", 10)
    c.drawString(50, 735, f"PROVEEDOR: {proveedor} RUC: {ruc}")
    c.drawString(50, 720, f"FECHA: {fecha}")
    c.drawString(50, 705, f"SEMANA: {semana} - {servicio_desc}")
    c.drawString(50, 690, f"POSICION 10 - 1000001 {servicio_desc} LINEA PLANIFICACION")
    c.save()


def create_oc_pdf(
    path: Path,
    *,
    oc_num: str,
    solped_num: str,
    proveedor: str,
    ruc: str,
    total: float,
    fecha: str,
    semana: int,
    servicio_desc: str = "SERV. SANEO",
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(path), pagesize=letter)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, 750, f"ORDEN DE SERVICIO N {oc_num}")
    c.setFont("Helvetica", 10)
    c.drawString(50, 735, f"PROVEEDOR: {proveedor} RUC: {ruc}")
    c.drawString(50, 720, f"FECHA DE EMISIÓN: {fecha}")
    c.drawString(50, 705, f"SOLPED: {solped_num}")
    c.drawString(50, 690, f"SEM - {semana} - {servicio_desc}")
    c.drawString(50, 675, f"NETO DESCUENTO SUB-TOTAL I.G.V. TOTAL {total:.2f}")
    c.drawString(50, 650, f"1000001 {servicio_desc} LINEA PLANIFICACION")
    c.save()


SYNTHETIC_DATASET = [
    # CHICLAYO (4 casos - Semana 41)
    {
        "site": "1130_CHICLAYO",
        "provincia": "CHICLAYO",
        "sem": 41,
        "tipo": "RXH",
        "doc": "E001-41",
        "oc": "430062018",
        "hes": "1000094356",
        "solped": "100021001",
        "ruc": "10400000041",
        "proveedor": "PROVEEDOR MODELO CHICLAYO UNO S.A.C.",
        "monto": 436.99,
        "fecha": "13/10/2026",
        "motivo": "SERV. SANEO",
    },
    {
        "site": "1130_CHICLAYO",
        "provincia": "CHICLAYO",
        "sem": 41,
        "tipo": "RXH",
        "doc": "E001-42",
        "oc": "430062019",
        "hes": "1000094357",
        "solped": "100021002",
        "ruc": "10400000042",
        "proveedor": "PROVEEDOR MODELO CHICLAYO DOS S.A.C.",
        "monto": 520.00,
        "fecha": "13/10/2026",
        "motivo": "SERV. SANEO",
    },
    {
        "site": "1130_CHICLAYO",
        "provincia": "CHICLAYO",
        "sem": 41,
        "tipo": "RXH",
        "doc": "E001-43",
        "oc": "430062020",
        "hes": "1000094358",
        "solped": "100021003",
        "ruc": "10400000043",
        "proveedor": "PROVEEDOR MODELO CHICLAYO TRES S.A.C.",
        "monto": 310.50,
        "fecha": "13/10/2026",
        "motivo": "SERV. VIGILANCIA",
    },
    {
        "site": "1130_CHICLAYO",
        "provincia": "CHICLAYO",
        "sem": 41,
        "tipo": "FACT",
        "doc": "E001-622",
        "oc": "430062021",
        "hes": "1000094359",
        "solped": "100021004",
        "ruc": "20400000622",
        "proveedor": "CORPORACION INDUSTRIAL NORTE S.A.C.",
        "monto": 849.60,
        "subtotal": 720.00,
        "fecha": "13/10/2026",
        "motivo": "SERV. TRANSPORTE",
    },
    # AREQUIPA (4 casos - Semana 41)
    {
        "site": "1110_AREQUIPA",
        "provincia": "AREQUIPA",
        "sem": 41,
        "tipo": "RXH",
        "doc": "E001-50",
        "oc": "430060349",
        "hes": "1000091025",
        "solped": "100021005",
        "ruc": "10400000050",
        "proveedor": "SERVICIOS DEL SUR AREQUIPA S.A.C.",
        "monto": 250.00,
        "fecha": "12/10/2026",
        "motivo": "SERV. MANTENIMIENTO",
    },
    {
        "site": "1110_AREQUIPA",
        "provincia": "AREQUIPA",
        "sem": 41,
        "tipo": "RXH",
        "doc": "E001-51",
        "oc": "430060350",
        "hes": "1000091027",
        "solped": "100021006",
        "ruc": "10400000051",
        "proveedor": "SERVICIOS DEL SUR AREQUIPA S.A.C.",
        "monto": 250.00,
        "fecha": "12/10/2026",
        "motivo": "SERV. MANTENIMIENTO",
    },
    {
        "site": "1110_AREQUIPA",
        "provincia": "AREQUIPA",
        "sem": 41,
        "tipo": "RXH",
        "doc": "E001-52",
        "oc": "430060351",
        "hes": "1000091028",
        "solped": "100021007",
        "ruc": "10400000052",
        "proveedor": "SERVICIOS DEL SUR AREQUIPA S.A.C.",
        "monto": 250.00,
        "fecha": "12/10/2026",
        "motivo": "SERV. MANTENIMIENTO",
    },
    {
        "site": "1110_AREQUIPA",
        "provincia": "AREQUIPA",
        "sem": 41,
        "tipo": "RXH",
        "doc": "E001-80",
        "oc": "430060346",
        "hes": "1000090996",
        "solped": "100021008",
        "ruc": "10400000080",
        "proveedor": "CONTRATISTAS ASOCIADOS AREQUIPA S.A.C.",
        "monto": 180.00,
        "fecha": "12/10/2026",
        "motivo": "SERV. SANEO",
    },
    # CHIMBOTE (3 casos - Semana 41)
    {
        "site": "1140_CHIMBOTE",
        "provincia": "CHIMBOTE",
        "sem": 41,
        "tipo": "RXH",
        "doc": "E001-10",
        "oc": "430063001",
        "hes": "1000095001",
        "solped": "100021009",
        "ruc": "10400000010",
        "proveedor": "SERVICIOS COSTEROS CHIMBOTE S.A.C.",
        "monto": 300.00,
        "fecha": "14/10/2026",
        "motivo": "SERV. LIMPIEZA",
    },
    {
        "site": "1140_CHIMBOTE",
        "provincia": "CHIMBOTE",
        "sem": 41,
        "tipo": "RXH",
        "doc": "E001-11",
        "oc": "430063002",
        "hes": "1000095002",
        "solped": "100021010",
        "ruc": "10400000011",
        "proveedor": "SERVICIOS COSTEROS CHIMBOTE S.A.C.",
        "monto": 300.00,
        "fecha": "14/10/2026",
        "motivo": "SERV. LIMPIEZA",
    },
    {
        "site": "1140_CHIMBOTE",
        "provincia": "CHIMBOTE",
        "sem": 41,
        "tipo": "FACT",
        "doc": "E001-12",
        "oc": "430063003",
        "hes": "1000095003",
        "solped": "100021011",
        "ruc": "20400000012",
        "proveedor": "OPERACIONES PORTUARIAS S.A.C.",
        "monto": 1180.00,
        "subtotal": 1000.00,
        "fecha": "14/10/2026",
        "motivo": "SERV. ESTIBA",
    },
    # ICA (3 casos: 1 en Semana 38, 2 en Semana 41)
    {
        "site": "1180_ICA",
        "provincia": "ICA",
        "sem": 38,
        "tipo": "RXH",
        "doc": "E001-134",
        "oc": "430061212",
        "hes": "1000092668",
        "solped": "100021012",
        "ruc": "10400000134",
        "proveedor": "SERVICIOS AGROINDUSTRIALES ICA S.A.C.",
        "monto": 450.00,
        "fecha": "21/09/2026",
        "motivo": "SERV. SANEO",
    },
    {
        "site": "1180_ICA",
        "provincia": "ICA",
        "sem": 39,
        "tipo": "FACT",
        "doc": "E001-1386",
        "oc": "430061211",
        "hes": "1000092666",
        "solped": "100021013",
        "ruc": "20400001386",
        "proveedor": "INGENIERIA Y SERVICIOS ICA S.A.C.",
        "monto": 950.00,
        "subtotal": 805.08,
        "fecha": "28/09/2026",
        "motivo": "SERV. ELECTRICIDAD",
    },
    {
        "site": "1180_ICA",
        "provincia": "ICA",
        "sem": 39,
        "tipo": "FACT",
        "doc": "E001-1388",
        "oc": "430061211",
        "hes": "1000092667",
        "solped": "100021013",
        "ruc": "20400001386",
        "proveedor": "INGENIERIA Y SERVICIOS ICA S.A.C.",
        "monto": 1200.00,
        "subtotal": 1016.95,
        "fecha": "28/09/2026",
        "motivo": "SERV. ELECTRICIDAD",
    },
    # AREQUIPA (Semana 34 - 10 casos sintéticos para soporte completo)
    *(
        [
            {
                "site": "1110_AREQUIPA",
                "provincia": "AREQUIPA",
                "sem": 34,
                "tipo": "RXH",
                "doc": f"E001-{100 + i}",
                "oc": f"430060{100 + i}",
                "hes": f"1000090{100 + i}",
                "solped": f"100021{100 + i}",
                "ruc": f"10400000{100 + i}",
                "proveedor": f"CONTRATISTA SINTETICO AQP {i} S.A.C.",
                "monto": 200.00 + (i * 15.0),
                "fecha": "20/08/2026",
                "motivo": "SERV. SANEO",
            }
            for i in range(10)
        ]
    ),
]


def generar_fixtures(target_dir: Path) -> int:
    target_dir.mkdir(parents=True, exist_ok=True)
    count = 0
    ica_oc_created = False

    for item in SYNTHETIC_DATASET:
        site_dir = target_dir / item["site"]
        prov = item["proveedor"]
        doc_code = item["doc"]
        tipo = item["tipo"]
        doc_label = "FACTURA" if tipo == "FACT" else "RXH"

        # 1. Comprobante PDF (Factura o RXH)
        doc_pdf_name = f"{doc_label} {doc_code} - {prov}.pdf"
        doc_pdf_path = site_dir / doc_pdf_name
        if tipo == "FACT":
            create_factura_pdf(
                doc_pdf_path,
                serie_num=doc_code,
                proveedor=prov,
                ruc=item["ruc"],
                subtotal=item.get("subtotal", item["monto"] / 1.18),
                total=item["monto"],
                fecha=item["fecha"],
                semana=item["sem"],
            )
        else:
            create_rxh_pdf(
                doc_pdf_path,
                serie_num=doc_code,
                proveedor=prov,
                ruc=item["ruc"],
                monto_neto=item["monto"],
                fecha=item["fecha"],
                semana=item["sem"],
            )
        count += 1

        # 2. HES PDF
        hes_pdf_name = f"HES {item['hes']} - {doc_label} {doc_code} - {prov}.pdf"
        hes_pdf_path = site_dir / hes_pdf_name
        create_hes_pdf(
            hes_pdf_path,
            hes_num=item["hes"],
            oc_num=item["oc"],
            proveedor=prov,
            ruc=item["ruc"],
            monto_sin_iva=item.get("subtotal", item["monto"]),
            fecha=item["fecha"],
            semana=item["sem"],
            servicio_desc=item["motivo"],
        )
        count += 1

        # 3. OC PDF
        # Para ICA E001-1386 y E001-1388, una sola OC cubre ambos documentos
        if item["doc"] in ("E001-1386", "E001-1388"):
            if not ica_oc_created:
                combined_name = f"OC 430061211 - FACTURA E001-1386 E001-1388 - {prov}.pdf"
                create_oc_pdf(
                    site_dir / combined_name,
                    oc_num="430061211",
                    solped_num="100021013",
                    proveedor=prov,
                    ruc=item["ruc"],
                    total=2150.00,
                    fecha=item["fecha"],
                    semana=item["sem"],
                    servicio_desc=item["motivo"],
                )
                count += 1
                ica_oc_created = True
        else:
            oc_pdf_name = f"OC {item['oc']} - {doc_label} {doc_code} - {prov}.pdf"
            oc_pdf_path = site_dir / oc_pdf_name
            create_oc_pdf(
                oc_pdf_path,
                oc_num=item["oc"],
                solped_num=item["solped"],
                proveedor=prov,
                ruc=item["ruc"],
                total=item["monto"],
                fecha=item["fecha"],
                semana=item["sem"],
                servicio_desc=item["motivo"],
            )
            count += 1

        # 4. SOLPED PDF (Para Semana 34 en Arequipa)
        if item["sem"] == 34:
            solped_name = f"SOLPED {item['solped']} - {doc_label} {doc_code} - {prov}.pdf"
            create_solped_pdf(
                site_dir / solped_name,
                solped_num=item["solped"],
                proveedor=prov,
                ruc=item["ruc"],
                fecha=item["fecha"],
                semana=item["sem"],
                servicio_desc=item["motivo"],
            )
            count += 1

    return count


if __name__ == "__main__":
    root = Path(__file__).resolve().parents[1]
    fixtures_dir = root / "fixtures" / "synthetic_expedientes"
    total_created = generar_fixtures(fixtures_dir)
    print(f"✅ Fixtures sintéticos generados exitosamente en: {fixtures_dir}")
    print(f"   Total de PDFs generados: {total_created} archivos.")
