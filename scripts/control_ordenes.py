#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Motor del maestro de control de órdenes de servicio de CALERA.

Este módulo no contabiliza documentos en SAP. Consolida expedientes, valida las
relaciones SOLPED/OC/HES, preserva datos manuales de un maestro anterior y crea
el libro de control A-W. El antiguo ``Masivo Reporte.xlsx`` sigue siendo un
entregable independiente.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import unicodedata
from collections import Counter, defaultdict
from copy import copy
from dataclasses import asdict, dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable, Iterator, Sequence

import openpyxl
from pypdf import PdfReader
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule, IconSetRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo


MASTER_HEADERS = [
    "ID ",
    "FECHA",
    "SEM",
    "Provincia ",
    "SOLICITANTE",
    "MOTIVO",
    "SOLPED",
    "RUC",
    "NOMBRE DEL PROVEEDOR ",
    "TIPO DOC",
    "N° Doc",
    "VALORIZADO",
    "DETALLE SERVICIO",
    "N° OC",
    "N° HES",
    "Lib Solped",
    "¿Solped Liberada",
    "Lib OS",
    "¿OS Liberada?",
    "OBSERVACIONES",
    "MIRO",
    "OBS. MIRO",
    "ESTATUS",
]

INCIDENT_HEADERS = [
    "SEM",
    "PROVINCIA",
    "DOCUMENTO",
    "FUENTE",
    "CAMPO",
    "GRAVEDAD",
    "MOTIVO",
]

FIELD_ALIASES = {
    "id": {"ID"},
    "fecha": {"FECHA", "FECHADOCUMENTO", "FECHASOLICITUD"},
    "sem": {"SEM", "SEMANA"},
    "provincia": {"PROVINCIA", "SEDE"},
    "solicitante": {"SOLICITANTE", "AREA"},
    "motivo": {"MOTIVO", "SERVICIO"},
    "solped": {"SOLPED", "SOLICITUDDEPEDIDO", "SOLICITUDPEDIDO"},
    "ruc": {"RUC", "CODIGO"},
    "proveedor": {
        "NOMBREDELPROVEEDOR",
        "NONBREDELPROVEEDOR",
        "MONBREDELPROVEEDOR",
        "PROVEEDOR",
        "PROVEEDOR2",
    },
    "tipo_doc": {"TIPODOC", "TIPODOCUMENTO"},
    "documento": {"NDOC", "RXHN", "DOCUMENTO", "NUMERODOCUMENTO"},
    "valorizado": {"VALORIZADO", "VALOR", "VALORTOTAL", "PRECIODEVALORACION"},
    "detalle": {"DETALLESERVICIO", "TEXTOBREVE", "DETALLE"},
    "oc": {"NOC", "PEDIDO", "ORDENDECOMPRA", "OS"},
    "hes": {"NHES", "HOJADEENTRADA", "HE"},
    "lib_solped": {"LIBSOLPED", "INDICADORLIBERACION"},
    "solped_liberada": {"SOLPEDLIBERADA"},
    "lib_os": {"LIBOS", "INDLIBERACION"},
    "os_liberada": {"OSLIBERADA"},
    "observaciones": {"OBSERVACIONES"},
    "miro": {"MIRO", "NMIRO"},
    "obs_miro": {"OBSMIRO"},
    "estatus": {"ESTATUS", "STATUS"},
    "centro": {"CENTRO"},
}

REQUIRED_CONTROL_FIELDS = (
    "fecha",
    "sem",
    "provincia",
    "ruc",
    "tipo_doc",
    "documento",
    "valorizado",
    "oc",
    "hes",
)

BLOCKING = "BLOQUEANTE"
WARNING = "ADVERTENCIA"
INFO = "INFORMATIVA"


@dataclass(frozen=True)
class Incident:
    week: str
    province: str
    document: str
    source: str
    field: str
    severity: str
    reason: str


@dataclass
class ConsolidationResult:
    week: str
    records: list[dict[str, Any]]
    incidents: list[Incident]
    selected_files: int
    selected_sites: list[str]
    reference_master: str | None = None

    def to_summary(self) -> dict[str, Any]:
        by_province = Counter(str(row.get("provincia") or "") for row in self.records)
        total_val = sum(Decimal(str(r["valorizado"])) for r in self.records if r.get("valorizado") is not None)
        supported_count = sum(bool(r.get("_physical_supported")) for r in self.records)
        reconciliation_pct = round((supported_count / max(1, len(self.records))) * 100, 2)
        return {
            "week": self.week,
            "records": len(self.records),
            "sites": dict(sorted(by_province.items())),
            "selected_files": self.selected_files,
            "selected_sites": self.selected_sites,
            "incidents": len(self.incidents),
            "blocking_incidents": sum(i.severity == BLOCKING for i in self.incidents),
            "warning_incidents": sum(i.severity == WARNING for i in self.incidents),
            "supported_records": supported_count,
            "closed_records": sum(r.get("estatus") == "CERRADO" for r in self.records),
            "observed_records": sum(r.get("estatus") == "OBSERVADO" for r in self.records),
            "pending_miro": sum(r.get("estatus") == "PENDIENTE MIRO" for r in self.records),
            "total_valorizado": float(total_val),
            "reconciliation_rate": reconciliation_pct,
        }


def normalize_text(value: Any) -> str:
    """Normaliza texto para comparar sin alterar el valor mostrado."""

    text = str(value or "").replace("�", "")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", text).strip().upper()


def normalize_header(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]+", "", normalize_text(value))


def clean_id(value: Any) -> str:
    """Convierte identificadores de Excel a texto sin notación decimal."""

    if value is None:
        return ""
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip().replace("\u00a0", "")
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


def clean_number(value: Any) -> float | None:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        return float(value)
    text = str(value).upper().replace("S/", "").replace("PEN", "")
    text = text.replace(",", "").strip()
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    try:
        return float(Decimal(match.group(0)))
    except InvalidOperation:
        return None


def normalize_document(value: Any) -> str:
    text = normalize_text(value)
    text = re.sub(r"^0[12]-0", "", text)
    text = re.sub(r"\b(?:RXH|RHE|FACTURA|FACT)\b", "", text)
    text = re.sub(r"[^A-Z0-9]+", "", text)
    match = re.fullmatch(r"([A-Z]\d{3})(\d+)", text)
    return f"{match.group(1)}{int(match.group(2))}" if match else text


def display_document(value: Any) -> str:
    key = normalize_document(value)
    match = re.fullmatch(r"([A-Z]\d{3})(\d+)", key)
    return f"{match.group(1)}-{int(match.group(2))}" if match else clean_id(value).strip()


def format_document_reference(value: Any, doc_type: Any) -> str:
    key = normalize_document(value)
    match = re.fullmatch(r"([A-Z]\d{3})(\d+)", key)
    if not match:
        return clean_id(value).strip()
    prefix = "01-0" if normalize_doc_type(doc_type) == "FACT" else "02-0"
    return f"{prefix}{match.group(1)}-{int(match.group(2)):07d}"


def normalize_doc_type(value: Any, ruc: Any = "") -> str:
    text = normalize_text(value)
    if "FACT" in text or text in {"01", "FAC"}:
        return "FACT"
    if "RXH" in text or text in {"RH", "RHE", "02"} or "HONOR" in text:
        return "RXH"
    return "FACT" if clean_id(ruc).startswith("20") else "RXH"


def normalize_province(value: Any) -> str:
    text = normalize_text(value)
    text = re.sub(r"^\d{3,5}[\s_-]*", "", text)
    return re.sub(r"[^A-Z0-9]+", "", text)


def province_display(value: Any) -> str:
    text = normalize_text(value)
    text = re.sub(r"^\d{3,5}[\s_-]*", "", text)
    return re.sub(r"[_-]+", " ", text).strip()


def week_display(value: Any) -> str:
    text = clean_id(value)
    match = re.search(r"\d+", text)
    return str(int(match.group(0))) if match else text


def parse_excel_date(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if isinstance(value, (int, float)):
        try:
            return openpyxl.utils.datetime.from_excel(value)
        except Exception:
            return None
    text = str(value or "").strip()
    for fmt in ("%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def find_field(header_value: Any) -> str | None:
    normalized = normalize_header(header_value)
    for field, aliases in FIELD_ALIASES.items():
        if normalized in aliases:
            return field
    return None


def detect_header_row(ws: Any, max_rows: int = 12) -> tuple[int, dict[str, int]]:
    best: tuple[int, int, dict[str, int]] | None = None
    for row_number in range(1, min(ws.max_row, max_rows) + 1):
        mapping: dict[str, int] = {}
        for column in range(1, ws.max_column + 1):
            field = find_field(ws.cell(row_number, column).value)
            if field and field not in mapping:
                mapping[field] = column
        score = len(mapping)
        if best is None or score > best[0]:
            best = (score, row_number, mapping)
    if not best or best[0] < 4:
        raise ValueError(f"No se reconocieron encabezados en la hoja {ws.title!r}")
    return best[1], best[2]


def value_from(row: Sequence[Any], columns: dict[str, int], field: str) -> Any:
    column = columns.get(field)
    if not column or column > len(row):
        return None
    return row[column - 1]


def infer_week(path: Path) -> str:
    for part in reversed(path.parts):
        match = re.search(r"SEM(?:ANA)?[\s_-]*(\d+)", normalize_text(part))
        if match:
            return str(int(match.group(1)))
    return ""


def infer_site(path: Path) -> tuple[str, str]:
    for part in reversed(path.parts):
        match = re.match(r"^(\d{3,5})[\s_-]+(.+)$", part.strip())
        if match:
            return match.group(1), province_display(match.group(2))
    return "", ""


def row_record(row: Sequence[Any], columns: dict[str, int], source: str) -> dict[str, Any]:
    ruc = clean_id(value_from(row, columns, "ruc"))
    record = {
        "id": clean_id(value_from(row, columns, "id")),
        "fecha": parse_excel_date(value_from(row, columns, "fecha")),
        "sem": week_display(value_from(row, columns, "sem")),
        "provincia": province_display(value_from(row, columns, "provincia")),
        "solicitante": clean_id(value_from(row, columns, "solicitante")).strip(),
        "motivo": clean_id(value_from(row, columns, "motivo")).strip(),
        "solped": clean_id(value_from(row, columns, "solped")),
        "ruc": ruc,
        "proveedor": clean_id(value_from(row, columns, "proveedor")).strip(),
        "tipo_doc": normalize_doc_type(value_from(row, columns, "tipo_doc"), ruc),
        "documento": display_document(value_from(row, columns, "documento")),
        "valorizado": clean_number(value_from(row, columns, "valorizado")),
        "detalle": clean_id(value_from(row, columns, "detalle")).strip(),
        "oc": clean_id(value_from(row, columns, "oc")),
        "hes": clean_id(value_from(row, columns, "hes")),
        "lib_solped": clean_id(value_from(row, columns, "lib_solped")),
        "solped_liberada": clean_number(value_from(row, columns, "solped_liberada")),
        "lib_os": clean_id(value_from(row, columns, "lib_os")),
        "os_liberada": clean_number(value_from(row, columns, "os_liberada")),
        "observaciones": clean_id(value_from(row, columns, "observaciones")).strip(),
        "miro": clean_id(value_from(row, columns, "miro")),
        "obs_miro": clean_id(value_from(row, columns, "obs_miro")).strip(),
        "estatus": clean_id(value_from(row, columns, "estatus")).strip(),
        "_source": source,
        "_physical_supported": False,
        "_evidence": {},
        "_blocking": [],
        "_warnings": [],
    }
    return record


def document_group_key(record: dict[str, Any]) -> tuple[str, ...]:
    return (
        week_display(record.get("sem")),
        normalize_province(record.get("provincia")),
        re.sub(r"\D", "", clean_id(record.get("ruc"))),
        normalize_doc_type(record.get("tipo_doc"), record.get("ruc")),
        normalize_document(record.get("documento")),
        re.sub(r"\D", "", clean_id(record.get("oc"))),
        re.sub(r"\D", "", clean_id(record.get("hes"))),
    )


def record_key(record: dict[str, Any]) -> tuple[str, ...]:
    amount = record.get("valorizado")
    amount_key = "" if amount is None else f"{float(amount):.2f}"
    return document_group_key(record) + (normalize_header(record.get("motivo")), amount_key)


def _iter_data_rows(ws: Any, header_row: int) -> Iterator[Sequence[Any]]:
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if any(value not in (None, "") for value in row):
            yield row


def read_bd_records(path: os.PathLike[str] | str) -> list[dict[str, Any]]:
    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if "BD" not in workbook.sheetnames:
        raise ValueError(f"El archivo {path} no contiene una hoja BD")
    ws = workbook["BD"]
    header_row, columns = detect_header_row(ws)
    records = []
    for row in _iter_data_rows(ws, header_row):
        record = row_record(row, columns, f"{Path(path).name}:BD")
        if record["fecha"] or record["documento"] or record["solped"]:
            records.append(record)
    workbook.close()
    return records


def _read_sheet_index(path: Path, sheet_name: str) -> tuple[list[dict[str, Any]], dict[str, int]]:
    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if sheet_name not in workbook.sheetnames:
        workbook.close()
        return [], {}
    ws = workbook[sheet_name]
    header_row, columns = detect_header_row(ws)
    values = []
    for row in _iter_data_rows(ws, header_row):
        values.append({field: value_from(row, columns, field) for field in columns})
    workbook.close()
    return values, columns


def read_reference_sources(path: os.PathLike[str] | str) -> dict[str, Any]:
    master_path = Path(path)
    result: dict[str, Any] = {
        "records": read_bd_records(master_path),
        "me5a": {},
        "me2n": {},
        "ml84": defaultdict(set),
    }

    me5a, _ = _read_sheet_index(master_path, "ME5A")
    for row in me5a:
        solped = clean_id(row.get("solped"))
        if not solped:
            continue
        result["me5a"].setdefault(solped, []).append(
            {
                "oc": clean_id(row.get("oc")),
                "lib_solped": clean_id(row.get("lib_solped")),
                "detalle": clean_id(row.get("detalle")).strip(),
                "solicitante": clean_id(row.get("solicitante")).strip(),
                "centro": clean_id(row.get("centro")),
            }
        )

    me2n, _ = _read_sheet_index(master_path, "ME2N")
    for row in me2n:
        oc = clean_id(row.get("oc"))
        if not re.fullmatch(r"\d{8,12}", oc):
            continue
        result["me2n"].setdefault(oc, []).append(
            {
                "lib_os": clean_id(row.get("lib_os")),
                "detalle": clean_id(row.get("detalle")).strip(),
                "valorizado": clean_number(row.get("valorizado")),
                "centro": clean_id(row.get("centro")),
            }
        )

    # ML84 tiene bloques irregulares. Las columnas M/N contienen pares OC/HES
    # válidos cuando M empieza por 43 y N por 10.
    workbook = openpyxl.load_workbook(master_path, data_only=True, read_only=True)
    if "ML84" in workbook.sheetnames:
        ws = workbook["ML84"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) < 14:
                continue
            oc = clean_id(row[12])
            hes = clean_id(row[13])
            if re.fullmatch(r"43\d{7,10}", oc) and re.fullmatch(r"10\d{8,10}", hes):
                result["ml84"][oc].add(hes)
    workbook.close()
    result["ml84"] = dict(result["ml84"])
    return result


def parse_consolidated(
    path: os.PathLike[str] | str,
    expected_week: str = "",
    expected_province: str = "",
) -> tuple[list[dict[str, Any]], list[Incident]]:
    file_path = Path(path)
    workbook = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
    sheet_name = next(
        (name for name in workbook.sheetnames if "CONSOLIDADO" in normalize_text(name)),
        workbook.sheetnames[0],
    )
    ws = workbook[sheet_name]
    header_row, columns = detect_header_row(ws)
    records: list[dict[str, Any]] = []
    incidents: list[Incident] = []
    normalized_expected_province = normalize_province(expected_province)
    for row in _iter_data_rows(ws, header_row):
        record = row_record(row, columns, f"{file_path.name}:{sheet_name}")
        if not (record["fecha"] or record["documento"] or record["solped"]):
            continue
        if expected_week and record["sem"] != week_display(expected_week):
            continue
        if normalized_expected_province:
            row_province = normalize_province(record["provincia"])
            if row_province != normalized_expected_province:
                continue
        record["_source_path"] = str(file_path)
        record["_site_dir"] = str(file_path.parent)
        records.append(record)
    workbook.close()
    if not records:
        incidents.append(
            Incident(
                week_display(expected_week),
                province_display(expected_province),
                "",
                file_path.name,
                "REGISTROS",
                BLOCKING,
                "El consolidado no contiene filas para la semana y provincia seleccionadas.",
            )
        )
    return records, incidents


def collect_input_files(roots: Iterable[os.PathLike[str] | str]) -> list[Path]:
    files: dict[str, Path] = {}
    for raw_root in roots:
        root = Path(raw_root).expanduser()
        if root.is_file():
            candidates = [root]
        elif root.is_dir():
            candidates = [p for p in root.rglob("*") if p.is_file()]
        else:
            continue
        for path in candidates:
            if path.name.startswith("~$"):
                continue
            if path.suffix.lower() not in {".xlsx", ".xls", ".pdf"}:
                continue
            key = str(path.resolve()).casefold()
            files[key] = path.resolve()
    return sorted(files.values(), key=lambda item: str(item).casefold())


def locate_reference_master(files: Iterable[Path]) -> Path | None:
    candidates = [
        path
        for path in files
        if path.suffix.lower() in {".xlsx", ".xls"}
        and "CONTROL" in normalize_text(path.name)
        and "ORDEN" in normalize_text(path.name)
        and "SERVICIO" in normalize_text(path.name)
    ]
    return sorted(candidates, key=lambda item: len(str(item)))[0] if candidates else None


MONTHS_MAP = {
    "enero": "01",
    "febrero": "02",
    "marzo": "03",
    "abril": "04",
    "mayo": "05",
    "junio": "06",
    "julio": "07",
    "agosto": "08",
    "setiembre": "09",
    "septiembre": "09",
    "octubre": "10",
    "noviembre": "11",
    "diciembre": "12",
}


def _filename_info(path: Path) -> dict[str, Any]:
    name = path.stem
    normalized = normalize_text(name)
    prefix_match = re.match(r"^(\d{10})\s*-\s*", normalized)
    base = normalized[prefix_match.end() :] if prefix_match else normalized
    site_code, province = infer_site(path.parent)
    week = infer_week(path)
    result: dict[str, Any] = {
        "path": str(path),
        "miro_prefix": prefix_match.group(1) if prefix_match else "",
        "site_code": site_code,
        "province": province,
        "week": week,
        "type": "OTRO",
        "number": "",
        "document": "",
        "documents": [],
        "document_type": "",
        "provider_hint": "",
    }
    for label, pattern in (
        ("HES", r"\bHES[\s_-]*(\d{8,12})"),
        ("SOLPED", r"\bSOLPED[\s_-]*(\d{7,12})"),
        ("OC", r"\bOC[\s_-]*(\d{8,12})"),
    ):
        match = re.search(pattern, base)
        if match:
            result["type"] = label
            result["number"] = match.group(1)
            break
    doc_matches = list(re.finditer(r"\b(RXH|FACTURA|FACT)[\s_-]*([A-Z]\d{3})[\s_-]*(\d+)", base))
    if doc_matches:
        docs = []
        for match in doc_matches:
            doc_code = f"{match.group(2)}{int(match.group(3))}"
            if doc_code not in docs:
                docs.append(doc_code)
        result["documents"] = docs
        result["document"] = docs[0]
        result["document_type"] = "FACT" if "FACT" in doc_matches[0].group(1) else "RXH"
        if result["type"] == "OTRO":
            result["type"] = result["document_type"]
    segments = re.split(r"\s+-\s+", base)
    if len(segments) > 1:
        result["provider_hint"] = segments[-1].strip()
    return result


def _pdf_document_total(path: Path, doc_type: str) -> float | None:
    text = " ".join(page.extract_text() or "" for page in PdfReader(str(path)).pages)
    text = re.sub(r"\s+", " ", normalize_text(text))
    patterns = (
        (r"IMPORTE\s+TOTAL\s*:?\s*(?:S/?\s*)?([0-9,]+\.\d{2})",)
        if doc_type == "FACT"
        else (
            r"TOTAL\s+NETO\s+RECIBIDO[\s\S]{0,180}?([0-9,]+\.\d{2})\s+SOLES",
            r"TOTAL\s+NETO\s+RECIBIDO\s*:?\s*(?:SOLES\s*)?([0-9,]+\.\d{2})",
            r"TOTAL\s+POR\s+HONORARIOS\s*:?\s*(?:S/?\s*)?([0-9,]+\.\d{2})",
            r"\b20\d{2}\s+([0-9,]+\.\d{2})\b",
        )
    )
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return float(match.group(1).replace(",", ""))
    return None


def _pdf_date(text: str) -> str | None:
    match = re.search(
        r"FECHA\s+DE\s+EMISI[OÓ]N\s*:?\s*(\d{1,2}[./-]\d{1,2}[./-]\d{4})",
        text,
        re.IGNORECASE,
    ) or re.search(
        r"N[UÚ]MERO\s*/?\s*FECHA\s+\d{8,12}\s*/\s*(\d{1,2}[./-]\d{1,2}[./-]\d{4})",
        text,
        re.IGNORECASE,
    )
    if match:
        return re.sub(r"[.-]", "/", match.group(1))
    match = re.search(
        r"(\d{1,2})\s+(?:DE\s+)?([A-ZÁÉÍÓÚÑ]+)\s+(?:DEL?\s+)?(\d{4})",
        text,
        re.IGNORECASE,
    )
    if match:
        month = MONTHS_MAP.get(normalize_text(match.group(2)).lower())
        if month:
            return f"{int(match.group(1)):02d}/{month}/{match.group(3)}"
    return None


def _pdf_delivery_date(text: str) -> str | None:
    match = re.search(
        r"\b\d{4}\s*-\s*CD-[A-ZÁÉÍÓÚÑ_-]+\s+(\d{1,2}[./-]\d{1,2}[./-]\d{4})",
        text,
        re.IGNORECASE,
    )
    return re.sub(r"[.-]", "/", match.group(1)) if match else None


def _service_motive(text: str) -> str:
    value = normalize_text(text)
    if re.search(r"MONTACARGA|MAQUINARIA PESADA", value):
        return "SERV. OP. MONTACARGA"
    if "LIMPIEZA" in value:
        return "SERV. LIMPIEZA"
    if re.search(r"ESTIBA|DESESTIBA", value):
        return "SERV. ESTIBA"
    if re.search(r"SANEO|RESANEO|SELECCION", value):
        return "SERV. SANEO"
    compact = re.sub(r"_{3,}[\s\S]*", "", value).strip()
    return compact[:80].strip()


def _clean_provider_name(raw_provider: str, word_hints: Sequence[str]) -> str:
    provider = re.sub(r"\s+", " ", str(raw_provider or "")).strip()
    if not provider:
        return ""
    words = sorted(
        {
            word
            for hint in word_hints
            for word in re.split(r"[^A-Z0-9]+", normalize_text(hint))
            if len(word) >= 3
        },
        key=len,
        reverse=True,
    )
    for word in words:
        escaped = re.escape(word)
        provider = re.sub(rf"([A-Z]{{2,}})({escaped})\b", r"\1 \2", provider)
        provider = re.sub(rf"\b({escaped})([A-Z]{{2,}})", r"\1 \2", provider)
    return re.sub(r"\s+", " ", provider).strip()


def _parse_pdf_info(path: Path) -> dict[str, Any]:
    base = _filename_info(path)
    try:
        raw_text = " ".join(page.extract_text() or "" for page in PdfReader(str(path)).pages)
        clean = re.sub(r"[\x00-\x1f\ufffd\xa0]", " ", raw_text)
        clean = re.sub(r"\s+", " ", clean).strip()
        upper = normalize_text(clean)
        is_fact = base["document_type"] == "FACT" or base["type"] == "FACT" or "FACTURA ELECTRONICA" in upper
        is_rxh = base["document_type"] == "RXH" or base["type"] == "RXH" or "RECIBO POR HONORARIOS" in upper
        kind = base["type"] if base["type"] in {"HES", "OC", "SOLPED"} else ("FACT" if is_fact else "RXH" if is_rxh else base["type"])
        total = _pdf_document_total(path, "FACT" if is_fact else "RXH") if (is_fact or is_rxh) else None
        service_total = None
        service_match = re.search(r"VALOR\s+TOTAL\s+SIN\s+IVA\s+PEN\s+([0-9,]+\.\d{2})", upper)
        if service_match:
            service_total = float(service_match.group(1).replace(",", ""))
        elif kind == "OC":
            net_match = re.search(r"\bNETO\s+DESCUENTO\s+SUB-TOTAL\s+I\.G\.V\.\s+TOTAL\s+([0-9,]+\.\d{2})", upper)
            if net_match:
                service_total = float(net_match.group(1).replace(",", ""))
        ruc = ""
        provider = ""
        hes_person_hint = ""
        vendor = re.search(r"PROVEEDOR\s*:\s*(.+?)\s+RUC\s*:\s*(\d{11})", upper)
        if vendor:
            provider = vendor.group(1).strip()
            ruc = vendor.group(2)
        if not ruc and is_fact:
            m = re.search(r"FACTURA\s+ELECTRONICA\s+RUC\s*:\s*(\d{11})", upper)
            if m:
                ruc = m.group(1)
        if not ruc and is_rxh:
            m = re.search(r"NRO\s*:\s*(\d{11})\s*-\s*[A-Z]\d{3}", upper)
            if m:
                ruc = m.group(1)
        hes_person = re.search(
            r"VALOR\s+TOTAL\s+SIN\s+IVA\s+PEN\s+[0-9,]+\.\d{2}\s+(.{3,65}?)\s+(?:JR\b|CALLE\b|AV\b|URB\b|AH\b|AA\s*HH\b|ASOC\b|ASC\b|FUNDO\b|GRAN\b|LAS\b|P\.?J\.?\b|MZA?\b|LOTE\b|LT\b|\d{5})",
            upper,
        )
        if hes_person:
            hes_person_hint = hes_person.group(1).strip()
        oc_match = re.search(r"POSICION\s+PEDIDO\s*/?\s*FECHA\s+(43\d{7,10})", upper) or re.search(
            r"ORDEN\s+DE\s+SERVICIO\s+N\s*(43\d{7,10})", upper
        )
        hes_match = re.search(r"N[UÚ]MERO\s*/?\s*FECHA\s+(10\d{8,10})", upper)
        week_match = re.search(r"\bSEM(?:ANA)?\s*[-:]?\s*(\d{1,2})\b", upper)
        detail = ""
        detail_match = re.search(r"\b\d{7}\s+(.{3,120}?)(?:L.{0,8}PLANIFICACI|LINEA\s+PLANIFICACION)", upper)
        if detail_match:
            detail = detail_match.group(1).strip()
        lead = re.split(r"_{5,}", clean)[0].strip()
        return {
            **base,
            "text": clean,
            "kind": kind,
            "total": total,
            "service_total": service_total,
            "date": _pdf_date(clean),
            "delivery_date": _pdf_delivery_date(clean),
            "week": str(int(week_match.group(1))) if week_match else base["week"],
            "ruc": ruc,
            "provider": provider,
            "hes_person_hint": hes_person_hint,
            "oc": oc_match.group(1) if oc_match else "",
            "hes": hes_match.group(1) if hes_match else "",
            "solicitante": "ALMACEN" if "ALM_LIMA_PRO" in upper else "",
            "motivo": _service_motive(lead) if base["type"] == "HES" else _service_motive(detail),
            "detalle": detail or lead,
            "readable": bool(clean),
        }
    except Exception:
        return {
            **base,
            "text": "",
            "kind": base["type"],
            "total": None,
            "service_total": None,
            "date": None,
            "delivery_date": None,
            "ruc": "",
            "provider": "",
            "hes_person_hint": "",
            "oc": "",
            "hes": "",
            "solicitante": "",
            "motivo": "",
            "detalle": "",
            "readable": False,
        }


def build_pdf_index(files: Iterable[Path]) -> dict[str, list[dict[str, Any]]]:
    index: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for path in files:
        if path.suffix.lower() != ".pdf":
            continue
        info = _parse_pdf_info(path)
        index[normalize_province(info["province"])].append(info)
    return dict(index)


def _infer_service_week(
    document_info: dict[str, Any] | None,
    oc_info: dict[str, Any] | None,
    hes_info: dict[str, Any] | None,
) -> str:
    if document_info and document_info.get("date"):
        doc_date = parse_excel_date(document_info["date"])
        if doc_date:
            iso = doc_date.isocalendar()[1]
            return str(iso if doc_date.weekday() == 6 else max(1, iso - 1))
    sap_date_str = (
        (oc_info and (oc_info.get("delivery_date") or oc_info.get("date")))
        or (hes_info and hes_info.get("date"))
        or None
    )
    if sap_date_str:
        sap_date = parse_excel_date(sap_date_str)
        if sap_date:
            return str(max(1, sap_date.isocalendar()[1] - 1))
    return ""


def _includes_document(info: dict[str, Any], document: str) -> bool:
    return bool(document) and (document in (info.get("documents") or []) or info.get("document") == document)


def _unique_value(values: Iterable[str]) -> str:
    unique = list(dict.fromkeys(v for v in values if v))
    return unique[0] if len(unique) == 1 else ""


def _records_from_pdf_index(
    pdf_index: dict[str, list[dict[str, Any]]],
    incidents: list[Incident],
) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for province_key, parsed_files in pdf_index.items():
        if not province_key:
            continue
        hes_files = [info for info in parsed_files if info.get("type") == "HES" and info.get("number")]
        for hes_info in hes_files:
            document = hes_info.get("document") or ""
            document_type = hes_info.get("document_type") or normalize_doc_type("", hes_info.get("ruc"))
            by_document = [info for info in parsed_files if _includes_document(info, document)]
            oc_candidates = [
                info
                for info in parsed_files
                if info.get("type") == "OC"
                and (
                    (hes_info.get("oc") and info.get("number") == hes_info.get("oc"))
                    or (not hes_info.get("oc") and _includes_document(info, document))
                )
            ]
            oc_number = hes_info.get("oc") or _unique_value(info.get("number", "") for info in oc_candidates)
            oc_info = next((info for info in oc_candidates if info.get("number") == oc_number), oc_candidates[0] if oc_candidates else None)
            solped_candidates = [info for info in by_document if info.get("type") == "SOLPED"]
            solped = _unique_value(info.get("number", "") for info in solped_candidates)
            solped_info = next((info for info in solped_candidates if info.get("number") == solped), None)
            doc_candidates = [info for info in by_document if info.get("type") in {"FACT", "RXH"}]
            doc_info = doc_candidates[0] if doc_candidates else None
            source_infos = [info for info in (doc_info, oc_info, hes_info, solped_info) if info]
            date_str = (
                (doc_info and doc_info.get("date"))
                or (oc_info and (oc_info.get("delivery_date") or oc_info.get("date")))
                or (hes_info and hes_info.get("date"))
                or (solped_info and solped_info.get("date"))
                or None
            )
            week = next((info.get("week") for info in source_infos if info.get("week")), "") or hes_info.get("week") or ""
            if not week:
                week = _infer_service_week(doc_info, oc_info, hes_info)
            raw_provider = next((info.get("provider") for info in source_infos if info.get("provider")), "") or hes_info.get("provider_hint") or ""
            provider = _clean_provider_name(
                raw_provider,
                [
                    hes_info.get("provider_hint", ""),
                    hes_info.get("hes_person_hint", ""),
                    oc_info.get("provider_hint", "") if oc_info else "",
                    doc_info.get("provider_hint", "") if doc_info else "",
                ],
            )
            valorizado = (
                doc_info["total"]
                if (doc_info and doc_info.get("total") is not None)
                else (
                    hes_info.get("service_total")
                    if hes_info.get("service_total") is not None
                    else (oc_info.get("service_total") if oc_info else None)
                )
            )
            has_solped_or_oc = bool(solped or oc_number)
            has_oc_and_hes = bool(oc_number and hes_info.get("number"))
            file_miro = next((info.get("miro_prefix") for info in source_infos if info.get("miro_prefix")), "")
            record: dict[str, Any] = {
                "id": "",
                "fecha": parse_excel_date(date_str),
                "sem": week,
                "provincia": hes_info.get("province", ""),
                "solicitante": next((info.get("solicitante") for info in source_infos if info.get("solicitante")), "") or "ALMACEN",
                "motivo": next((info.get("motivo") for info in source_infos if info.get("motivo")), ""),
                "solped": solped,
                "ruc": next((info.get("ruc") for info in source_infos if info.get("ruc")), ""),
                "proveedor": provider,
                "tipo_doc": document_type,
                "documento": display_document(document),
                "valorizado": valorizado,
                "detalle": next((info.get("detalle") for info in source_infos if info.get("detalle")), ""),
                "oc": oc_number,
                "hes": hes_info.get("number", ""),
                "lib_solped": "2" if has_solped_or_oc else "",
                "solped_liberada": 2 if has_solped_or_oc else 0,
                "lib_os": "G" if has_oc_and_hes else "",
                "os_liberada": 2 if has_oc_and_hes else 0,
                "observaciones": "",
                "miro": file_miro,
                "obs_miro": "",
                "estatus": "",
                "_source": hes_info["path"],
                "_source_path": hes_info["path"],
                "_site_dir": str(Path(hes_info["path"]).parent),
                "_site_code": hes_info.get("site_code", ""),
                "_physical_supported": False,
                "_evidence": {},
                "_blocking": [],
                "_warnings": [],
            }
            if not hes_info.get("readable"):
                _add_incident(
                    incidents,
                    record,
                    hes_info["path"],
                    "LECTURA PDF",
                    WARNING,
                    "El PDF HES no permitió extraer texto; se conservaron los datos disponibles en el nombre del archivo.",
                )
            if len(oc_candidates) > 1:
                _add_incident(incidents, record, "Expediente PDF", "N° OC", BLOCKING, "Se encontraron varios PDF de OC candidatos para el mismo expediente.")
            if len(solped_candidates) > 1:
                _add_incident(incidents, record, "Expediente PDF", "SOLPED", BLOCKING, "Se encontraron varios PDF de SOLPED candidatos para el mismo documento.")
            if len(doc_candidates) > 1:
                _add_incident(incidents, record, "Expediente PDF", "DOCUMENTO", BLOCKING, "Se encontraron varios comprobantes PDF para el mismo documento.")
            records.append(record)
    return records


def _add_incident(
    incidents: list[Incident],
    record: dict[str, Any],
    source: str,
    field: str,
    severity: str,
    reason: str,
) -> None:
    incident = Incident(
        week_display(record.get("sem")),
        province_display(record.get("provincia")),
        display_document(record.get("documento")),
        source,
        field,
        severity,
        reason,
    )
    if incident not in incidents:
        incidents.append(incident)
    target = "_blocking" if severity == BLOCKING else "_warnings"
    if reason not in record[target]:
        record[target].append(reason)


def _physical_evidence(record: dict[str, Any], pdf_index: dict[str, list[dict[str, Any]]]) -> dict[str, list[str]]:
    province_files = pdf_index.get(normalize_province(record.get("provincia")), [])
    doc = normalize_document(record.get("documento"))
    oc = re.sub(r"\D", "", clean_id(record.get("oc")))
    hes = re.sub(r"\D", "", clean_id(record.get("hes")))
    solped = re.sub(r"\D", "", clean_id(record.get("solped")))
    evidence: dict[str, list[str]] = {"documento": [], "oc": [], "hes": [], "solped": []}
    for info in province_files:
        if doc and info.get("document") == doc and info.get("type") in {"RXH", "FACT"}:
            evidence["documento"].append(info["path"])
        if oc and info.get("type") == "OC" and info.get("number") == oc:
            evidence["oc"].append(info["path"])
        if hes and info.get("type") == "HES" and info.get("number") == hes:
            evidence["hes"].append(info["path"])
        if solped and info.get("type") == "SOLPED" and info.get("number") == solped:
            evidence["solped"].append(info["path"])
    return evidence


def _validate_required(record: dict[str, Any], incidents: list[Incident]) -> None:
    for field in REQUIRED_CONTROL_FIELDS:
        if record.get(field) in (None, ""):
            _add_incident(
                incidents,
                record,
                record.get("_source", "Expediente PDF"),
                field.upper(),
                BLOCKING,
                f"Falta el campo esencial {field.upper()}.",
            )


def _validate_physical(record: dict[str, Any], pdf_index: dict[str, list[dict[str, Any]]], incidents: list[Incident]) -> None:
    evidence = _physical_evidence(record, pdf_index)
    record["_evidence"] = evidence
    if len(evidence["documento"]) == 1:
        document_path = evidence["documento"][0]
        document = next(
            info
            for info in pdf_index.get(normalize_province(record.get("provincia")), [])
            if info["path"] == document_path
        )
        if document.get("total") is not None:
            record["valorizado"] = document["total"]
    for field in ("oc", "hes"):
        if not evidence[field]:
            _add_incident(
                incidents,
                record,
                "Expediente PDF",
                field.upper(),
                BLOCKING,
                f"No se encontró el PDF de {field.upper()} en la carpeta de la provincia.",
            )
        elif len(evidence[field]) > 1:
            _add_incident(
                incidents,
                record,
                "Expediente PDF",
                field.upper(),
                BLOCKING,
                f"Se encontraron varios PDF candidatos para {field.upper()}.",
            )
    if len(evidence["documento"]) > 1:
        _add_incident(
            incidents,
            record,
            "Expediente PDF",
            "DOCUMENTO",
            BLOCKING,
            "Se encontraron varios PDF candidatos para DOCUMENTO.",
        )
    elif len(evidence["documento"]) == 0 and (not record.get("documento") or record.get("valorizado") is None):
        _add_incident(
            incidents,
            record,
            "Expediente PDF",
            "DOCUMENTO",
            BLOCKING,
            "No se encontró el PDF de DOCUMENTO ni respaldo del importe en la carpeta de la provincia.",
        )
    record["_physical_supported"] = (
        len(evidence["oc"]) == 1
        and len(evidence["hes"]) == 1
        and (
            len(evidence["documento"]) == 1
            or (len(evidence["documento"]) == 0 and bool(record.get("documento")) and record.get("valorizado") is not None)
        )
    )


def _deduplicate(records: Sequence[dict[str, Any]], incidents: list[Incident]) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    seen: dict[tuple[str, ...], dict[str, Any]] = {}
    for record in records:
        key = record_key(record)
        if key not in seen:
            seen[key] = record
            output.append(record)
            continue
        existing = seen[key]
        if record.get("_source_path") != existing.get("_source_path"):
            _add_incident(
                incidents,
                existing,
                "Consolidación",
                "DUPLICADO",
                WARNING,
                "El mismo registro fue encontrado en más de una selección y se conservó una sola vez.",
            )
    return output


def load_manual_values(path: os.PathLike[str] | str) -> tuple[dict[tuple[str, ...], dict[str, str]], list[Incident]]:
    records = read_bd_records(path)
    by_group: dict[tuple[str, ...], list[dict[str, Any]]] = defaultdict(list)
    for record in records:
        by_group[document_group_key(record)].append(record)
    output: dict[tuple[str, ...], dict[str, str]] = {}
    incidents: list[Incident] = []
    for key, group in by_group.items():
        miros = {clean_id(item.get("miro")) for item in group if clean_id(item.get("miro"))}
        if len(miros) > 1:
            sample = group[0]
            incidents.append(
                Incident(
                    week_display(sample.get("sem")),
                    province_display(sample.get("provincia")),
                    display_document(sample.get("documento")),
                    Path(path).name,
                    "MIRO",
                    BLOCKING,
                    "El maestro editado contiene más de un MIRO para la misma clave documental.",
                )
            )
            continue
        output[key] = {
            "miro": next(iter(miros), ""),
            "obs_miro": next((clean_id(item.get("obs_miro")).strip() for item in group if clean_id(item.get("obs_miro")).strip()), ""),
            "observaciones": next((clean_id(item.get("observaciones")).strip() for item in group if clean_id(item.get("observaciones")).strip()), ""),
        }
    return output, incidents


def _apply_manual_values(
    records: Sequence[dict[str, Any]],
    manual_values: dict[tuple[str, ...], dict[str, str]],
    incidents: list[Incident],
) -> None:
    for record in records:
        values = manual_values.get(document_group_key(record))
        if not values:
            continue
        miro = values.get("miro", "")
        if miro and not re.fullmatch(r"\d{10}", miro):
            _add_incident(
                incidents,
                record,
                "Maestro editado",
                "MIRO",
                BLOCKING,
                f"El valor MIRO {miro!r} no tiene 10 dígitos.",
            )
        else:
            record["miro"] = miro
        record["obs_miro"] = values.get("obs_miro", "")
        manual_observation = values.get("observaciones", "")
        if manual_observation:
            record["_manual_observation"] = manual_observation


def _finalize_record(record: dict[str, Any]) -> None:
    messages = []
    if record.get("_blocking"):
        messages.append("BLOQUEANTE: " + " | ".join(record["_blocking"]))
    if record.get("_warnings"):
        messages.append("ADVERTENCIA: " + " | ".join(record["_warnings"]))
    if record.get("_manual_observation"):
        messages.append(record["_manual_observation"])
    record["observaciones"] = " | ".join(messages)

    miro_clean = clean_id(record.get("miro"))
    is_valid_miro = bool(re.fullmatch(r"\d{10}", miro_clean))

    if record.get("_blocking"):
        record["estatus"] = "OBSERVADO"
    elif not miro_clean:
        record["estatus"] = "PENDIENTE MIRO"
    elif is_valid_miro and record.get("_physical_supported", True):
        record["estatus"] = "CERRADO"
    else:
        record["estatus"] = "OBSERVADO"


def consolidate(
    roots: Iterable[os.PathLike[str] | str],
    *,
    mode: str = "expedientes",
    week: str | int | None = None,
    reference_master: os.PathLike[str] | str | None = None,
    previous_master: os.PathLike[str] | str | None = None,
) -> ConsolidationResult:
    mode = mode.lower().strip()
    if mode not in {"expedientes", "completo"}:
        raise ValueError("El modo debe ser 'expedientes' o 'completo'.")
    all_files = collect_input_files(roots)
    pdf_files = [path for path in all_files if path.suffix.lower() == ".pdf"]
    if not pdf_files:
        raise ValueError("No se encontraron archivos PDF compatibles en las rutas seleccionadas.")
    pdf_index = build_pdf_index(pdf_files)
    incidents: list[Incident] = []
    requested_week = week_display(week) if week not in (None, "") else ""

    records = _records_from_pdf_index(pdf_index, incidents)
    if requested_week:
        records = [record for record in records if week_display(record.get("sem")) == requested_week]
    if not records:
        raise ValueError("No se encontraron expedientes PDF HES para los criterios seleccionados.")

    for record in records:
        _validate_physical(record, pdf_index, incidents)
        _validate_required(record, incidents)

    if previous_master:
        manual, manual_incidents = load_manual_values(previous_master)
        incidents.extend(manual_incidents)
        _apply_manual_values(records, manual, incidents)

    records = _deduplicate(records, incidents)
    records.sort(
        key=lambda item: (
            item.get("_site_code") or item.get("provincia") or "",
            int(item["hes"]) if str(item.get("hes", "")).isdigit() else 0,
        )
    )
    for index, record in enumerate(records, start=1):
        record["id"] = index
        _finalize_record(record)

    selected_sites = sorted({province_display(row.get("provincia")) for row in records if row.get("provincia")})
    detected_weeks = sorted({week_display(row.get("sem")) for row in records if row.get("sem")})
    result_week = requested_week or (detected_weeks[0] if len(detected_weeks) == 1 else "Todas")

    return ConsolidationResult(
        week=result_week,
        records=records,
        incidents=incidents,
        selected_files=len(pdf_files),
        selected_sites=selected_sites,
        reference_master=None,
    )


def _apply_header_style(ws: Any, row: int, columns: int, yellow_columns: set[int] | None = None) -> None:
    yellow_columns = yellow_columns or set()
    green = PatternFill("solid", fgColor="00B050")
    yellow = PatternFill("solid", fgColor="FFF200")
    white = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    black = Font(name="Calibri", size=10, bold=True, color="000000")
    border = Border(bottom=Side(style="thin", color="808080"))
    for column in range(1, columns + 1):
        cell = ws.cell(row, column)
        cell.fill = yellow if column in yellow_columns else green
        cell.font = black if column in yellow_columns else white
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border


def _master_row_values(record: dict[str, Any], row_number: int) -> list[Any]:
    blocking_formula = (
        f'=IF(OR(B{row_number}="",C{row_number}="",D{row_number}="",'
        f'H{row_number}="",J{row_number}="",K{row_number}="",L{row_number}="",N{row_number}="",O{row_number}="",'
        f'LEFT(T{row_number},10)="BLOQUEANTE"),"OBSERVADO",'
        f'IF(U{row_number}="","PENDIENTE MIRO","CERRADO"))'
    )
    return [
        record.get("id"),
        record.get("fecha"),
        int(record["sem"]) if str(record.get("sem", "")).isdigit() else record.get("sem"),
        record.get("provincia"),
        record.get("solicitante"),
        record.get("motivo"),
        record.get("solped"),
        record.get("ruc"),
        record.get("proveedor"),
        record.get("tipo_doc"),
        format_document_reference(record.get("documento"), record.get("tipo_doc")),
        record.get("valorizado"),
        record.get("detalle"),
        record.get("oc"),
        record.get("hes"),
        record.get("lib_solped"),
        f'=IFERROR(IF(VALUE(P{row_number})=2,2,0),0)',
        record.get("lib_os"),
        f'=IF(R{row_number}="G",2,IF(R{row_number}="B",0,1))',
        record.get("observaciones"),
        record.get("miro"),
        record.get("obs_miro"),
        blocking_formula,
    ]


def build_workbook(result: ConsolidationResult) -> Workbook:
    workbook = Workbook()
    ws = workbook.active
    ws.title = "BD"
    ws.sheet_view.showGridLines = False
    ws.append(MASTER_HEADERS)
    for row_number, record in enumerate(result.records, start=2):
        ws.append(_master_row_values(record, row_number))

    _apply_header_style(ws, 1, 23, {3, 5, 7, 13, 14, 15, 17, 19, 20})
    ws.row_dimensions[1].height = 30
    widths = [8, 12, 7, 15, 16, 24, 14, 14, 38, 12, 14, 16, 38, 14, 14, 14, 17, 12, 16, 46, 16, 28, 18]
    thin_gray = Side(style="thin", color="D9E1D5")
    body_border = Border(bottom=thin_gray)
    for column, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(column)].width = width
    for row in ws.iter_rows(min_row=2, max_row=max(ws.max_row, 2), max_col=23):
        for cell in row:
            cell.font = Font(name="Calibri", size=10, color="1F1F1F")
            cell.alignment = Alignment(vertical="center", wrap_text=False)
            cell.border = body_border
        row[1].number_format = "dd/mm/yyyy"
        row[6].number_format = "@"
        row[7].number_format = "@"
        row[10].number_format = "@"
        row[11].number_format = '"S/ "#,##0.00'
        row[13].number_format = "@"
        row[14].number_format = "@"
        row[20].number_format = "@"
        row[14].fill = PatternFill("solid", fgColor="F4CCCC")
        row[20].fill = PatternFill("solid", fgColor="FFF2CC")
        row[21].fill = PatternFill("solid", fgColor="FFF2CC")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:W{max(ws.max_row, 1)}"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    if ws.max_row >= 2:
        table = Table(displayName="TablaControlOS", ref=f"A1:W{ws.max_row}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium4",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(table)
        miro_validation = DataValidation(
            type="custom",
            formula1='=OR(U2="",AND(LEN(U2)=10,ISNUMBER(--U2)))',
            allow_blank=True,
        )
        miro_validation.error = "El MIRO debe tener exactamente 10 dígitos."
        miro_validation.errorTitle = "MIRO inválido"
        miro_validation.prompt = "Ingresa el número MIRO de 10 dígitos o deja la celda vacía."
        miro_validation.promptTitle = "MIRO manual"
        miro_validation.showErrorMessage = True
        miro_validation.showInputMessage = True
        ws.add_data_validation(miro_validation)
        miro_validation.add(f"U2:U{ws.max_row}")

        ws.conditional_formatting.add(
            f"Q2:Q{ws.max_row}",
            IconSetRule("3TrafficLights1", "num", [0, 1, 2], showValue=False),
        )
        ws.conditional_formatting.add(
            f"S2:S{ws.max_row}",
            IconSetRule("3TrafficLights1", "num", [0, 1, 2], showValue=False),
        )
        ws.conditional_formatting.add(
            f"W2:W{ws.max_row}",
            FormulaRule(formula=['$W2="OBSERVADO"'], fill=PatternFill("solid", fgColor="F4CCCC"), font=Font(color="9C0006", bold=True)),
        )
        ws.conditional_formatting.add(
            f"W2:W{ws.max_row}",
            FormulaRule(formula=['$W2="PENDIENTE MIRO"'], fill=PatternFill("solid", fgColor="FFF2CC"), font=Font(color="9C6500", bold=True)),
        )
        ws.conditional_formatting.add(
            f"W2:W{ws.max_row}",
            FormulaRule(formula=['$W2="CERRADO"'], fill=PatternFill("solid", fgColor="E2F0D9"), font=Font(color="375623", bold=True)),
        )

    # La plantilla de referencia oculta estas columnas auxiliares.
    for column in ("M", "P", "R"):
        ws.column_dimensions[column].hidden = True

    summary = workbook.create_sheet("RESUMEN")
    summary.sheet_view.showGridLines = False
    summary["A1"] = f"Control de órdenes de servicio - Semana {result.week}"
    summary["A1"].font = Font(name="Calibri", size=16, bold=True, color="375623")
    summary["A3"] = "Indicador"
    summary["B3"] = "Valor"
    summary_rows = [
        ("Registros", f'=COUNTA(BD!B2:B{max(ws.max_row, 2)})'),
        ("Importe total", f'=SUM(BD!L2:L{max(ws.max_row, 2)})'),
        ("Pendientes de MIRO", f'=COUNTIF(BD!W2:W{max(ws.max_row, 2)},"PENDIENTE MIRO")'),
        ("Observados", f'=COUNTIF(BD!W2:W{max(ws.max_row, 2)},"OBSERVADO")'),
        ("Cerrados", f'=COUNTIF(BD!W2:W{max(ws.max_row, 2)},"CERRADO")'),
        ("Incidencias", len(result.incidents)),
    ]
    for item in summary_rows:
        summary.append(item)
    summary["B5"].number_format = '"S/ "#,##0.00'
    province_start = 11
    province_headers = ["Provincia", "Registros", "RXH", "FACT", "Importe", "Pendiente MIRO", "Observado", "Cerrado"]
    for column, value in enumerate(province_headers, start=1):
        summary.cell(province_start, column, value)
    provinces = sorted({province_display(row.get("provincia")) for row in result.records})
    for row_number, province in enumerate(provinces, start=province_start + 1):
        summary.cell(row_number, 1, province)
        summary.cell(row_number, 2, f'=COUNTIF(BD!D:D,A{row_number})')
        summary.cell(row_number, 3, f'=COUNTIFS(BD!D:D,A{row_number},BD!J:J,"RXH")')
        summary.cell(row_number, 4, f'=COUNTIFS(BD!D:D,A{row_number},BD!J:J,"FACT")')
        summary.cell(row_number, 5, f'=SUMIF(BD!D:D,A{row_number},BD!L:L)')
        summary.cell(row_number, 6, f'=COUNTIFS(BD!D:D,A{row_number},BD!W:W,"PENDIENTE MIRO")')
        summary.cell(row_number, 7, f'=COUNTIFS(BD!D:D,A{row_number},BD!W:W,"OBSERVADO")')
        summary.cell(row_number, 8, f'=COUNTIFS(BD!D:D,A{row_number},BD!W:W,"CERRADO")')
        summary.cell(row_number, 5).number_format = '"S/ "#,##0.00'
    _apply_header_style(summary, 3, 2)
    _apply_header_style(summary, province_start, 8)
    summary.column_dimensions["A"].width = 26
    for column in range(2, 9):
        summary.column_dimensions[get_column_letter(column)].width = 18
    summary.freeze_panes = f"A{province_start + 1}"

    incident_ws = workbook.create_sheet("INCIDENCIAS")
    incident_ws.sheet_view.showGridLines = False
    incident_ws.append(INCIDENT_HEADERS)
    for incident in result.incidents:
        incident_ws.append(
            [
                incident.week,
                incident.province,
                incident.document,
                incident.source,
                incident.field,
                incident.severity,
                incident.reason,
            ]
        )
    _apply_header_style(incident_ws, 1, len(INCIDENT_HEADERS))
    incident_widths = [9, 18, 18, 28, 22, 16, 90]
    for column, width in enumerate(incident_widths, start=1):
        incident_ws.column_dimensions[get_column_letter(column)].width = width
    for row in incident_ws.iter_rows(min_row=2, max_col=7):
        row[6].alignment = Alignment(wrap_text=True, vertical="top")
        if row[5].value == BLOCKING:
            row[5].fill = PatternFill("solid", fgColor="F4CCCC")
        elif row[5].value == WARNING:
            row[5].fill = PatternFill("solid", fgColor="FFF2CC")
    incident_ws.freeze_panes = "A2"
    incident_ws.auto_filter.ref = f"A1:G{max(incident_ws.max_row, 1)}"

    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    return workbook


def save_workbook(result: ConsolidationResult, output_path: os.PathLike[str] | str) -> Path:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = build_workbook(result)
    try:
        workbook.save(output)
        return output
    except PermissionError:
        alternative = output.with_name(f"{output.stem} (Actualizado){output.suffix}")
        workbook.save(alternative)
        return alternative


def default_output_name(week: str) -> str:
    return f"Control_Ordenes_Servicio_Semana_{week}_Generado.xlsx"


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Genera el maestro de control de órdenes de servicio CALERA.")
    parser.add_argument("rutas", nargs="*", default=["expedientes"], help="Carpetas de expedientes que se consolidarán.")
    parser.add_argument("--modo", choices=("expedientes", "completo"), default="expedientes")
    parser.add_argument("--semana", help="Semana a generar cuando se seleccionan varios periodos.")
    parser.add_argument("--maestro", help="Control_Ordenes_Servicio utilizado como fuente BD/ME5A/ME2N/ML84.")
    parser.add_argument("--maestro-editado", help="Maestro previo del que se recuperarán MIRO y observaciones manuales.")
    parser.add_argument("--salida", help="Ruta del archivo XLSX generado.")
    parser.add_argument("--resumen-json", action="store_true", help="Imprime un resumen JSON verificable.")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    result = consolidate(
        args.rutas,
        mode=args.modo,
        week=args.semana,
        reference_master=args.maestro,
        previous_master=args.maestro_editado,
    )
    output = Path(args.salida) if args.salida else Path.cwd() / default_output_name(result.week)
    saved = save_workbook(result, output)
    if args.resumen_json:
        payload = result.to_summary()
        payload["output"] = str(saved.resolve())
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print(f"Maestro generado: {saved.resolve()}")
        print(f"Registros: {len(result.records)} | Incidencias: {len(result.incidents)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
