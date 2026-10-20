#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Herramienta local, segura y reversible para prefijar PDF con su MIRO.

La interfaz nunca cambia archivos durante el análisis. El renombrado requiere
una confirmación explícita y registra hashes/rutas para poder revertirlo.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Sequence

try:
    from control_ordenes import (
        build_pdf_index,
        clean_id,
        document_group_key,
        normalize_document,
        normalize_province,
        province_display,
        read_bd_records,
    )
except ImportError:  # Permite importar como scripts.renombrar_documentos_miro
    from scripts.control_ordenes import (  # type: ignore
        build_pdf_index,
        clean_id,
        document_group_key,
        normalize_document,
        normalize_province,
        province_display,
        read_bd_records,
    )


MIRO_PATTERN = re.compile(r"^\d{10}$")
EXISTING_PREFIX_PATTERN = re.compile(r"^(\d{10})\s+-\s+", re.IGNORECASE)
HISTORY_DIRECTORY = ".calera_renombrado"
HISTORY_FILE = "historial_renombrado.json"


@dataclass
class RenameCandidate:
    province: str
    hes: str
    document: str
    miro: str
    current_path: str
    new_path: str
    kind: str
    status: str
    reason: str = ""


def _safe_resolve(path: os.PathLike[str] | str, root: Path) -> Path:
    resolved = Path(path).resolve(strict=False)
    try:
        resolved.relative_to(root)
    except ValueError as exc:
        raise ValueError(f"La ruta queda fuera de la carpeta autorizada: {resolved}") from exc
    return resolved


def _iter_safe_pdfs(root: Path) -> Iterable[Path]:
    root = root.resolve(strict=True)
    for current, dirnames, filenames in os.walk(root, followlinks=False):
        current_path = Path(current).resolve(strict=True)
        _safe_resolve(current_path, root)
        dirnames[:] = [
            name
            for name in dirnames
            if not (Path(current, name).is_symlink() or name == HISTORY_DIRECTORY)
        ]
        for filename in filenames:
            path = Path(current, filename)
            if path.suffix.lower() != ".pdf" or path.is_symlink():
                continue
            yield _safe_resolve(path, root)


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _record_file_matches(record: dict[str, Any], index: dict[str, list[dict[str, str]]], scope: str) -> list[tuple[str, Path]]:
    province_files = index.get(normalize_province(record.get("provincia")), [])
    document = normalize_document(record.get("documento"))
    identifiers = {
        "HES": re.sub(r"\D", "", clean_id(record.get("hes"))),
        "OC": re.sub(r"\D", "", clean_id(record.get("oc"))),
        "SOLPED": re.sub(r"\D", "", clean_id(record.get("solped"))),
    }
    allowed = {"HES"} if scope == "hes" else {"HES", "OC", "SOLPED", "RXH", "FACT"}
    matches: list[tuple[str, Path]] = []
    for info in province_files:
        kind = info.get("type", "OTRO")
        if kind not in allowed:
            continue
        if kind in {"RXH", "FACT"}:
            matched = bool(document and info.get("document") == document)
        else:
            matched = bool(identifiers.get(kind) and info.get("number") == identifiers[kind])
        if matched:
            matches.append((kind, Path(info["path"])))
    return matches


def analyze_renames(
    master_path: os.PathLike[str] | str,
    expedients_root: os.PathLike[str] | str,
    *,
    scope: str = "hes",
) -> list[RenameCandidate]:
    """Crea una vista previa. No modifica archivos."""

    if scope not in {"hes", "todos"}:
        raise ValueError("El alcance debe ser 'hes' o 'todos'.")
    master = Path(master_path).resolve(strict=True)
    root = Path(expedients_root).resolve(strict=True)
    if not root.is_dir():
        raise ValueError("La ruta de expedientes debe ser una carpeta.")
    records = read_bd_records(master)
    pdfs = list(_iter_safe_pdfs(root))
    pdf_index = build_pdf_index(pdfs)
    candidates: list[RenameCandidate] = []

    seen_groups: set[tuple[str, ...]] = set()
    for record in records:
        miro = clean_id(record.get("miro"))
        if not miro:
            continue
        base = {
            "province": province_display(record.get("provincia")),
            "hes": clean_id(record.get("hes")),
            "document": clean_id(record.get("documento")),
            "miro": miro,
        }
        if not MIRO_PATTERN.fullmatch(miro):
            candidates.append(
                RenameCandidate(
                    **base,
                    current_path="",
                    new_path="",
                    kind="",
                    status="MIRO INVÁLIDO",
                    reason="El MIRO debe tener exactamente 10 dígitos.",
                )
            )
            continue

        group_key = document_group_key(record)
        if group_key in seen_groups:
            continue
        seen_groups.add(group_key)
        matches = _record_file_matches(record, pdf_index, scope)
        required_kinds = {"HES"} if scope == "hes" else {"HES", "OC", normalize_document(record.get("documento")) and record.get("tipo_doc", "RXH")}
        if not matches:
            candidates.append(
                RenameCandidate(
                    **base,
                    current_path="",
                    new_path="",
                    kind="HES" if scope == "hes" else "",
                    status="NO ENCONTRADO",
                    reason="No se encontró un PDF asociado por provincia e identificadores documentales.",
                )
            )
            continue
        for kind, current in matches:
            current = _safe_resolve(current, root)
            prefix = EXISTING_PREFIX_PATTERN.match(current.name)
            if prefix:
                if prefix.group(1) == miro:
                    candidates.append(
                        RenameCandidate(
                            **base,
                            current_path=str(current),
                            new_path=str(current),
                            kind=kind,
                            status="YA RENOMBRADO",
                            reason="El archivo ya tiene el prefijo MIRO correcto.",
                        )
                    )
                else:
                    candidates.append(
                        RenameCandidate(
                            **base,
                            current_path=str(current),
                            new_path="",
                            kind=kind,
                            status="CONFLICTO",
                            reason=f"El archivo ya tiene un MIRO diferente ({prefix.group(1)}).",
                        )
                    )
                continue
            destination = _safe_resolve(current.with_name(f"{miro} - {current.name}"), root)
            if destination.exists() and destination != current:
                status = "COLISIÓN"
                reason = "Ya existe un archivo con el nombre de destino."
            else:
                status = "LISTO"
                reason = ""
            candidates.append(
                RenameCandidate(
                    **base,
                    current_path=str(current),
                    new_path=str(destination),
                    kind=kind,
                    status=status,
                    reason=reason,
                )
            )

    # Un archivo asociado a MIRO diferentes queda bloqueado.
    by_source: dict[str, list[RenameCandidate]] = {}
    for candidate in candidates:
        if candidate.current_path:
            by_source.setdefault(candidate.current_path.casefold(), []).append(candidate)
    for group in by_source.values():
        miros = {item.miro for item in group}
        destinations = {item.new_path.casefold() for item in group if item.new_path}
        if len(miros) > 1 or len(destinations) > 1:
            for item in group:
                item.status = "CONFLICTO"
                item.reason = "El mismo PDF quedó asociado a más de un MIRO."

    # Un mismo archivo propuesto varias veces con el mismo MIRO se muestra una vez.
    deduplicated: list[RenameCandidate] = []
    seen: set[tuple[str, str]] = set()
    for candidate in candidates:
        key = (candidate.current_path.casefold(), candidate.new_path.casefold())
        if candidate.current_path and key in seen:
            continue
        seen.add(key)
        deduplicated.append(candidate)
    return deduplicated


def _history_path(root: Path) -> Path:
    return root / HISTORY_DIRECTORY / HISTORY_FILE


def _read_history(root: Path) -> dict[str, Any]:
    path = _history_path(root)
    if not path.exists():
        return {"version": 1, "operations": []}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"No se pudo leer el historial de renombrado: {exc}") from exc
    if not isinstance(data, dict) or not isinstance(data.get("operations"), list):
        raise ValueError("El historial de renombrado tiene un formato inválido.")
    return data


def _write_history(root: Path, history: dict[str, Any]) -> None:
    directory = root / HISTORY_DIRECTORY
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / HISTORY_FILE
    temporary = directory / f"{HISTORY_FILE}.tmp"
    temporary.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, path)


def apply_renames(
    candidates: Sequence[RenameCandidate],
    expedients_root: os.PathLike[str] | str,
    *,
    master_path: os.PathLike[str] | str = "",
    scope: str = "hes",
) -> dict[str, Any]:
    """Aplica exclusivamente candidatos LISTO y registra una operación reversible."""

    root = Path(expedients_root).resolve(strict=True)
    ready = [item for item in candidates if item.status == "LISTO"]
    operation = {
        "id": datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S.%fZ"),
        "created_at": datetime.now(timezone.utc).isoformat(),
        "root": str(root),
        "master": str(Path(master_path).resolve(strict=False)) if master_path else "",
        "scope": scope,
        "reverted_at": None,
        "entries": [],
    }
    for candidate in ready:
        old_path = _safe_resolve(candidate.current_path, root)
        new_path = _safe_resolve(candidate.new_path, root)
        entry = {
            "old_path": str(old_path),
            "new_path": str(new_path),
            "miro": candidate.miro,
            "kind": candidate.kind,
            "status": "ERROR",
            "error": "",
        }
        try:
            if not old_path.exists() or not old_path.is_file():
                raise FileNotFoundError("El archivo original ya no existe.")
            if old_path.is_symlink():
                raise ValueError("No se permiten enlaces simbólicos.")
            if new_path.exists():
                raise FileExistsError("El nombre de destino ya existe.")
            prefix = EXISTING_PREFIX_PATTERN.match(old_path.name)
            if prefix:
                raise ValueError("El archivo ya tiene un prefijo MIRO.")
            before = old_path.stat()
            digest = _hash_file(old_path)
            old_path.rename(new_path)
            after = new_path.stat()
            entry.update(
                {
                    "status": "RENOMBRADO",
                    "sha256": digest,
                    "size": after.st_size,
                    "source_mtime_ns": before.st_mtime_ns,
                    "renamed_mtime_ns": after.st_mtime_ns,
                }
            )
        except Exception as exc:  # Registrar cada error sin eliminar documentos.
            entry["error"] = str(exc)
        operation["entries"].append(entry)
    history = _read_history(root)
    history["operations"].append(operation)
    _write_history(root, history)
    return operation


def undo_last_rename(expedients_root: os.PathLike[str] | str) -> dict[str, Any]:
    root = Path(expedients_root).resolve(strict=True)
    history = _read_history(root)
    operation = next(
        (
            item
            for item in reversed(history["operations"])
            if not item.get("reverted_at") and any(entry.get("status") == "RENOMBRADO" for entry in item.get("entries", []))
        ),
        None,
    )
    if operation is None:
        raise ValueError("No existe una operación pendiente de reversión.")
    errors: list[str] = []
    reverted = 0
    for entry in reversed(operation["entries"]):
        if entry.get("status") != "RENOMBRADO":
            continue
        old_path = _safe_resolve(entry["old_path"], root)
        new_path = _safe_resolve(entry["new_path"], root)
        try:
            if old_path.exists():
                raise FileExistsError(f"No se puede restaurar porque ya existe: {old_path}")
            if not new_path.exists() or not new_path.is_file():
                raise FileNotFoundError(f"No existe el archivo renombrado: {new_path}")
            if new_path.is_symlink():
                raise ValueError("No se permiten enlaces simbólicos.")
            if new_path.stat().st_size != entry.get("size") or _hash_file(new_path) != entry.get("sha256"):
                raise ValueError(f"El archivo cambió después del renombrado: {new_path}")
            new_path.rename(old_path)
            entry["status"] = "REVERTIDO"
            reverted += 1
        except Exception as exc:
            errors.append(str(exc))
    if not errors:
        operation["reverted_at"] = datetime.now(timezone.utc).isoformat()
    operation["undo_errors"] = errors
    _write_history(root, history)
    return {"operation_id": operation["id"], "reverted": reverted, "errors": errors}


class RenameApp:
    def __init__(self) -> None:
        import tkinter as tk
        from tkinter import ttk

        self.tk = tk
        self.root = tk.Tk()
        self.root.title("CALERA - Renombrado documental por MIRO")
        self.root.geometry("1220x680")
        self.master_var = tk.StringVar()
        self.folder_var = tk.StringVar()
        self.scope_var = tk.StringVar(value="hes")
        self.status_var = tk.StringVar(value="Selecciona el maestro editado y la carpeta de expedientes.")
        self.candidates: list[RenameCandidate] = []

        frame = ttk.Frame(self.root, padding=14)
        frame.pack(fill="both", expand=True)
        controls = ttk.Frame(frame)
        controls.pack(fill="x")
        ttk.Button(controls, text="Seleccionar Excel", command=self.select_master).grid(row=0, column=0, padx=4, pady=4)
        ttk.Entry(controls, textvariable=self.master_var, width=80, state="readonly").grid(row=0, column=1, padx=4, pady=4, sticky="ew")
        ttk.Button(controls, text="Seleccionar expedientes", command=self.select_folder).grid(row=1, column=0, padx=4, pady=4)
        ttk.Entry(controls, textvariable=self.folder_var, width=80, state="readonly").grid(row=1, column=1, padx=4, pady=4, sticky="ew")
        controls.columnconfigure(1, weight=1)

        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=(8, 10))
        ttk.Radiobutton(actions, text="Solo PDF de HES", variable=self.scope_var, value="hes").pack(side="left", padx=4)
        ttk.Radiobutton(actions, text="Todos los PDF relacionados", variable=self.scope_var, value="todos").pack(side="left", padx=4)
        ttk.Button(actions, text="Analizar MIRO", command=self.analyze).pack(side="left", padx=10)
        ttk.Button(actions, text="Vista previa", command=self.show_preview).pack(side="left", padx=4)
        ttk.Button(actions, text="Renombrar documentos", command=self.rename).pack(side="left", padx=4)
        ttk.Button(actions, text="Deshacer último renombrado", command=self.undo).pack(side="left", padx=4)

        columns = ("province", "hes", "document", "miro", "current", "new", "status")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", height=22)
        labels = {
            "province": "Provincia",
            "hes": "HES",
            "document": "Documento",
            "miro": "MIRO",
            "current": "Archivo actual",
            "new": "Nombre nuevo",
            "status": "Estado",
        }
        widths = {"province": 110, "hes": 110, "document": 110, "miro": 110, "current": 270, "new": 310, "status": 120}
        for column in columns:
            self.tree.heading(column, text=labels[column])
            self.tree.column(column, width=widths[column], anchor="w")
        yscroll = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.pack(fill="both", expand=True, side="left")
        yscroll.pack(fill="y", side="right")
        xscroll.pack(fill="x", side="bottom")
        ttk.Label(self.root, textvariable=self.status_var, padding=(14, 6)).pack(fill="x")

    def select_master(self) -> None:
        from tkinter import filedialog

        selected = filedialog.askopenfilename(title="Seleccionar maestro con MIRO", filetypes=[("Excel", "*.xlsx")])
        if selected:
            self.master_var.set(selected)

    def select_folder(self) -> None:
        from tkinter import filedialog

        selected = filedialog.askdirectory(title="Seleccionar carpeta raíz de expedientes")
        if selected:
            self.folder_var.set(selected)

    def analyze(self) -> None:
        from tkinter import messagebox

        try:
            self.candidates = analyze_renames(
                self.master_var.get(),
                self.folder_var.get(),
                scope=self.scope_var.get(),
            )
            self.show_preview()
        except Exception as exc:
            messagebox.showerror("No se pudo analizar", str(exc))

    def show_preview(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        for candidate in self.candidates:
            current_name = Path(candidate.current_path).name if candidate.current_path else ""
            new_name = Path(candidate.new_path).name if candidate.new_path else ""
            self.tree.insert(
                "",
                "end",
                values=(candidate.province, candidate.hes, candidate.document, candidate.miro, current_name, new_name, candidate.status),
            )
        counts: dict[str, int] = {}
        for candidate in self.candidates:
            counts[candidate.status] = counts.get(candidate.status, 0) + 1
        self.status_var.set(" | ".join(f"{status}: {count}" for status, count in sorted(counts.items())) or "Sin registros con MIRO.")

    def rename(self) -> None:
        from tkinter import messagebox

        ready = sum(item.status == "LISTO" for item in self.candidates)
        if not ready:
            messagebox.showinfo("Sin cambios", "No hay archivos en estado LISTO.")
            return
        if not messagebox.askyesno("Confirmar renombrado", f"Se renombrarán {ready} PDF. ¿Deseas continuar?"):
            return
        operation = apply_renames(
            self.candidates,
            self.folder_var.get(),
            master_path=self.master_var.get(),
            scope=self.scope_var.get(),
        )
        renamed = sum(entry.get("status") == "RENOMBRADO" for entry in operation["entries"])
        errors = sum(entry.get("status") == "ERROR" for entry in operation["entries"])
        messagebox.showinfo("Resultado", f"Renombrados: {renamed}\nErrores: {errors}")
        self.analyze()

    def undo(self) -> None:
        from tkinter import messagebox

        if not messagebox.askyesno("Confirmar reversión", "Se intentará restaurar el último renombrado. ¿Continuar?"):
            return
        try:
            result = undo_last_rename(self.folder_var.get())
            messagebox.showinfo("Reversión", f"Restaurados: {result['reverted']}\nErrores: {len(result['errors'])}")
            if self.master_var.get() and self.folder_var.get():
                self.analyze()
        except Exception as exc:
            messagebox.showerror("No se pudo revertir", str(exc))

    def run(self) -> None:
        self.root.mainloop()


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Vista previa y renombrado seguro de PDF por MIRO.")
    parser.add_argument("--excel", help="Maestro Excel con MIRO completado.")
    parser.add_argument("--expedientes", help="Carpeta raíz de expedientes.")
    parser.add_argument("--alcance", choices=("hes", "todos"), default="hes")
    parser.add_argument("--analizar", action="store_true", help="Imprime la vista previa en JSON sin cambiar archivos.")
    parser.add_argument("--aplicar", action="store_true", help="Aplica los candidatos LISTO.")
    parser.add_argument("--confirmar", action="store_true", help="Confirmación adicional obligatoria para --aplicar.")
    parser.add_argument("--deshacer", action="store_true", help="Revierte la última operación válida.")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    if not any((args.analizar, args.aplicar, args.deshacer)):
        RenameApp().run()
        return 0
    if not args.expedientes:
        raise SystemExit("--expedientes es obligatorio.")
    if args.deshacer:
        print(json.dumps(undo_last_rename(args.expedientes), ensure_ascii=False, indent=2))
        return 0
    if not args.excel:
        raise SystemExit("--excel es obligatorio para analizar o aplicar.")
    candidates = analyze_renames(args.excel, args.expedientes, scope=args.alcance)
    if args.aplicar:
        if not args.confirmar:
            raise SystemExit("Por seguridad, --aplicar también requiere --confirmar.")
        result = apply_renames(candidates, args.expedientes, master_path=args.excel, scope=args.alcance)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(json.dumps([asdict(item) for item in candidates], ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
