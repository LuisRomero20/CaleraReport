# Guía técnica del maestro de control y renombrado MIRO

## Arquitectura

- `index.html` y `app.js`: interfaz estática, carga acumulativa, conciliación,
  vista previa y generación de los dos formatos Excel.
- `vendor/xlsx.bundle.js` y `vendor/jszip.min.js`: dependencias locales para
  leer, escribir y terminar los XLSX con fila superior inmovilizada, sin CDN.
- `scripts/control_ordenes.py`: motor Python y CLI del maestro A–W.
- `scripts/renombrar_documentos_miro.py`: interfaz Tkinter y motor reversible de
  renombrado.
- `scripts/procesar_masivo.py`: generador histórico del masivo SAP A–U.
- `tests/test_control_ordenes.py`: integración con los ejemplos reales y pruebas
  destructivas únicamente sobre copias temporales.

## Identidad documental

La clave de grupo usa semana, provincia normalizada, RUC, tipo, comprobante, OC
y HES. La clave de fila añade motivo e importe para conservar casos con varias
líneas de servicio dentro del mismo documento.

Los datos manuales se recuperan por la clave de grupo. Si el mismo grupo contiene
MIRO diferentes, el sistema no elige uno: genera una incidencia bloqueante.

## Precedencia

1. Coincidencia documental exacta en `BD`.
2. Consolidado de la provincia y semana seleccionadas.
3. `ME5A` para SOLPED, OC, texto breve y liberación.
4. `ME2N` para liberación de la orden y texto de servicio.
5. `ML84` para la relación OC–HES.
6. Nombres PDF para confirmar la existencia física.

No se usa el nombre del proveedor como clave suficiente. Las comparaciones de
provincia ignoran acentos, separadores y mayúsculas, pero el resultado conserva
la denominación del negocio.

## Incidencias

Una incidencia `BLOQUEANTE` impide marcar una fila como cerrada. Ejemplos:

- Falta SOLPED, OC, HES, comprobante, RUC o provincia.
- ME5A relaciona la SOLPED con otra OC.
- ML84 relaciona la OC con otra HES.
- Falta un PDF obligatorio o existen varios candidatos.
- El expediente no fue cargado en modo maestro completo.
- Un maestro editado contiene MIRO contradictorios.

Una `ADVERTENCIA` conserva el registro y deja trazabilidad, por ejemplo cuando
el texto de negocio difiere entre el consolidado y `BD`.

## Seguridad del renombrado

El análisis y la aplicación son fases independientes. Solo se aplican candidatos
con estado `LISTO`. Cada ruta se resuelve dentro de la carpeta raíz autorizada.
No se siguen enlaces simbólicos y nunca se sobrescribe un destino existente.

El historial conserva ruta anterior, ruta nueva, fecha, alcance, tamaño y hash
SHA-256. La reversión exige que el contenido no haya cambiado desde la operación.
Una ejecución repetida detecta el prefijo existente y devuelve `YA RENOMBRADO`.

## Diferencia entre los Excel

| Archivo | Columna U | Uso |
|---|---|---|
| Control de órdenes de servicio A–W | MIRO manual | Seguimiento y organización documental. |
| Masivo SAP A–U | N° OC | Entrada para los bots MIRO existentes. |

El motor de control no ejecuta contabilizaciones SAP.

## Validación en Windows

```powershell
pip install -r scripts/requirements.txt
python -m unittest discover -s tests -v
python scripts/control_ordenes.py expedientes/Semana_41 --semana 41 --resumen-json
python scripts/renombrar_documentos_miro.py
```

Para validar la web sin publicar:

```powershell
python -m http.server 8000
```

Después abre `http://localhost:8000` en Chrome o Edge.
