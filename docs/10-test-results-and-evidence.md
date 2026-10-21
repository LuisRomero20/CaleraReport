# CaleraReport Enterprise — Resultados de Pruebas y Evidencias (M6)

> **Documento:** `docs/10-test-results-and-evidence.md`  
> **Versión:** 1.0.0 (Enterprise Specification)  
> **Fecha:** 21 de octubre de 2026  
> **Estado:** Aprobado / Especificación Oficial  
> **Módulo:** Matriz de Aseguramiento de Calidad (QA), Resultados de Pruebas y Benchmark

---

## 1. Resumen Ejecutivo del Estado de Calidad

* **Total de Pruebas Automatizadas:** 21 pruebas unitarias y de integración.
* **Tasa de Aprobación:** **100% (21/21 exitosas)**.
* **Tiempo Total de Ejecución:** **4.06 segundos**.
* **Fixtures Utilizados:** Dataset sintético 100% reproducible y versionado ([fixtures/synthetic_expedientes/](file:///c:/Users/Luis%20Romero/Documents/Proyectos/Calera/fixtures/synthetic_expedientes)).
* **Entornos Validados:** Windows 11 (Entorno local de desarrollo) y Linux/Ubuntu (Pipeline CI GitHub Actions) sobre Python 3.11, 3.12 y 3.13.

---

## 2. Benchmark de Rendimiento de la Suite de Pruebas

| Métrica | Baseline Inicial (v3.0) | Suite Enterprise Optimizada | Mejora Relativa |
|---|:---:|:---:|:---:|
| **Tiempo Total de Ejecución** | 28.68 segundos | **4.06 segundos** | **85.8% más rápido** ⚡ |
| **Total de Métodos de Prueba** | 13 métodos | **21 métodos** | **+61.5% cobertura** |
| **Dependencia de Archivos Privados** | Requiere `expedientes/` local | **Cero dependencia** (Fixtures sintéticos) | **100% reproducible** |
| **Pruebas de Seguridad SAP Fail-Closed** | 0 pruebas | **6 pruebas con Mocks COM** | **Control Financiero Auditado** |

---

## 3. Matriz Detallada de Pruebas y Aserciones

### 3.1 Suite de Ingeniería Documental y Control ([`tests/test_control_ordenes.py`](file:///c:/Users/Luis%20Romero/Documents/Proyectos/Calera/tests/test_control_ordenes.py))

| # | Método de Prueba | Alcance y Aserciones Clave | Resultado |
|:---:|---|---|:---:|
| 1 | `test_all_expedientes_has_24_verified_records_across_4_provinces` | Consolida 24 registros en 4 provincias (`AREQUIPA: 14`, `CHICLAYO: 4`, `CHIMBOTE: 3`, `ICA: 3`). Soporte físico 100% y 0 incidencias bloqueantes. | **OK** |
| 2 | `test_week_41_three_sites_has_11_verified_records` | Filtro por semana 41 retorna exactamente 11 registros en Arequipa, Chiclayo y Chimbote. | **OK** |
| 3 | `test_real_chiclayo_relation_is_exact` | Valida relación exacta HES (`1000094356`), OC (`430062018`), Comprobante (`E001-41`), motivo e importe (`436.99`) usando razón social sintética. | **OK** |
| 4 | `test_week_filter_selects_requested_period` | Filtros por semanas 34 y 38 seleccionan los registros correctos con relación a SOLPED. | **OK** |
| 5 | `test_generated_workbook_has_required_structure_and_live_status` | Estructura de 3 hojas (`BD`, `RESUMEN`, `INCIDENCIAS`), inmovilización de paneles en `A2`, fórmulas de semáforos Q y S, y tipos de celda estrictos. | **OK** |
| 6 | `test_manual_fields_survive_regeneration_by_document_key` | Preservación de MIRO (`4300039999`) y notas contables ante regeneración por coincidencia documental. | **OK** |
| 7 | `test_invalid_short_miro_never_gives_cerrado_status` | Valida que MIROs de 3 dígitos (`"123"`) o alfanuméricos (`"ABC..."`) resulten en `OBSERVADO`, y solo un código de 10 dígitos numérico resulte en `CERRADO`. | **OK** |
| 8 | `test_to_summary_metrics_contains_full_audit_contract` | Verifica las 14 métricas del resumen de auditoría (`reconciliation_rate = 100.0%`, volumen valorizado, conteo de bloqueos). | **OK** |
| 9 | `test_old_sap_helpers_keep_their_contract` | Normalización de referencias SAP (`01-`, `02-`) y limpieza de textos de servicio. | **OK** |
| 10 | `test_sap_reference_round_trips_to_document_number` | Idas y vueltas entre referencias SAP (`02-0E001-0000047`) y documentos normalizados (`E001-47`). | **OK** |
| 11 | `test_invoice_uses_importe_total_not_valor_venta` | Facturas toman exclusivamente el Importe Total con IGV (`849.60`), no el Valor Venta (`720.00`). | **OK** |
| 12 | `test_receipt_uses_net_received_total` | Recibos por honorarios toman el Total Neto Recibido (`110.00`). | **OK** |
| 13 | `test_workbook_uses_control_formulas_and_reference_format` | Comprobación de fórmulas de semáforo `=IFERROR(IF(VALUE(P2)=2,2,0),0)`. | **OK** |
| 14 | `test_invalid_miro_is_never_ready` | La herramienta de renombrado rechaza MIROs de longitud diferente a 10 dígitos. | **OK** |
| 15 | `test_preview_apply_idempotence_and_undo_on_copies` | Ciclo completo: Análisis (`LISTO`), renombrado (`RENOMBRADO`), idempotencia (`YA RENOMBRADO`) y reversión por hash SHA-256 (`reverted = 1, errors = 0`). | **OK** |

### 3.2 Suite de Seguridad y Control Financiero SAP ([`tests/test_sap_security.py`](file:///c:/Users/Luis%20Romero/Documents/Proyectos/Calera/tests/test_sap_security.py))

| # | Método de Prueba | Alcance y Aserciones Clave | Resultado |
|:---:|---|---|:---:|
| 16 | `test_empty_confirmation_fails_closed` | Presionar Enter sin confirmación explícita cancela de inmediato la operación. | **OK** |
| 17 | `test_negative_confirmation_cancels` | Ingresar `"no"` aborta la contabilización y limpia el modal. | **OK** |
| 18 | `test_simulation_mode_never_commits_and_cancels_cleanly` | Modo `--simular` no contabiliza, retorna `"SIMULADO"` y envía `F12` para dejar SAP limpio. | **OK** |
| 19 | `test_ambiguous_sap_response_is_blocked_from_marking_success` | Si SAP emite un mensaje de advertencia o no incluye los números de factura/asiento, se reporta error y se envía a reconciliación manual; jamás `"REGISTRADO"`. | **OK** |
| 20 | `test_successful_posting_captures_both_numbers` | Extracción inequívoca de N° Factura MIRO (`5105773695`) y Asiento Contable FI (`4300035763`). | **OK** |
| 21 | `test_text_cleaning_removes_week_mention_for_sap` | Limpieza de strings como `"SERVICIO DE SANEO SEMANA 41"` hacia `"SERVICIO DE SANEO"` para BKTXT. | **OK** |

---

## 4. Evidencia Oficial de Ejecución de Pruebas

```text
PS C:\Users\Luis Romero\Documents\Proyectos\Calera> python -m unittest discover -s tests -v
test_all_expedientes_has_24_verified_records_across_4_provinces (test_control_ordenes.ControlOrdersIntegrationTests) ... ok
test_generated_workbook_has_required_structure_and_live_status (test_control_ordenes.ControlOrdersIntegrationTests) ... ok
test_invalid_short_miro_never_gives_cerrado_status (test_control_ordenes.ControlOrdersIntegrationTests) ... ok
test_manual_fields_survive_regeneration_by_document_key (test_control_ordenes.ControlOrdersIntegrationTests) ... ok
test_old_sap_helpers_keep_their_contract (test_control_ordenes.ControlOrdersIntegrationTests) ... ok
test_real_chiclayo_relation_is_exact (test_control_ordenes.ControlOrdersIntegrationTests) ... ok
test_to_summary_metrics_contains_full_audit_contract (test_control_ordenes.ControlOrdersIntegrationTests) ... ok
test_week_41_three_sites_has_11_verified_records (test_control_ordenes.ControlOrdersIntegrationTests) ... ok
test_week_filter_selects_requested_period (test_control_ordenes.ControlOrdersIntegrationTests) ... ok
test_invalid_miro_is_never_ready (test_control_ordenes.RenameToolTests) ... ok
test_preview_apply_idempotence_and_undo_on_copies (test_control_ordenes.RenameToolTests) ... ok
test_invoice_uses_importe_total_not_valor_venta (test_control_ordenes.ReportAmountTests) ... ok
test_receipt_uses_net_received_total (test_control_ordenes.ReportAmountTests) ... ok
test_sap_reference_round_trips_to_document_number (test_control_ordenes.ReportAmountTests) ... ok
test_workbook_uses_control_formulas_and_reference_format (test_control_ordenes.ReportAmountTests) ... ok
test_ambiguous_sap_response_is_blocked_from_marking_success (test_sap_security.SapSecurityFailClosedTests) ... ok
test_empty_confirmation_fails_closed (test_sap_security.SapSecurityFailClosedTests) ... ok
test_negative_confirmation_cancels (test_sap_security.SapSecurityFailClosedTests) ... ok
test_simulation_mode_never_commits_and_cancels_cleanly (test_sap_security.SapSecurityFailClosedTests) ... ok
test_successful_posting_captures_both_numbers (test_sap_security.SapSecurityFailClosedTests) ... ok
test_text_cleaning_removes_week_mention_for_sap (test_sap_security.SapSecurityFailClosedTests) ... ok

----------------------------------------------------------------------
Ran 21 tests in 4.063s

OK
```
