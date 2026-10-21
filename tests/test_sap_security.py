#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Pruebas unitarias de control y seguridad financiera para automatización SAP MIRO.

Evalúa los controles Fail-Closed, simulación controlada, prevención de falsos positivos
y validación de confirmación humana explícita mediante mocks de sesión SAP GUI.
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from importlib import import_module
sap_scripting = import_module("1_sap_miro_scripting")
registrar_fila_miro = sap_scripting.registrar_fila_miro
limpiar_texto_servicio = sap_scripting.limpiar_texto_servicio


class SapSecurityFailClosedTests(unittest.TestCase):
    def setUp(self) -> None:
        self.sleep_patcher = patch("time.sleep", return_value=None)
        self.sleep_patcher.start()

    def tearDown(self) -> None:
        self.sleep_patcher.stop()

    def _create_mock_session(self, msg_text: str = "", msg_type: str = "S") -> MagicMock:
        session = MagicMock()
        session.Children.Count = 1

        components: dict[str, MagicMock] = {}
        wnd0 = MagicMock()
        sbar = MagicMock()
        sbar.text = msg_text
        sbar.messageType = msg_type

        def find_by_id(id_str: str) -> MagicMock:
            if "sbar" in id_str:
                return sbar
            if id_str == "wnd[0]":
                return wnd0
            if id_str not in components:
                components[id_str] = MagicMock()
            return components[id_str]

        wnd0.findById.side_effect = find_by_id
        session.findById.side_effect = find_by_id
        session._wnd0 = wnd0
        return session

    def test_text_cleaning_removes_week_mention_for_sap(self) -> None:
        self.assertEqual(limpiar_texto_servicio("SERVICIO DE SANEO SEMANA 41"), "SERVICIO DE SANEO")
        self.assertEqual(limpiar_texto_servicio("MANTENIMIENTO SEM 38 PROVINCIA"), "MANTENIMIENTO")
        self.assertEqual(limpiar_texto_servicio("OPERACION LOGISTICA"), "OPERACION LOGISTICA")

    @patch("builtins.input", return_value="")
    def test_empty_confirmation_fails_closed(self, mock_input: MagicMock) -> None:
        session = self._create_mock_session()
        exito, msg, doc_miro, asiento = registrar_fila_miro(
            session,
            fecha_doc="13.10.2026",
            referencia="01-0E001-0000622",
            ruc="20400000622",
            importe=849.60,
            texto_pos="SERVICIO DE TRANSPORTE",
            oc="430062021",
            ind_impto="C1",
            interactivo=True,
            modo_simulacion=False,
        )
        self.assertFalse(exito)
        self.assertIn("Fail-Closed", msg)
        self.assertEqual(doc_miro, "")
        self.assertEqual(asiento, "")

    @patch("builtins.input", return_value="no")
    def test_negative_confirmation_cancels(self, mock_input: MagicMock) -> None:
        session = self._create_mock_session()
        exito, msg, doc_miro, asiento = registrar_fila_miro(
            session,
            fecha_doc="13.10.2026",
            referencia="02-0E001-0000041",
            ruc="10400000041",
            importe=436.99,
            texto_pos="SERV. SANEO",
            oc="430062018",
            ind_impto="C0",
            interactivo=True,
            modo_simulacion=False,
        )
        self.assertFalse(exito)
        self.assertIn("Fail-Closed", msg)
        self.assertEqual(doc_miro, "")

    def test_simulation_mode_never_commits_and_cancels_cleanly(self) -> None:
        session = self._create_mock_session()
        exito, msg, doc_miro, asiento = registrar_fila_miro(
            session,
            fecha_doc="13.10.2026",
            referencia="01-0E001-0000622",
            ruc="20400000622",
            importe=849.60,
            texto_pos="SERVICIO DE TRANSPORTE",
            oc="430062021",
            ind_impto="C1",
            interactivo=False,
            modo_simulacion=True,
        )
        self.assertTrue(exito)
        self.assertIn("Modo prueba", msg)
        self.assertEqual(doc_miro, "SIMULADO")
        self.assertEqual(asiento, "SIMULADO")
        # Verificar que se envió tecla de cancelación (vKey 12 = F12)
        session._wnd0.sendVKey.assert_called_with(12)

    @patch("builtins.input", return_value="SI")
    def test_ambiguous_sap_response_is_blocked_from_marking_success(self, mock_input: MagicMock) -> None:
        # Mensaje de advertencia o informativo que NO contiene números de factura ni asiento
        session = self._create_mock_session(msg_text="Verifique la fecha de contabilización", msg_type="W")
        exito, msg, doc_miro, asiento = registrar_fila_miro(
            session,
            fecha_doc="13.10.2026",
            referencia="01-0E001-0000622",
            ruc="20400000622",
            importe=849.60,
            texto_pos="SERVICIO DE TRANSPORTE",
            oc="430062021",
            ind_impto="C1",
            interactivo=True,
            modo_simulacion=False,
        )
        self.assertFalse(exito)
        self.assertIn("Alerta de Control Financiero", msg)
        self.assertIn("reconciliación manual", msg)
        self.assertEqual(doc_miro, "")
        self.assertEqual(asiento, "")

    @patch("builtins.input", return_value="SI")
    def test_successful_posting_captures_both_numbers(self, mock_input: MagicMock) -> None:
        # Formato oficial de SAP MIRO al contabilizar exitosamente
        official_msg = "Se ha contabilizado el documentoc de factura 5105773695 ( Asiento contable: 4300035763 )"
        session = self._create_mock_session(msg_text=official_msg, msg_type="S")
        exito, msg, doc_miro, asiento = registrar_fila_miro(
            session,
            fecha_doc="13.10.2026",
            referencia="01-0E001-0000622",
            ruc="20400000622",
            importe=849.60,
            texto_pos="SERVICIO DE TRANSPORTE",
            oc="430062021",
            ind_impto="C1",
            interactivo=True,
            modo_simulacion=False,
        )
        self.assertTrue(exito)
        self.assertEqual(doc_miro, "5105773695")
        self.assertEqual(asiento, "4300035763")


if __name__ == "__main__":
    unittest.main()
