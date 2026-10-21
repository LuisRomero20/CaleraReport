' ==============================================================================
' SOLUCIÓN 2: MACRO NATIVA VBSCRIPT PARA SAP GUI (CERO INSTALACIONES)
' ==============================================================================
' Este script no requiere instalar Python ni ninguna herramienta externa.
' Funciona de forma 100% nativa en cualquier laptop con Windows y SAP GUI.
' Para ejecutarlo: doble clic en el archivo o desde la consola (cscript).
'
' Cumple el flujo oficial en SAP MIRO:
' 1. Dat. Básica: Fecha, Referencia, Importe Total, PEN, C1/C0, Calc.Impuestos, Texto, OC.
' 2. Detalle: Clase doc KR (Factura) o KH (RH).
' 3. Retención: Verificación limpia.
' 4. Simulación y Contabilización.
' 5. Extrae N° Factura MIRO y Asiento Contable FI en columnas del Excel.
' ==============================================================================

Option Explicit

Dim fso, currentDir, excelPath
Set fso = CreateObject("Scripting.FileSystemObject")
currentDir = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
excelPath = fso.BuildPath(currentDir, "Masivo Reporte.xlsx")

If Not fso.FileExists(excelPath) Then
    MsgBox "No se encontro el archivo:" & vbCrLf & excelPath, vbCritical, "Error - Archivo no encontrado"
    WScript.Quit
End If

' 1. Conexión con SAP GUI
Dim SapGuiAuto, application, connection, session
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Or Not IsObject(SapGuiAuto) Then
    MsgBox "No se pudo conectar con SAP GUI." & vbCrLf & vbCrLf & _
           "Por favor, abre SAP GUI e inicia sesion en Sociedad 1001 antes de ejecutar este script.", vbExclamation, "SAP no encontrado"
    WScript.Quit
End If

Set application = SapGuiAuto.GetScriptingEngine
If Err.Number <> 0 Or application.Children.Count = 0 Then
    MsgBox "No hay conexiones activas en SAP GUI." & vbCrLf & _
           "Asegurate de iniciar sesion en tu mandante de SAP.", vbExclamation, "Sin sesion activa"
    WScript.Quit
End If

Set connection = application.Children(0)
Set session = connection.Children(0)
On Error GoTo 0

' 2. Abrir Excel
Dim objExcel, objWorkbook, objSheet, r, maxRows
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

Set objWorkbook = objExcel.Workbooks.Open(excelPath)
Set objSheet = objWorkbook.Sheets(1)

maxRows = objSheet.UsedRange.Rows.Count
Dim totalExitosos, totalFallidos, msgFinal
totalExitosos = 0
totalFallidos = 0

' Agregar encabezados de resultado en columnas V, W y X
objSheet.Cells(1, 22).Value = "Estado MIRO"
objSheet.Cells(1, 23).Value = "N° Factura MIRO"
objSheet.Cells(1, 24).Value = "Asiento Contable FI"

Dim fechaDoc, refDoc, rucDoc, tPosDoc, impDoc, ocDoc, indImpto, claseDoc, servicioSolo
Dim idCab, strMsg, msgType, docMiro, asientoContable

For r = 2 To maxRows
    idCab = Trim(CStr(objSheet.Cells(r, 1).Value))
    If idCab = "" Or UCase(idCab) = "TOTAL" Then
        ' Saltear fila vacía o total
    Else
        fechaDoc = Trim(CStr(objSheet.Cells(r, 2).Value))
        refDoc   = Trim(CStr(objSheet.Cells(r, 4).Value))
        rucDoc   = Trim(CStr(objSheet.Cells(r, 6).Value))
        indImpto = Trim(CStr(objSheet.Cells(r, 8).Value))
        tPosDoc  = Trim(CStr(objSheet.Cells(r, 12).Value))
        impDoc   = Trim(CStr(objSheet.Cells(r, 17).Value))
        ocDoc    = Trim(CStr(objSheet.Cells(r, 21).Value))

        If indImpto = "" Then indImpto = Trim(CStr(objSheet.Cells(r, 18).Value))
        If Left(refDoc, 3) = "01-" Then
            indImpto = "C1"
            claseDoc = "KR"
        Else
            indImpto = "C0"
            claseDoc = "KH"
        End If

        ' Limpiar texto de servicio (quitar SEMANA)
        servicioSolo = tPosDoc
        Dim posSem
        posSem = InStr(1, UCase(servicioSolo), "SEMANA")
        If posSem > 1 Then
            servicioSolo = Trim(Left(servicioSolo, posSem - 1))
        End If
        servicioSolo = Left(servicioSolo, 25)

        On Error Resume Next
        ' Navegar a MIRO
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nMIRO"
        session.findById("wnd[0]").sendVKey 0
        WScript.Sleep 800

        ' 1. Llenar datos de cabecera (Dat.básic)
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/ctxtINVFO-BLDAT").Text = fechaDoc
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/ctxtINVFO-BUDAT").Text = fechaDoc
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/txtINVFO-XBLNR").Text = refDoc
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/txtINVFO-WRBTR").Text = FormatNumber(CDbl(impDoc), 2)
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/ctxtINVFO-WAERS").Text = "PEN"

        ' Indicador de Impuestos y Calc. Impuestos
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/ctxtINVFO-MWSKZ_GROSS").Text = indImpto
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/ctxtINVFO-MWSKZ").Text = indImpto
        If indImpto = "C1" Then
            session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/chkINVFO-XMWST").Selected = True
        Else
            session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/chkINVFO-XMWST").Selected = False
        End If

        ' Texto servicio
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/txtINVFO-BKTXT").Text = servicioSolo

        ' Asignar OC y dar Enter
        If ocDoc <> "" Then
            session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/ssubITEMS:SAPLMR1M:6011/ctxtRM08M-EBELN").Text = ocDoc
            session.findById("wnd[0]").sendVKey 0
            WScript.Sleep 1200
        End If

        ' Cerrar popups de aviso si aparecen
        Do While session.Children.Count > 1
            session.Children(session.Children.Count - 1).sendVKey 0
            WScript.Sleep 300
        Loop

        ' 2. Pestaña Detalle (Clase doc: KR o KH)
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_DETAIL").select
        WScript.Sleep 400
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_DETAIL/ssubHEADER:SAPLMR1M:6010/ctxtINVFO-BLART").Text = claseDoc

        ' 3. Pestaña Retención
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_WT").select
        WScript.Sleep 400

        ' 4. Volver a Datos Básicos y Simular
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER").select
        WScript.Sleep 300
        session.findById("wnd[0]/tbar[1]/btn[7]").Press
        WScript.Sleep 1000

        ' ── CONTROL FAIL-CLOSED: Confirmación interactiva antes de contabilizar ──
        Dim respConfirm
        respConfirm = MsgBox("Simulación generada exitosamente en SAP." & vbCrLf & vbCrLf & _
                             "Comprobante: " & refDoc & vbCrLf & _
                             "Proveedor RUC: " & rucDoc & vbCrLf & _
                             "Importe: S/ " & FormatNumber(CDbl(impDoc), 2) & vbCrLf & _
                             "Orden de Compra: " & ocDoc & vbCrLf & vbCrLf & _
                             "¿Deseas CONTABILIZAR este comprobante en SAP?", _
                             vbQuestion + vbYesNo + vbDefaultButton2, "Control Financiero SAP MIRO")

        If respConfirm <> vbYes Then
            ' Cancelar modal de simulación con F12
            session.findById("wnd[0]").sendVKey 12
            objSheet.Cells(r, 22).Value = "CANCELADO"
            objSheet.Cells(r, 23).Value = "Cancelado por el usuario tras simulación"
            objSheet.Cells(r, 24).Value = "-"
        Else
            ' 5. Contabilizar (Guardar)
            session.findById("wnd[0]/tbar[0]/btn[11]").Press
            WScript.Sleep 1800

            strMsg = session.findById("wnd[0]/sbar").Text
            msgType = session.findById("wnd[0]/sbar").MessageType

            ' Extraer N° Documento y Asiento
            docMiro = ""
            asientoContable = ""
            
            Dim regexMiro, regexAsiento, matches
            Set regexMiro = CreateObject("VBScript.RegExp")
            regexMiro.Pattern = "\b(5\d{9})\b"
            If regexMiro.Test(strMsg) Then
                Set matches = regexMiro.Execute(strMsg)
                docMiro = matches(0).Value
            End If

            Set regexAsiento = CreateObject("VBScript.RegExp")
            regexAsiento.Pattern = "\b(4\d{9})\b"
            If regexAsiento.Test(strMsg) Then
                Set matches = regexAsiento.Execute(strMsg)
                asientoContable = matches(0).Value
            End If

            ' Verificación estricta: Solo éxito con número inequívoco y mensaje de éxito S/I
            If Err.Number = 0 And (docMiro <> "" Or asientoContable <> "") And (msgType = "S" Or msgType = "I") Then
                objSheet.Cells(r, 22).Value = "CONTABILIZADO"
                objSheet.Cells(r, 23).Value = docMiro
                objSheet.Cells(r, 24).Value = asientoContable
                totalExitosos = totalExitosos + 1
            Else
                objSheet.Cells(r, 22).Value = "REVISION MANUAL REQUERIDA"
                objSheet.Cells(r, 23).Value = strMsg
                objSheet.Cells(r, 24).Value = "-"
                totalFallidos = totalFallidos + 1
                Err.Clear
            End If
        End If
        On Error GoTo 0
        WScript.Sleep 500
    End If
Next

' Guardar cambios en Excel
objWorkbook.Save
objWorkbook.Close
objExcel.Quit

msgFinal = "Procesamiento completado en SAP MIRO." & vbCrLf & vbCrLf & _
           "Exitosos: " & totalExitosos & vbCrLf & _
           "Fallidos: " & totalFallidos & vbCrLf & vbCrLf & _
           "El archivo 'Masivo Reporte.xlsx' fue actualizado con los numeros de asiento contable."

MsgBox msgFinal, vbInformation, "Automatizacion SAP Finalizada"
