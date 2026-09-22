' ==============================================================================
' SOLUCIÓN 2: MACRO NATIVA VBSCRIPT PARA SAP GUI (CERO INSTALACIONES)
' ==============================================================================
' Este script no requiere instalar Python ni ninguna herramienta externa.
' Funciona de forma 100% nativa en cualquier laptop con Windows y SAP GUI.
' Para ejecutarlo: doble clic en el archivo o desde la consola (cscript).
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
           "Por favor, abre SAP GUI e inicia sesion antes de ejecutar este script.", vbExclamation, "SAP no encontrado"
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

' Agregar encabezados de resultado en columnas V y W
objSheet.Cells(1, 22).Value = "Estado MIRO"
objSheet.Cells(1, 23).Value = "N° Documento MIRO"

Dim fechaDoc, refDoc, rucDoc, tPosDoc, impDoc, ocDoc
Dim idCab, strMsg, msgType

For r = 2 To maxRows
    idCab = Trim(CStr(objSheet.Cells(r, 1).Value))
    If idCab = "" Or UCase(idCab) = "TOTAL" Then
        ' Saltear fila vacía o total
    Else
        fechaDoc = Trim(CStr(objSheet.Cells(r, 2).Value))
        refDoc   = Trim(CStr(objSheet.Cells(r, 4).Value))
        rucDoc   = Trim(CStr(objSheet.Cells(r, 6).Value))
        tPosDoc  = Trim(CStr(objSheet.Cells(r, 12).Value))
        impDoc   = Trim(CStr(objSheet.Cells(r, 17).Value))
        ocDoc    = Trim(CStr(objSheet.Cells(r, 21).Value))

        On Error Resume Next
        ' Navegar a MIRO
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nMIRO"
        session.findById("wnd[0]").sendVKey 0
        WScript.Sleep 800

        ' Llenar datos de cabecera
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/ctxtINVFO-BLDAT").Text = fechaDoc
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/ctxtINVFO-BUDAT").Text = fechaDoc
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/txtINVFO-XBLNR").Text = refDoc
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/txtINVFO-WRBTR").Text = FormatNumber(CDbl(impDoc), 2)
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsTS/tabpTAB_HEADER/ssubHEADER:SAPLMR1M:6010/ctxtINVFO-WAERS").Text = "PEN"

        ' Asignar OC
        If ocDoc <> "" Then
            session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/ssubITEMS:SAPLMR1M:6011/ctxtRM08M-EBELN").Text = ocDoc
            session.findById("wnd[0]").sendVKey 0
            WScript.Sleep 1000
        End If

        ' Contabilizar (Guardar)
        session.findById("wnd[0]/tbar[0]/btn[11]").Press
        WScript.Sleep 1500

        strMsg = session.findById("wnd[0]/sbar").Text
        msgType = session.findById("wnd[0]/sbar").MessageType

        If Err.Number = 0 And msgType <> "E" Then
            objSheet.Cells(r, 22).Value = "CONTABILIZADO"
            objSheet.Cells(r, 23).Value = strMsg
            totalExitosos = totalExitosos + 1
        Else
            objSheet.Cells(r, 22).Value = "ERROR"
            objSheet.Cells(r, 23).Value = Err.Description & " " & strMsg
            totalFallidos = totalFallidos + 1
            Err.Clear
        End If
        On Error GoTo 0
        WScript.Sleep 500
    End If
Next

' Guardar archivo de auditoría
Dim outPath
outPath = Replace(excelPath, ".xlsx", " (CON MIRO VBS).xlsx")
objWorkbook.SaveAs outPath
objWorkbook.Close False
objExcel.Quit

Set objSheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing

msgFinal = "Proceso MIRO finalizado exitosamente." & vbCrLf & vbCrLf & _
           "Facturas procesadas: " & (totalExitosos + totalFallidos) & vbCrLf & _
           "Exitosas: " & totalExitosos & vbCrLf & _
           "Con observaciones: " & totalFallidos & vbCrLf & vbCrLf & _
           "Archivo con resultados guardado en:" & vbCrLf & outPath

MsgBox msgFinal, vbInformation, "Registro Masivo SAP MIRO Completado"
