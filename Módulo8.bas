Attribute VB_Name = "Módulo8"
Option Explicit

Public Sub ProcesarCarpetaRecursiva(ByVal folderPath As String, _
                                     ByRef dicAut As Object, ByRef dicAgg As Object, ByVal tolSec As Long, _
                                     ByVal wsKM As Worksheet, ByRef nmCount As Long, _
                                     ByRef processedRows As Long, ByRef matchedRows As Long, _
                                     ByRef noMatchRows As Long, ByRef processedSheets As Long, _
                                     ByVal wsTL As Worksheet)
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Dim f As String

    f = Dir(folderPath & "*.csv"): Do While Len(f) > 0
        currentFileIndex = currentFileIndex + 1
        SetStatus "Procesando archivo " & currentFileIndex & " / " & totalFilesExpected & "  (" & f & ")"
        ProcesarArchivoVE folderPath & f, ExtraerDivisionDesdeNombre(f), dicAut, dicAgg, tolSec, wsKM, nmCount, f, True, processedRows, matchedRows, noMatchRows, processedSheets, wsTL
        totalFiles = totalFiles + 1
        f = Dir
    Loop

    f = Dir(folderPath & "*.xlsx"): Do While Len(f) > 0
        currentFileIndex = currentFileIndex + 1
        SetStatus "Procesando archivo " & currentFileIndex & " / " & totalFilesExpected & "  (" & f & ")"
        ProcesarArchivoVE folderPath & f, ExtraerDivisionDesdeNombre(f), dicAut, dicAgg, tolSec, wsKM, nmCount, f, False, processedRows, matchedRows, noMatchRows, processedSheets, wsTL
        totalFiles = totalFiles + 1
        f = Dir
    Loop

    f = Dir(folderPath & "*.xlsm"): Do While Len(f) > 0
        currentFileIndex = currentFileIndex + 1
        SetStatus "Procesando archivo " & currentFileIndex & " / " & totalFilesExpected & "  (" & f & ")"
        ProcesarArchivoVE folderPath & f, ExtraerDivisionDesdeNombre(f), dicAut, dicAgg, tolSec, wsKM, nmCount, f, False, processedRows, matchedRows, noMatchRows, processedSheets, wsTL
        totalFiles = totalFiles + 1
        f = Dir
    Loop

    f = Dir(folderPath & "*.xlsb"): Do While Len(f) > 0
        currentFileIndex = currentFileIndex + 1
        SetStatus "Procesando archivo " & currentFileIndex & " / " & totalFilesExpected & "  (" & f & ")"
        ProcesarArchivoVE folderPath & f, ExtraerDivisionDesdeNombre(f), dicAut, dicAgg, tolSec, wsKM, nmCount, f, False, processedRows, matchedRows, noMatchRows, processedSheets, wsTL
        totalFiles = totalFiles + 1
        f = Dir
    Loop

    f = Dir(folderPath & "*.xls"): Do While Len(f) > 0
        currentFileIndex = currentFileIndex + 1
        SetStatus "Procesando archivo " & currentFileIndex & " / " & totalFilesExpected & "  (" & f & ")"
        ProcesarArchivoVE folderPath & f, ExtraerDivisionDesdeNombre(f), dicAut, dicAgg, tolSec, wsKM, nmCount, f, False, processedRows, matchedRows, noMatchRows, processedSheets, wsTL
        totalFiles = totalFiles + 1
        f = Dir
    Loop

    Dim subf As String
    subf = Dir(folderPath & "*", vbDirectory)
    Do While Len(subf) > 0
        If subf <> "." And subf <> ".." Then
            If (GetAttr(folderPath & subf) And vbDirectory) <> 0 Then
                Call ProcesarCarpetaRecursiva(folderPath & subf, dicAut, dicAgg, tolSec, wsKM, nmCount, processedRows, matchedRows, noMatchRows, processedSheets, wsTL)
            End If
        End If
        subf = Dir
    Loop
End Sub

Public Sub ProcesarArchivoVE(ByVal fullPath As String, ByVal divisionName As String, _
                              ByRef dicAut As Object, ByRef dicAgg As Object, ByVal tolSec As Long, _
                              ByVal wsKM As Worksheet, ByRef nmCount As Long, ByVal fileShort As String, _
                              ByVal isCSV As Boolean, _
                              ByRef processedRows As Long, ByRef matchedRows As Long, _
                              ByRef noMatchRows As Long, ByRef processedSheets As Long, ByVal wsTL As Worksheet)

    Dim wbTmp As Workbook, ws As Worksheet, onlyOne As Boolean

    If isCSV Then
        Set wbTmp = Workbooks.Open(fileName:=fullPath, Local:=True)
        onlyOne = True
    Else
        Set wbTmp = Workbooks.Open(fileName:=fullPath, ReadOnly:=True)
        onlyOne = False
    End If

    If onlyOne Then
        Set ws = wbTmp.Worksheets(1)
        totalSheets = totalSheets + 1
        LogFileSheet fileShort, ws.name, divisionName, Not IsReporteUnidadesSheet(ws.name)
        If Not IsReporteUnidadesSheet(ws.name) Then
            If ws.UsedRange.rows.count > 1 And ws.UsedRange.Columns.count > 1 Then
                ProcesarHojaVE ws, divisionName, dicAut, dicAgg, tolSec, wsKM, nmCount, fileShort, _
                               processedRows, matchedRows, noMatchRows, wsTL
                processedSheets = processedSheets + 1
            End If
        End If
    Else
        Dim wsIt As Worksheet
        For Each wsIt In wbTmp.Worksheets
            totalSheets = totalSheets + 1
            LogFileSheet fileShort, wsIt.name, divisionName, Not IsReporteUnidadesSheet(wsIt.name)
            If Not IsReporteUnidadesSheet(wsIt.name) Then
                If wsIt.UsedRange.rows.count > 1 And wsIt.UsedRange.Columns.count > 1 Then
                    ProcesarHojaVE wsIt, divisionName, dicAut, dicAgg, tolSec, wsKM, nmCount, _
                                   fileShort & " | " & wsIt.name, processedRows, matchedRows, noMatchRows, wsTL
                    processedSheets = processedSheets + 1
                End If
            End If
        Next wsIt
    End If

    wbTmp.Close SaveChanges:=False
End Sub

Public Sub ProcesarHojaVE(ByVal ws As Worksheet, ByVal divisionName As String, _
    ByRef dicAut As Object, ByRef dicAgg As Object, ByVal tolSec As Long, _
    ByVal wsKM As Worksheet, ByRef nmCount As Long, ByVal srcTag As String, _
    ByRef processedRows As Long, ByRef matchedRows As Long, _
    ByRef noMatchRows As Long, ByVal wsTL As Worksheet)

    Dim hdr As Long
    Dim aVeh, aFS, aFFS, aH1, aH2, aKms
    Dim vVeh As Long, vFS As Long, vFFS As Long, vH1 As Long, vH2 As Long, vKms As Long
    Dim lrVeh As Long, lrKms As Long, lastRow As Long
    Dim i As Long
    Dim veh As String
    Dim d1 As Double, d2 As Double
    Dim h1 As Long, h2 As Long
    Dim kms As Double
    Dim diaIni As Long, diaFin As Long
    Dim dd As Long
    Dim segIni As Long, segFin As Long
    Dim movTotalDur As Double
    Dim segIniD As Long, segFinD As Long, durD As Long
    Dim kmsDia As Double
    Dim overlapTotal As Long
    Dim keyD As String
    Dim col As Collection
    Dim j As Long
    Dim ar As Variant
    Dim iniA As Long, finA As Long
    Dim s As Long, e As Long, ov As Long
    Dim kmsMatchD As Double
    Dim kmsNoMatchD As Double
    Dim absStart As Double, absEnd As Double, catPick As String

    hdr = FindHeaderRow(ws)
    aVeh = Array("Unidad", "Carro", "Vehiculo", "Vehículo")
    aFS = Array("Fecha Inicio", "F Servicio", "Fecha", "F_Servicio", "FServicio")
    aFFS = Array("Fecha Fin", "F FServicio", "F_FServicio", "FFServicio", "Fecha")
    aH1 = Array("Hora Inicio", "HHMM1", "HoraInicial", "Inicio", "HI")
    aH2 = Array("Hora Fin", "HHMM2", "HoraFinal", "Fin", "HF")
    aKms = Array("Kilómetros", "kms", "km", "kilometros", "kilómetros", "kilometraje")

    vVeh = FindColAnyInRow(ws, hdr, aVeh)
    vFS = FindColAnyInRow(ws, hdr, aFS)
    vFFS = FindColAnyInRow(ws, hdr, aFFS)
    vH1 = FindColAnyInRow(ws, hdr, aH1)
    vH2 = FindColAnyInRow(ws, hdr, aH2)
    vKms = FindColAnyInRow(ws, hdr, aKms)

    If vVeh = 0 Or vFS = 0 Or vKms = 0 Then
        LogLine ThisWorkbook, srcTag, hdr, vVeh, vFS, vFFS, vH1, vH2, vKms, 0, 0, 0, "Columnas clave no encontradas"
        Exit Sub
    End If
    If vFFS = 0 Then vFFS = vFS
    If vH1 = 0 Then vH1 = vFS
    If vH2 = 0 Then vH2 = vFFS

    lrVeh = ws.Cells(ws.rows.count, vVeh).End(xlUp).row
    lrKms = ws.Cells(ws.rows.count, vKms).End(xlUp).row
    lastRow = IIf(lrVeh > lrKms, lrVeh, lrKms)
    If lastRow <= hdr Then
        LogLine ThisWorkbook, srcTag, hdr, vVeh, vFS, vFFS, vH1, vH2, vKms, lrVeh, lrKms, 0, "Sin datos bajo encabezado"
        Exit Sub
    End If

    LogLine ThisWorkbook, srcTag, hdr, vVeh, vFS, vFFS, vH1, vH2, vKms, lrVeh, lrKms, lastRow - hdr, ""

    For i = hdr + 1 To lastRow
        If (i - hdr) Mod 200 = 0 Then
            Dim pctHoja As Double: pctHoja = (i - hdr) / (lastRow - hdr)
            SetStatus "Archivo " & currentFileIndex & "/" & totalFilesExpected & " | Hoja: " & ws.name, pctHoja
        End If

        Dim rowKey As String
        rowKey = Normalize(srcTag) & "|" & CStr(i)
        If seenRows.Exists(rowKey) Then GoTo NextRow Else seenRows.Add rowKey, 1

        veh = Trim$(CStr(ws.Cells(i, vVeh).value))
If veh = "" Then GoTo NextRow
If Not VehLenOK(veh, 6) Then GoTo NextRow


        d1 = DateOnlyEx2(ws.Cells(i, vFS).value, MOV_DATE_ORDER)
        d2 = DateOnlyEx2(ws.Cells(i, vFFS).value, MOV_DATE_ORDER)
        If d1 = 0 And d2 > 0 Then d1 = d2
        If d2 = 0 And d1 > 0 Then d2 = d1
        If d1 = 0 Then GoTo NextRow

        h1 = TimeToSecEx(ws.Cells(i, vH1).value)
        h2 = TimeToSecEx(ws.Cells(i, vH2).value)
        If h2 < h1 And d2 = d1 Then d2 = d1 + 1

        kms = KmsToDouble(ws.Cells(i, vKms).value)
        If kms <= 0 Then GoTo NextRow
        processedRows = processedRows + 1

        diaIni = CLng(d1)
        diaFin = CLng(d2)
        If diaFin < diaIni Then diaFin = diaIni

        movTotalDur = 0
        Dim absMovStart As Double, absMovEnd As Double
        absMovStart = d1 * 86400# + h1
        absMovEnd = d2 * 86400# + h2
        For dd = diaIni To diaFin
            If dd = diaIni Then segIni = h1 Else segIni = 0
            If dd = diaFin Then segFin = h2 Else segFin = 86400
            If segFin < segIni Then segFin = segIni
            movTotalDur = movTotalDur + (segFin - segIni)
        Next dd
        If movTotalDur <= 0 Then GoTo NextRow

        ' Si hay llegada real dentro del movimiento => NO SLOPE (asignar 100% km al día de llegada)
        Dim allocDay As Long, hasArr As Boolean
        hasArr = FindArrivalDayForMovement(veh, absMovStart, absMovEnd, allocDay)

        For dd = diaIni To diaFin
            If dd = diaIni Then segIniD = h1 Else segIniD = 0
            If dd = diaFin Then segFinD = h2 Else segFinD = 86400
            If segFinD < segIniD Then segFinD = segIniD
            durD = segFinD - segIniD
            If durD <= 0 Then GoTo SiguienteDia

            If hasArr Then
                If dd = allocDay Then
                    kmsDia = kms
                Else
                    kmsDia = 0
                End If
            Else
                kmsDia = kms * (durD / movTotalDur)
            End If
            If kmsDia <= 0 Then GoTo SiguienteDia

            Call WriteTimelineForDaySegment(divisionName, veh, dd, segIniD, segFinD, _
                                            kmsDia, durD, dicAut, tolSec, dicVisitWins, wsTL)

            overlapTotal = 0
            keyD = veh & "|" & CStr(dd)
            If dicAut.Exists(keyD) Then
                Set col = dicAut(keyD)
                For j = 1 To col.count
                    ar = col(j)
                    iniA = CLng(ar(1)) - (TOLERANCIA_MIN * 60)
                    finA = CLng(ar(2)) + (TOLERANCIA_MIN * 60)
                    If iniA < 0 Then iniA = 0
                    If finA > 86400 Then finA = 86400
                    s = IIf(segIniD > iniA, segIniD, iniA)
                    e = IIf(segFinD < finA, segFinD, finA)
                    If e > s Then ov = e - s Else ov = 0
                    overlapTotal = overlapTotal + ov
                Next j
            End If

            If overlapTotal <= 0 Then
                kmsMatchD = 0
                noMatchRows = noMatchRows + 1
                absStart = CDbl(dd) * 86400# + segIniD
                absEnd = CDbl(dd) * 86400# + segFinD
                catPick = ""
                Call AttributeNoMatchToVisitCats(divisionName, veh, dd, absStart, absEnd, kmsDia, durD, srcTag, wsKM, catPick)
            Else
                If overlapTotal > durD Then overlapTotal = durD
                kmsMatchD = kmsDia * (CDbl(overlapTotal) / CDbl(durD))
                matchedRows = matchedRows + 1

                kmsNoMatchD = kmsDia - kmsMatchD
                If kmsNoMatchD > 0.0000005 Then
                    absStart = CDbl(dd) * 86400# + segIniD
                    absEnd = CDbl(dd) * 86400# + segFinD
                    catPick = ""
                    Call AttributeNoMatchToVisitCats(divisionName, veh, dd, _
                                                     absStart, absEnd, _
                                                     kmsNoMatchD, durD, _
                                                     srcTag, wsKM, catPick)
                End If
            End If

            AggSum dicAgg, divisionName, veh, dd, kmsMatch:=kmsMatchD, kmsTot:=kmsDia
SiguienteDia:
        Next dd

NextRow:
    Next i
End Sub

Private Sub AggSum(ByRef dicAgg As Object, ByVal divn As String, ByVal veh As String, ByVal fecha As Double, ByVal kmsMatch As Double, ByVal kmsTot As Double)
    Dim key As String
    Dim vals As Variant
    key = divn & "|" & veh & "|" & CStr(fecha)
    If dicAgg.Exists(key) Then
        vals = dicAgg(key)
        vals(0) = CDbl(vals(0)) + kmsTot
        vals(1) = CDbl(vals(1)) + kmsMatch
        dicAgg(key) = vals
    Else
        vals = Array(CDbl(kmsTot), CDbl(kmsMatch))
        dicAgg.Add key, vals
    End If
End Sub

Public Sub LoadBaseDatosFromSheet(ByVal wsBD As Worksheet, ByRef dLleg As Object)
    ' A: Unidad | C: División | D: Cliente | F: Tipo de servicio
    ' G: Fecha ini | H: Hora ini | I: Fecha fin | J: Hora fin
    ' Q: Hora programada | S: Hora real | T: DiffMin (abs) | U: Estado
    Dim lr As Long, r As Long
    Dim veh As String, divn As String, cliente As String, tserv As String
    Dim dIni As Double, dFin As Double, hIni As Long, hFin As Long
    Dim hProg As Long, hReal As Long, absLleg As Double
    Dim diffMin As Double, estado As String
    Dim dayInt As Long

    lr = wsBD.Cells(wsBD.rows.count, 1).End(xlUp).row
    If lr < 2 Then Exit Sub

    For r = 2 To lr
        veh = Trim$(CStr(wsBD.Cells(r, 1).value))
        If Len(veh) = 0 Then GoTo NextR

        divn = Trim$(CStr(wsBD.Cells(r, 3).value))
        cliente = Trim$(CStr(wsBD.Cells(r, 4).value))
        tserv = Trim$(CStr(wsBD.Cells(r, 6).value))

        dIni = DateOnlyEx2(wsBD.Cells(r, 7).value, "DMY")
        dFin = DateOnlyEx2(wsBD.Cells(r, 9).value, "DMY")
        hIni = TimeToSecEx(wsBD.Cells(r, 8).value)
        hFin = TimeToSecEx(wsBD.Cells(r, 10).value)

        hProg = TimeToSecEx(wsBD.Cells(r, 17).value)
        hReal = TimeToSecEx(wsBD.Cells(r, 19).value)
        diffMin = Val(wsBD.Cells(r, 20).value)
        estado = Trim$(CStr(wsBD.Cells(r, 21).value))

        If dIni > 0 And hReal > 0 Then
            absLleg = dIni * 86400# + hReal
        ElseIf dFin > 0 And hReal > 0 Then
            absLleg = dFin * 86400# + hReal
        ElseIf dIni > 0 And hIni > 0 Then
            absLleg = dIni * 86400# + hIni
        Else
            GoTo NextR
        End If

        dayInt = CLng(Int(absLleg / 86400#))
        If Not dLleg.Exists(veh) Then dLleg.Add veh, New Collection
        dLleg(veh).Add Array(absLleg, cliente, divn, diffMin, estado, tserv, dayInt)
NextR:
    Next r
End Sub

