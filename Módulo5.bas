Attribute VB_Name = "Módulo5"
Option Explicit

Public Function BuildKMCotizadosPorCliente(ByVal wsCot As Worksheet, ByRef dDiasVistos As Object) As Object
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare
    If wsCot Is Nothing Then Set BuildKMCotizadosPorCliente = dic: Exit Function

    Dim hdr As Long: hdr = FindHeaderRowBD(wsCot)
    Dim cCli As Long, cFS As Long, cFFS As Long, cH1 As Long, cH2 As Long
    cCli = FindColAnyInRow(wsCot, hdr, Array("C_Cliente", "Cliente"))
    cFS = FindColAnyInRow(wsCot, hdr, Array("F_Servicio", "Fecha Inicio", "Fecha", "F_Servicio", "FServicio"))
    cFFS = FindColAnyInRow(wsCot, hdr, Array("F_FServicio", "Fecha Fin", "F_FServicio", "FFServicio", "Fecha")): If cFFS = 0 Then cFFS = cFS
    cH1 = FindColAnyInRow(wsCot, hdr, Array("HoraInicial", "Hora Inicio", "HHMM1", "Inicio", "HI"))
    cH2 = FindColAnyInRow(wsCot, hdr, Array("HoraFinal", "Hora Fin", "HHMM2", "Fin", "HF"))
    If cCli = 0 Or cFS = 0 Or cH1 = 0 Or cH2 = 0 Then Set BuildKMCotizadosPorCliente = dic: Exit Function

    Dim KMcols As Variant: KMcols = Array(23, 24, 25, 26, 27) ' W:X:Y:Z:AA
    Dim lr As Long: lr = wsCot.Cells(wsCot.rows.count, cCli).End(xlUp).row
    Dim r As Long, ii As Long, cli As String
    Dim d1 As Double, d2 As Double, h1 As Long, h2 As Long
    Dim diaIni As Long, diaFin As Long, dd As Long
    Dim segIni As Long, segFin As Long, durD As Long, totDur As Double
    Dim kmCotRow As Double

    For r = hdr + 1 To lr
        cli = Trim$(CStr(wsCot.Cells(r, cCli).value)): If Len(cli) = 0 Then GoTo NextR

        kmCotRow = 0
        For ii = LBound(KMcols) To UBound(KMcols)
            kmCotRow = kmCotRow + KmsToDouble(wsCot.Cells(r, KMcols(ii)).value)
        Next ii
        If kmCotRow <= 0 Then GoTo NextR

        d1 = DateOnlyEx2(wsCot.Cells(r, cFS).value, COT_DATE_ORDER)
        d2 = DateOnlyEx2(wsCot.Cells(r, cFFS).value, COT_DATE_ORDER)
        If d1 = 0 And d2 > 0 Then d1 = d2
        If d2 = 0 And d1 > 0 Then d2 = d1
        If d1 = 0 Then GoTo NextR
        h1 = TimeToSecEx(wsCot.Cells(r, cH1).value)
        h2 = TimeToSecEx(wsCot.Cells(r, cH2).value)
        If h2 < h1 And d2 = d1 Then d2 = d1 + 1

        diaIni = CLng(d1): diaFin = CLng(d2): If diaFin < diaIni Then diaFin = diaIni
        totDur = 0
        For dd = diaIni To diaFin
            If dd = diaIni Then segIni = h1 Else segIni = 0
            If dd = diaFin Then segFin = h2 Else segFin = 86400
            If segFin < segIni Then segFin = segIni
            totDur = totDur + (segFin - segIni)
        Next dd
        If totDur <= 0 Then totDur = 1

        For dd = diaIni To diaFin
            If (dDiasVistos Is Nothing) Or dDiasVistos.Exists(CStr(dd)) Then
                If dd = diaIni Then segIni = h1 Else segIni = 0
                If dd = diaFin Then segFin = h2 Else segFin = 86400
                If segFin < segIni Then segFin = segIni
                durD = segFin - segIni
                If durD > 0 Then
                    If dic.Exists(cli) Then
                        dic(cli) = CDbl(dic(cli)) + kmCotRow * (durD / totDur)
                    Else
                        dic.Add cli, kmCotRow * (durD / totDur)
                    End If
                End If
            End If
        Next dd
NextR:
    Next r
    Set BuildKMCotizadosPorCliente = dic
End Function

Public Function BuildTipoCortoPorCliente(ByVal wsCot As Worksheet, ByRef dDiasVistos As Object) As Object
    Dim outDic As Object: Set outDic = CreateObject("Scripting.Dictionary")
    outDic.CompareMode = vbTextCompare
    If wsCot Is Nothing Then Set BuildTipoCortoPorCliente = outDic: Exit Function

    Dim cCli As Long, cTipo As Long, cFS As Long, cFF As Long
    cCli = ColIndexExact(wsCot, "C_Cliente"): If cCli = 0 Then cCli = FindColAnyInRow(wsCot, 1, Array("C_Cliente", "Cliente"))
    cTipo = ColIndexExact(wsCot, "D_Servicio_Tipo_Corto"): If cTipo = 0 Then cTipo = wsCot.Range("V1").Column
    cFS = ColIndexExact(wsCot, "F_Servicio")
    cFF = ColIndexExact(wsCot, "F_FServicio"): If cFF = 0 Then cFF = cFS

    Dim missingCols As String
    If cCli = 0 Then missingCols = missingCols & vbCrLf & " - C_Cliente"
    If cTipo = 0 Then missingCols = missingCols & vbCrLf & " - D_Servicio_Tipo_Corto"
    If cFS = 0 Then missingCols = missingCols & vbCrLf & " - F_Servicio"
    If Len(missingCols) > 0 Then
        WarnMissingColumns wsCot.name, missingCols
        Set BuildTipoCortoPorCliente = outDic
        Exit Function
    End If

    Dim store As Object: Set store = CreateObject("Scripting.Dictionary")
    store.CompareMode = vbTextCompare

    Dim lr As Long: lr = wsCot.Cells(wsCot.rows.count, cCli).End(xlUp).row
    Dim r As Long, cli0 As String, Tipo0 As String
    Dim d1 As Double, d2 As Double, dd As Long
    Dim includeRow As Boolean

    For r = 2 To lr
        cli0 = Trim$(CStr(wsCot.Cells(r, cCli).value))
        Tipo0 = UCase$(Trim$(CStr(wsCot.Cells(r, cTipo).value)))
        If Tipo0 <> "GD" And Tipo0 <> "VE" And Tipo0 <> "FC" Then GoTo NextR

        d1 = DateOnlyEx2(wsCot.Cells(r, cFS).value, COT_DATE_ORDER)
        d2 = DateOnlyEx2(wsCot.Cells(r, cFF).value, COT_DATE_ORDER)
        If d1 = 0 And d2 > 0 Then d1 = d2
        If d2 = 0 And d1 > 0 Then d2 = d1
        If d1 = 0 Then GoTo NextR
        If d2 < d1 Then d2 = d1

        includeRow = (dDiasVistos Is Nothing)
        If Not includeRow Then
            For dd = CLng(d1) To CLng(d2)
                If dDiasVistos.Exists(CStr(dd)) Then includeRow = True: Exit For
            Next dd
        End If
        If Not includeRow Then GoTo NextR

        cli0 = MapClienteNull(cli0, Tipo0)
        If Len(cli0) = 0 Then GoTo NextR

        Dim dCli As Object, rec As Variant, lastSeen As Long
        If store.Exists(cli0) Then
            Set dCli = store(cli0)
        Else
            Set dCli = CreateObject("Scripting.Dictionary")
            dCli.CompareMode = vbTextCompare
            store.Add cli0, dCli
        End If

        If dCli.Exists(Tipo0) Then
            rec = dCli(Tipo0)
        Else
            rec = Array(0&, 0&)
        End If
        rec(0) = CLng(rec(0)) + 1
        lastSeen = CLng(d2)
        If lastSeen > CLng(rec(1)) Then rec(1) = lastSeen
        dCli(Tipo0) = rec

NextR:
    Next r

    Dim k As Variant, dt As Object, t As Variant
    Dim bestType As String, bestCnt As Long, bestLast As Long, v As Variant
    For Each k In store.keys
        Set dt = store(k)
        bestType = "": bestCnt = -1: bestLast = -1
        For Each t In dt.keys
            v = dt(t)
            If CLng(v(0)) > bestCnt Or (CLng(v(0)) = bestCnt And CLng(v(1)) > bestLast) Then
                bestCnt = CLng(v(0)): bestLast = CLng(v(1)): bestType = CStr(t)
            End If
        Next t
        If bestType <> "" Then
            outDic(k) = bestType
        End If
    Next k

    Set BuildTipoCortoPorCliente = outDic
End Function

Public Function CalcularKMVaciosAsignadosDesdeTL( _
    ByVal wsTL As Worksheet, _
    ByRef dDiasVistos As Object, _
    Optional ByRef dLleg As Object = Nothing, _
    Optional ByRef totalVaciosTL As Double = 0#, _
    Optional ByRef totalAsignado As Double = 0#) As Object
    Dim dicOut As Object: Set dicOut = CreateObject("Scripting.Dictionary"): dicOut.CompareMode = vbTextCompare
    If wsTL Is Nothing Then On Error Resume Next: Set wsTL = ThisWorkbook.Worksheets(HOJA_TL): On Error GoTo 0
    If wsTL Is Nothing Then Set CalcularKMVaciosAsignadosDesdeTL = dicOut: Exit Function

    Dim cVeh As Long, cTipo As Long, cIni As Long, cKm As Long, cCli As Long
    cVeh = ColIndexExact(wsTL, "Vehiculo")
    cTipo = ColIndexExact(wsTL, "Tipo")
    cIni = ColIndexExact(wsTL, "Inicio")
    cKm = ColIndexExact(wsTL, "Km")
    cCli = GetClienteSiteVisitColumn(wsTL)

    Dim missingTLCols As String
    If cVeh = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Vehiculo"
    If cTipo = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Tipo"
    If cIni = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Inicio"
    If cKm = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Km"
    If cCli = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Cliente / SiteVisit"
    If Len(missingTLCols) > 0 Then
        WarnMissingColumns wsTL.name, missingTLCols
        Set CalcularKMVaciosAsignadosDesdeTL = dicOut
        Exit Function
    End If

    Const FALLBACK_CLIENTE As String = "(Sin cliente)"

    ' veh -> colección de eventos: Array(ts As Double, isCli As Boolean, cli As String, isVac As Boolean, km As Double)
    Dim bucket As Object: Set bucket = CreateObject("Scripting.Dictionary")
    bucket.CompareMode = vbTextCompare

    Dim totalVaciosTL_local As Double: totalVaciosTL_local = 0#
    Dim totalAsignado_local As Double: totalAsignado_local = 0#

    Dim lr As Long: lr = wsTL.Cells(wsTL.rows.count, cVeh).End(xlUp).row
    Dim r As Long, veh As String, tipo As String, cli As String, ts As Double, km As Double, dd As Long
    For r = 2 To lr
        veh = Trim$(CStr(wsTL.Cells(r, cVeh).value)): If Len(veh) = 0 Then GoTo NextR
        If Not IsDate(wsTL.Cells(r, cIni).value) Then GoTo NextR
        ts = CDbl(CDate(wsTL.Cells(r, cIni).value))
        dd = CLng(Int(ts))
        If Not (dDiasVistos Is Nothing) Then If Not dDiasVistos.Exists(CStr(dd)) Then GoTo NextR

        tipo = LCase$(Trim$(CStr(wsTL.Cells(r, cTipo).value)))
        km = 0: If IsNumeric(wsTL.Cells(r, cKm).value) Then km = CDbl(wsTL.Cells(r, cKm).value)
        If km <= 0 Then GoTo NextR

        Dim isVac As Boolean
        isVac = ( _
            tipo = "otros" Or _
            tipo = "taller" Or _
            tipo = "diesel" Or _
            tipo = "patio" Or _
            tipo = "inicio" Or _
            tipo = "fin" Or _
            tipo = "enganche")
        Dim isCli As Boolean: isCli = (tipo = "match")
        If Not (isVac Or isCli) Then GoTo NextR

        If isCli Then cli = Trim$(CStr(wsTL.Cells(r, cCli).value)) Else cli = ""

        If isVac Then totalVaciosTL_local = totalVaciosTL_local + km

        If Not bucket.Exists(veh) Then bucket.Add veh, New Collection
        bucket(veh).Add Array(ts, isCli, cli, isVac, km)
NextR:
    Next r

    ' Procesa por vehículo en orden cronológico (arrastre entre días)
    Dim k As Variant, col As Collection, n As Long
    Dim evs() As Variant, i As Long
    For Each k In bucket.keys
        Set col = bucket(k): n = col.count
        If n = 0 Then GoTo NextK

        ' Copiar a arreglo y ordenar por ts asc usando insertion sort
        ReDim evs(1 To n)
        For i = 1 To n
            evs(i) = col(i)
        Next i

        Dim j As Long
        Dim temp As Variant
        For i = 2 To n
            temp = evs(i)
            j = i - 1
            Do While j >= 1 And CDbl(evs(j)(0)) > CDbl(temp(0))
                evs(j + 1) = evs(j)
                j = j - 1
            Loop
            evs(j + 1) = temp
        Next i

        Dim vac_pending As Double: vac_pending = 0#
        Dim half_for_next As Double: half_for_next = 0#
        Dim lastCli As String: lastCli = ""
        Dim lastTwoA As String: lastTwoA = "": Dim lastTwoB As String: lastTwoB = ""
        Dim vacSinceLastCli As Boolean: vacSinceLastCli = False

        For i = 1 To n
            Dim currentEvent As Variant
            currentEvent = evs(i)

            Dim eIsCli As Boolean: eIsCli = CBool(currentEvent(1))
            Dim eCli As String:     eCli = CStr(currentEvent(2))
            Dim eIsVac As Boolean:  eIsVac = CBool(currentEvent(3))
            Dim eKm As Double:      eKm = CDbl(currentEvent(4))
            Dim eDay As Long:       eDay = CLng(Int(CDbl(currentEvent(0))))

            If eIsVac Then
                ' Acumula vacíos consecutivos en un solo bloque
                vac_pending = vac_pending + eKm
                vacSinceLastCli = True

            ElseIf eIsCli Then
                Dim resolvedCli As String
                resolvedCli = eCli
                If Len(resolvedCli) = 0 Then
                    resolvedCli = ResolveFallbackClienteTL(CStr(k), CDbl(currentEvent(0)), eDay, dLleg, FALLBACK_CLIENTE)
                End If

                ' Paga la mitad reservada de un enganche previo, si existe
                If half_for_next > 0# And Len(resolvedCli) > 0 Then
                    If dicOut.Exists(resolvedCli) Then
                        dicOut(resolvedCli) = CDbl(dicOut(resolvedCli)) + half_for_next
                    Else
                        dicOut.Add resolvedCli, half_for_next
                    End If
                    totalAsignado_local = totalAsignado_local + half_for_next
                    half_for_next = 0#
                End If

                ' Regla principal: asigna los vacíos acumulados al siguiente cliente
                If vac_pending > 0# Then
                    Dim nextIsClient As Boolean: nextIsClient = False
                    If i < n Then
                        Dim nextEvent As Variant
                        nextEvent = evs(i + 1)
                        If CBool(nextEvent(1)) Then nextIsClient = True  ' enganche cliente->cliente sin vacíos intermedios
                    End If

                    If nextIsClient Then
                        ' Enganche: 50/50 entre este cliente y el siguiente
                        Dim share As Double: share = vac_pending / 2#
                        If Len(resolvedCli) > 0 Then
                            If dicOut.Exists(resolvedCli) Then
                                dicOut(resolvedCli) = CDbl(dicOut(resolvedCli)) + share
                            Else
                                dicOut.Add resolvedCli, share
                            End If
                            totalAsignado_local = totalAsignado_local + share
                        End If
                        half_for_next = share
                        vac_pending = 0#
                    Else
                        ' Todo el bloque al cliente actual
                        If Len(resolvedCli) > 0 Then
                            If dicOut.Exists(resolvedCli) Then
                                dicOut(resolvedCli) = CDbl(dicOut(resolvedCli)) + vac_pending
                            Else
                                dicOut.Add resolvedCli, vac_pending
                            End If
                            totalAsignado_local = totalAsignado_local + vac_pending
                        End If
                        vac_pending = 0#
                    End If
                End If

                ' Tracking de pares consecutivos de clientes (para cierre sin cliente futuro)
                If Len(lastCli) > 0 And Not vacSinceLastCli Then
                    lastTwoA = lastCli: lastTwoB = resolvedCli
                Else
                    lastTwoA = resolvedCli: lastTwoB = ""
                End If
                lastCli = resolvedCli
                vacSinceLastCli = False
            End If
        Next i

        ' Fin de datos: si aún hay vacíos pendientes, asigna al/los últimos clientes vistos
        If vac_pending > 0# Then
            If Len(lastTwoB) > 0 Then
                Dim half As Double: half = vac_pending / 2#
                If Len(lastTwoA) > 0 Then
                    If dicOut.Exists(lastTwoA) Then
                        dicOut(lastTwoA) = CDbl(dicOut(lastTwoA)) + half
                    Else
                        dicOut.Add lastTwoA, half
                    End If
                    totalAsignado_local = totalAsignado_local + half
                End If
                If dicOut.Exists(lastTwoB) Then
                    dicOut(lastTwoB) = CDbl(dicOut(lastTwoB)) + half
                Else
                    dicOut.Add lastTwoB, half
                End If
                totalAsignado_local = totalAsignado_local + half
            ElseIf Len(lastCli) > 0 Then
                If dicOut.Exists(lastCli) Then
                    dicOut(lastCli) = CDbl(dicOut(lastCli)) + vac_pending
                Else
                    dicOut.Add lastCli, vac_pending
                End If
                totalAsignado_local = totalAsignado_local + vac_pending
            Else
                Debug.Print "[CalcularKMVaciosAsignadosDesdeTL] Vacíos sin cliente para veh=" & CStr(k) & _
                            " -> " & Format(vac_pending, "0.00") & " km"
                If dicOut.Exists(FALLBACK_CLIENTE) Then
                    dicOut(FALLBACK_CLIENTE) = CDbl(dicOut(FALLBACK_CLIENTE)) + vac_pending
                Else
                    dicOut.Add FALLBACK_CLIENTE, vac_pending
                End If
                totalAsignado_local = totalAsignado_local + vac_pending
            End If
            vac_pending = 0#
        End If

NextK:
    Next k

    totalVaciosTL = totalVaciosTL_local
    totalAsignado = totalAsignado_local

    If Abs(totalAsignado_local - totalVaciosTL_local) > 0.01 Then
        Debug.Print "[CalcularKMVaciosAsignadosDesdeTL] Diferencia en KM vacíos asignados: TL=" & Format(totalVaciosTL_local, "0.00") & _
                    " vs asignados=" & Format(totalAsignado_local, "0.00")
    End If

    Set CalcularKMVaciosAsignadosDesdeTL = dicOut
End Function

Private Function ResolveFallbackClienteTL( _
    ByVal vehiculo As String, _
    ByVal ts As Double, _
    ByVal dayInt As Long, _
    ByRef dLleg As Object, _
    ByVal defaultCliente As String) As String

    Dim fallback As String: fallback = defaultCliente

    If dayInt <= 0 Then
        ResolveFallbackClienteTL = fallback
        Exit Function
    End If

    If Not dLleg Is Nothing Then
        If dLleg.Exists(vehiculo) Then
            Dim col As Collection
            Set col = dLleg(vehiculo)
            Dim idx As Long, rec As Variant
            Dim bestDiff As Double: bestDiff = -1#
            Dim targetAbs As Double: targetAbs = CDbl(ts) * 86400#
            For idx = 1 To col.count
                rec = col(idx)
                If IsArray(rec) Then
                    If UBound(rec) >= 6 Then
                        If CLng(rec(6)) = dayInt Then
                            Dim cli0 As String
                            cli0 = Trim$(CStr(rec(1)))
                            If Len(cli0) > 0 Then
                                Dim absLleg As Double: absLleg = CDbl(rec(0))
                                Dim diff As Double: diff = Abs(absLleg - targetAbs)
                                If bestDiff < 0# Or diff < bestDiff Then
                                    bestDiff = diff
                                    fallback = cli0
                                    If diff = 0# Then Exit For
                                End If
                            End If
                        End If
                    End If
                End If
            Next idx
        End If
    End If

    ResolveFallbackClienteTL = fallback
End Function

Public Sub CrearResumenClienteTotal_RangoFijo( _
    ByVal wb As Workbook, _
    ByRef dLleg As Object, _
    Optional ByRef dicAgg As Object = Nothing, _
    Optional ByRef dicKMCotizados As Object = Nothing, _
    Optional ByVal wsBD As Worksheet = Nothing, _
    Optional ByVal fechaInicio As Variant, _
    Optional ByVal fechaFin As Variant)

    Dim dDiasRango As Object: Set dDiasRango = CreateObject("Scripting.Dictionary")
    Dim dIni As Double, dFin As Double, dd As Long
    Dim fechaInicioRaw As Variant, fechaFinRaw As Variant

    If IsMissing(fechaInicio) Or IsEmpty(fechaInicio) Then
        fechaInicioRaw = Trim$(CStr(InputBox("Ingrese la fecha de inicio (dd/mm/aaaa)", _
                                            "Rango de fechas", Format$(Date, "dd/mm/yyyy"))))
        If Len(fechaInicioRaw) = 0 Then Exit Sub
    Else
        fechaInicioRaw = fechaInicio
    End If

    If IsMissing(fechaFin) Or IsEmpty(fechaFin) Then
        Dim defaultFin As String
        If IsDate(fechaInicioRaw) Then
            defaultFin = Format$(CDate(fechaInicioRaw), "dd/mm/yyyy")
        Else
            defaultFin = CStr(fechaInicioRaw)
        End If
        fechaFinRaw = Trim$(CStr(InputBox("Ingrese la fecha de fin (dd/mm/aaaa)", _
                                          "Rango de fechas", defaultFin)))
        If Len(fechaFinRaw) = 0 Then Exit Sub
    Else
        fechaFinRaw = fechaFin
    End If

    If IsDate(fechaInicioRaw) Then
        dIni = CDbl(CDate(fechaInicioRaw))
    Else
        dIni = DateOnlyEx2(CStr(fechaInicioRaw), "DMY")
    End If

    If IsDate(fechaFinRaw) Then
        dFin = CDbl(CDate(fechaFinRaw))
    Else
        dFin = DateOnlyEx2(CStr(fechaFinRaw), "DMY")
    End If

    If dIni = 0 Or dFin = 0 Then
        MsgBox "No se pudo interpretar las fechas proporcionadas.", vbExclamation
        Exit Sub
    End If
    If dFin < dIni Then dFin = dIni
    For dd = CLng(dIni) To CLng(dFin)
        dDiasRango.Add CStr(dd), True
    Next dd
    Dim rangoTxt As String
    rangoTxt = Format$(CDate(dIni), "dd/mm/yyyy") & " - " & Format$(CDate(dFin), "dd/mm/yyyy")

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("Resumen_Cliente_Total")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.count))
        ws.name = "Resumen_Cliente_Total"
    End If
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Clear

    ws.Range("A1:R1").value = Array( _
        "Cliente", "Fecha", "Servicios", "En tiempo", "Retardo", "% En tiempo", _
        "Min. prom. diferencia", "Min. tot. diferencia", "D_Servicio_Tipo_Corto", _
        "Servicio SS", "Servicio SM", "Servicio RESS", "Servicio RESM", _
        "KM_Realizados", "KM_vacios_asignados", "KM_Cotizados", _
        "Realizados_vs_Cotizados", "Eficiencia_Cliente")

    ' Reutiliza las mismas fuentes y reglas que el resumen normal, pero filtrando por dDiasRango
    Dim wsTL As Worksheet
    On Error Resume Next: Set wsTL = wb.Worksheets(HOJA_TL): On Error GoTo 0

    Dim dicCliMatchTL As Object
    Set dicCliMatchTL = SumarKMMatchDesdeTLPorCliente(wsTL)

    Dim dicBase As Object, dicTipos As Object
    Set dicBase = CreateObject("Scripting.Dictionary"): dicBase.CompareMode = vbTextCompare
    Set dicTipos = CreateObject("Scripting.Dictionary"): dicTipos.CompareMode = vbTextCompare

    Dim k As Variant, col As Collection, i As Long
    Dim wv As Variant, cliente As String, estado As String, diffMin As Double, tipo As String, dayInt As Long, vbArr As Variant
    If Not dLleg Is Nothing Then
        For Each k In dLleg.keys
            Set col = dLleg(k)
            For i = 1 To col.count
                wv = col(i)
                dayInt = CLng(wv(6))
                If dDiasRango.Exists(CStr(dayInt)) Then
                    cliente = Trim$(CStr(wv(1))): If Len(cliente) = 0 Then GoTo NextW
                    estado = LCase$(CStr(wv(4)))
                    diffMin = CDbl(wv(3))
                    tipo = CStr(wv(5))
                    If dicBase.Exists(cliente) Then vbArr = dicBase(cliente) Else vbArr = Array(0#, 0#, 0#, 0#)
                    vbArr(0) = vbArr(0) + 1
                    If estado = "en tiempo" Then vbArr(1) = vbArr(1) + 1 Else vbArr(2) = vbArr(2) + 1
                    vbArr(3) = vbArr(3) + diffMin
                    dicBase(cliente) = vbArr
                    Dim keyTipo As String: keyTipo = cliente & "||" & tipo
                    If dicTipos.Exists(keyTipo) Then dicTipos(keyTipo) = CLng(dicTipos(keyTipo)) + 1 Else dicTipos.Add keyTipo, 1
                End If
NextW:
            Next i
        Next k
    End If

    Dim dicCliKM_BD As Object
    Set dicCliKM_BD = SumarKMRealizadosDesdeBDPorCliente(wsBD, dDiasRango, dLleg, dicAgg)

    Dim wsCot As Worksheet, dicTipoCorto As Object
    On Error Resume Next: Set wsCot = wb.Worksheets(HOJA_COT): On Error GoTo 0
    Set dicTipoCorto = BuildTipoCortoPorCliente(wsCot, dDiasRango)

    Dim dicVacAsign As Object
    Dim totalVacTL_Rango As Double, totalVacAsignado_Rango As Double
    Set dicVacAsign = CalcularKMVaciosAsignadosDesdeTL(wsTL, dDiasRango, dLleg, totalVacTL_Rango, totalVacAsignado_Rango)

    Dim dicAll As Object: Set dicAll = CreateObject("Scripting.Dictionary"): dicAll.CompareMode = vbTextCompare
    Dim aCli As Variant
    For Each aCli In dicCliMatchTL.keys
        If Not dicAll.Exists(CStr(aCli)) Then dicAll.Add CStr(aCli), True
    Next aCli
    For Each aCli In dicBase.keys
        If Not dicAll.Exists(CStr(aCli)) Then dicAll.Add CStr(aCli), True
    Next aCli
    For Each aCli In dicVacAsign.keys
        If Not dicAll.Exists(CStr(aCli)) Then dicAll.Add CStr(aCli), True
    Next aCli

    Dim r As Long: r = 2
    Dim s As Double, onT As Double, ret As Double, sumD As Double
    Dim pOn As Double, avgD As Double, kmReal As Double, kmCot As Double
    Dim ssCt As Long, smCt As Long, ressCt As Long, resmCt As Long
    Dim tipoCorto As String, kmVacAsign As Double, ratioRealVsCot As Double, effOp As Double
    Dim totalVacResumen As Double: totalVacResumen = 0#

    For Each aCli In dicAll.keys
        If dicCliMatchTL.Exists(aCli) Then kmReal = CDbl(dicCliMatchTL(aCli)) Else kmReal = 0
        If Not dicKMCotizados Is Nothing And dicKMCotizados.Exists(CStr(aCli)) Then kmCot = CDbl(dicKMCotizados(CStr(aCli))) Else kmCot = 0

        If dicBase.Exists(aCli) Then
            vbArr = dicBase(aCli)
            s = vbArr(0): onT = vbArr(1): ret = vbArr(2): sumD = vbArr(3)
        Else
            s = 0: onT = 0: ret = 0: sumD = 0
        End If
        If s > 0 Then pOn = onT / s Else pOn = 0
        If s > 0 Then avgD = sumD / s Else avgD = 0

        ssCt = 0: smCt = 0: ressCt = 0: resmCt = 0
        If dicTipos.Exists(CStr(aCli) & "||SS") Then ssCt = CLng(dicTipos(CStr(aCli) & "||SS"))
        If dicTipos.Exists(CStr(aCli) & "||SM") Then smCt = CLng(dicTipos(CStr(aCli) & "||SM"))
        If dicTipos.Exists(CStr(aCli) & "||RESS") Then ressCt = CLng(dicTipos(CStr(aCli) & "||RESS"))
        If dicTipos.Exists(CStr(aCli) & "||RESM") Then resmCt = CLng(dicTipos(CStr(aCli) & "||RESM"))

        If Not dicTipoCorto Is Nothing And dicTipoCorto.Exists(CStr(aCli)) Then tipoCorto = CStr(dicTipoCorto(CStr(aCli))) Else tipoCorto = ""
        If dicVacAsign.Exists(CStr(aCli)) Then kmVacAsign = CDbl(dicVacAsign(CStr(aCli))) Else kmVacAsign = 0
        totalVacResumen = totalVacResumen + kmVacAsign

        If kmCot > 0 Then ratioRealVsCot = kmReal / kmCot Else ratioRealVsCot = 0
        If (kmReal + kmVacAsign) > 0 Then effOp = kmReal / (kmReal + kmVacAsign) Else effOp = 0

        ws.Cells(r, 1).value = aCli
        ws.Cells(r, 2).value = rangoTxt
        ws.Cells(r, 3).value = s
        ws.Cells(r, 4).value = onT
        ws.Cells(r, 5).value = ret
        ws.Cells(r, 6).value = pOn
        ws.Cells(r, 7).value = avgD
        ws.Cells(r, 8).value = sumD
        ws.Cells(r, 9).value = tipoCorto
        ws.Cells(r, 10).value = ssCt
        ws.Cells(r, 11).value = smCt
        ws.Cells(r, 12).value = ressCt
        ws.Cells(r, 13).value = resmCt
        ws.Cells(r, 14).value = kmReal
        ws.Cells(r, 15).value = kmVacAsign
        ws.Cells(r, 16).value = kmCot
        ws.Cells(r, 17).value = ratioRealVsCot
        ws.Cells(r, 18).value = effOp
        r = r + 1
    Next aCli

    If Abs(totalVacResumen - totalVacAsignado_Rango) > 0.01 Then
        Debug.Print "[Resumen_Cliente_Total] Diferencia entre resumen y asignación de KM vacíos: resumen=" & _
                    Format(totalVacResumen, "0.00") & " vs asignados=" & Format(totalVacAsignado_Rango, "0.00")
    End If
    If Abs(totalVacAsignado_Rango - totalVacTL_Rango) > 0.01 Then
        Debug.Print "[Resumen_Cliente_Total] Diferencia entre KM vacíos TL y asignados: TL=" & _
                    Format(totalVacTL_Rango, "0.00") & " vs asignados=" & Format(totalVacAsignado_Rango, "0.00")
    End If

    If r > 2 Then
        ws.Range("F2:F" & r - 1).NumberFormat = "0.0%"
        ws.Range("G2:H" & r - 1).NumberFormat = "0.0"
        ws.Range("N2:P" & r - 1).NumberFormat = "0.0"
        ws.Range("Q2:Q" & r - 1).NumberFormat = "0.0%"
        ws.Range("R2:R" & r - 1).NumberFormat = "0.0%"
        ws.Columns.AutoFit
        ws.Range("A1").AutoFilter
    Else
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
    End If
End Sub

Public Sub CrearResumenClienteFiltrado( _
    ByVal wb As Workbook, _
    ByRef dLleg As Object, _
    ByRef dDiasVistos As Object, _
    Optional ByRef dicAgg As Object = Nothing, _
    Optional ByRef dicKMCotizados As Object = Nothing, _
    Optional ByVal wsBD As Worksheet = Nothing)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("Resumen_Cliente")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.count))
        ws.name = "Resumen_Cliente"
    End If

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Clear

    ' A Cliente | B Fecha(rango) | C..R resto
    ws.Range("A1:R1").value = Array( _
        "Cliente", "Fecha", "Servicios", "En tiempo", "Retardo", "% En tiempo", _
        "Min. prom. diferencia", "Min. tot. diferencia", "D_Servicio_Tipo_Corto", _
        "Servicio SS", "Servicio SM", "Servicio RESS", "Servicio RESM", _
        "KM_Realizados", "KM_vacios_asignados", "KM_Cotizados", _
        "Realizados_vs_Cotizados", "Eficiencia_Cliente")

    Dim dicBase As Object, dicTipos As Object
    Set dicBase = CreateObject("Scripting.Dictionary"): dicBase.CompareMode = vbTextCompare
    Set dicTipos = CreateObject("Scripting.Dictionary"): dicTipos.CompareMode = vbTextCompare

    Dim k As Variant, col As Collection, i As Long
    Dim wv As Variant, cliente As String, estado As String, tipo As String
    Dim diffMin As Double, dayInt As Long, vbArr As Variant

    If Not dLleg Is Nothing Then
        For Each k In dLleg.keys
            Set col = dLleg(k)
            For i = 1 To col.count
                wv = col(i) ' [absLleg, Cliente, División, DiffMin, Estado, TipoServ, dayInt]
                dayInt = CLng(wv(6))
                If dDiasVistos Is Nothing Or dDiasVistos.Exists(CStr(dayInt)) Then
                    cliente = Trim$(CStr(wv(1))): If Len(cliente) = 0 Then GoTo NextW
                    estado = LCase$(CStr(wv(4)))
                    diffMin = CDbl(wv(3))
                    tipo = CStr(wv(5))

                    If dicBase.Exists(cliente) Then
                        vbArr = dicBase(cliente)
                    Else
                        vbArr = Array(0#, 0#, 0#, 0#)
                    End If
                    vbArr(0) = vbArr(0) + 1
                    If estado = "en tiempo" Then vbArr(1) = vbArr(1) + 1 Else vbArr(2) = vbArr(2) + 1
                    vbArr(3) = vbArr(3) + diffMin
                    dicBase(cliente) = vbArr

                    Dim keyTipo As String: keyTipo = cliente & "||" & tipo
                    If dicTipos.Exists(keyTipo) Then
                        dicTipos(keyTipo) = CLng(dicTipos(keyTipo)) + 1
                    Else
                        dicTipos.Add keyTipo, 1
                    End If
                End If
NextW:
            Next i
        Next k
    End If

    Dim wsTL As Worksheet
    On Error Resume Next: Set wsTL = wb.Worksheets(HOJA_TL): On Error GoTo 0

    Dim dicCliMatchTL As Object
    Set dicCliMatchTL = SumarKMMatchDesdeTLPorCliente(wsTL)

    Dim dicTotalAsignado As Object, dicSoporte As Object
    Set dicSoporte = CreateObject("Scripting.Dictionary"): dicSoporte.CompareMode = vbTextCompare
    Set dicTotalAsignado = SumarKMCliente_TLxBD(wsTL, wsBD, dDiasVistos, dLleg, dicSoporte)

    Dim dicCliKM_BD As Object
    Set dicCliKM_BD = SumarKMRealizadosDesdeBDPorCliente(wsBD, dDiasVistos, dLleg, dicAgg)

    Dim wsCot As Worksheet, dicTipoCorto As Object
    On Error Resume Next: Set wsCot = wb.Worksheets(HOJA_COT): On Error GoTo 0
    Set dicTipoCorto = BuildTipoCortoPorCliente(wsCot, dDiasVistos)

    Dim dicVacAsign As Object
    Dim totalVacTL_Filtro As Double, totalVacAsignado_Filtro As Double
    Set dicVacAsign = CalcularKMVaciosAsignadosDesdeTL(wsTL, dDiasVistos, dLleg, totalVacTL_Filtro, totalVacAsignado_Filtro)

    Dim dicAll As Object: Set dicAll = CreateObject("Scripting.Dictionary"): dicAll.CompareMode = vbTextCompare
    Dim aCli As Variant
    For Each aCli In dicCliMatchTL.keys
        If Not dicAll.Exists(CStr(aCli)) Then dicAll.Add CStr(aCli), True
    Next aCli
    For Each aCli In dicBase.keys
        If Not dicAll.Exists(CStr(aCli)) Then dicAll.Add CStr(aCli), True
    Next aCli
    For Each aCli In dicVacAsign.keys
        If Not dicAll.Exists(CStr(aCli)) Then dicAll.Add CStr(aCli), True
    Next aCli

    ' Rango de fechas visual (min-max)
    Dim minD As Double: minD = 1E+99
    Dim maxD As Double: maxD = -1
    If Not dDiasVistos Is Nothing And dDiasVistos.count > 0 Then
        For Each k In dDiasVistos.keys
            If CDbl(k) < minD Then minD = CDbl(k)
            If CDbl(k) > maxD Then maxD = CDbl(k)
        Next k
    ElseIf Not dicAgg Is Nothing And dicAgg.count > 0 Then
        Dim arrK() As String
        For Each k In dicAgg.keys
            arrK = Split(CStr(k), "|")
            If UBound(arrK) = 2 Then
                If CDbl(arrK(2)) < minD Then minD = CDbl(arrK(2))
                If CDbl(arrK(2)) > maxD Then maxD = CDbl(arrK(2))
            End If
        Next k
    End If
    Dim rangoTxt As String
    If maxD >= 0 And minD < 1E+99 Then
        rangoTxt = Format(CDate(minD), "dd/mm/yyyy") & " - " & Format(CDate(maxD), "dd/mm/yyyy")
    Else
        rangoTxt = ""
    End If

    Dim r As Long: r = 2
    Dim s As Double, onT As Double, ret As Double, sumD As Double
    Dim pOn As Double, avgD As Double, kmReal As Double, kmCot As Double
    Dim ssCt As Long, smCt As Long, ressCt As Long, resmCt As Long
    Dim prefBase As String, kTipoKey As String, tipoCorto As String
    Dim kmVacAsign As Double, ratioRealVsCot As Double, effOp As Double
    Dim totalVacResumen As Double: totalVacResumen = 0#

    For Each aCli In dicAll.keys
        If dicCliMatchTL.Exists(aCli) Then kmReal = CDbl(dicCliMatchTL(aCli)) Else kmReal = 0
        If Not dicKMCotizados Is Nothing And dicKMCotizados.Exists(CStr(aCli)) Then kmCot = CDbl(dicKMCotizados(CStr(aCli))) Else kmCot = 0

        If dicBase.Exists(aCli) Then
            vbArr = dicBase(aCli)
            s = vbArr(0): onT = vbArr(1): ret = vbArr(2): sumD = vbArr(3)
        Else
            s = 0: onT = 0: ret = 0: sumD = 0
        End If
        If s > 0 Then pOn = onT / s Else pOn = 0
        If s > 0 Then avgD = sumD / s Else avgD = 0

        ssCt = 0: smCt = 0: ressCt = 0: resmCt = 0
        prefBase = CStr(aCli) & "||"
        kTipoKey = prefBase & "SS":   If dicTipos.Exists(kTipoKey) Then ssCt = CLng(dicTipos(kTipoKey))
        kTipoKey = prefBase & "SM":   If dicTipos.Exists(kTipoKey) Then smCt = CLng(dicTipos(kTipoKey))
        kTipoKey = prefBase & "RESS": If dicTipos.Exists(kTipoKey) Then ressCt = CLng(dicTipos(kTipoKey))
        kTipoKey = prefBase & "RESM": If dicTipos.Exists(kTipoKey) Then resmCt = CLng(dicTipos(kTipoKey))

        If Not dicTipoCorto Is Nothing And dicTipoCorto.Exists(CStr(aCli)) Then
            tipoCorto = CStr(dicTipoCorto(CStr(aCli)))
        Else
            tipoCorto = ""
        End If

        If dicVacAsign.Exists(CStr(aCli)) Then kmVacAsign = CDbl(dicVacAsign(CStr(aCli))) Else kmVacAsign = 0
        totalVacResumen = totalVacResumen + kmVacAsign

        If kmCot > 0 Then ratioRealVsCot = kmReal / kmCot Else ratioRealVsCot = 0
        If (kmReal + kmVacAsign) > 0 Then effOp = kmReal / (kmReal + kmVacAsign) Else effOp = 0

        ws.Cells(r, 1).value = aCli
        ws.Cells(r, 2).value = rangoTxt          ' Fecha rango
        ws.Cells(r, 3).value = s
        ws.Cells(r, 4).value = onT
        ws.Cells(r, 5).value = ret
        ws.Cells(r, 6).value = pOn
        ws.Cells(r, 7).value = avgD
        ws.Cells(r, 8).value = sumD
        ws.Cells(r, 9).value = tipoCorto
        ws.Cells(r, 10).value = ssCt
        ws.Cells(r, 11).value = smCt
        ws.Cells(r, 12).value = ressCt
        ws.Cells(r, 13).value = resmCt
        ws.Cells(r, 14).value = kmReal
        ws.Cells(r, 15).value = kmVacAsign
        ws.Cells(r, 16).value = kmCot
        ws.Cells(r, 17).value = ratioRealVsCot
        ws.Cells(r, 18).value = effOp
        r = r + 1
    Next aCli

    If Abs(totalVacResumen - totalVacAsignado_Filtro) > 0.01 Then
        Debug.Print "[Resumen_Cliente] Diferencia entre resumen y asignación de KM vacíos: resumen=" & _
                    Format(totalVacResumen, "0.00") & " vs asignados=" & Format(totalVacAsignado_Filtro, "0.00")
    End If
    If Abs(totalVacAsignado_Filtro - totalVacTL_Filtro) > 0.01 Then
        Debug.Print "[Resumen_Cliente] Diferencia entre KM vacíos TL y asignados: TL=" & _
                    Format(totalVacTL_Filtro, "0.00") & " vs asignados=" & Format(totalVacAsignado_Filtro, "0.00")
    End If

    If r > 2 Then
        ws.Range("F2:F" & r - 1).NumberFormat = "0.0%"
        ws.Range("G2:G" & r - 1).NumberFormat = "0.0"
        ws.Range("H2:H" & r - 1).NumberFormat = "0.0"
        ws.Range("N2:P" & r - 1).NumberFormat = "0.0"
        ws.Range("Q2:Q" & r - 1).NumberFormat = "0.0%"
        ws.Range("R2:R" & r - 1).NumberFormat = "0.0%"
        ws.Columns.AutoFit
        ws.Range("A1").AutoFilter
    Else
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
    End If
End Sub

Public Function SumarKMRealizadosPorCliente_TLPrimero( _
    ByVal wsTL As Worksheet, _
    ByRef dicAgg As Object, _
    ByVal wsBD As Worksheet, _
    ByRef dDiasVistos As Object, _
    ByRef dicLleg As Object) As Object

    Dim dicOut As Object: Set dicOut = CreateObject("Scripting.Dictionary"): dicOut.CompareMode = vbTextCompare
    If dicAgg Is Nothing Then Set SumarKMRealizadosPorCliente_TLPrimero = dicOut: Exit Function
    If wsTL Is Nothing Then On Error Resume Next: Set wsTL = ThisWorkbook.Worksheets(HOJA_TL): On Error GoTo 0

    Dim cDiv As Long, cVeh As Long, cTipo As Long, cIni As Long, cKm As Long, cCliCol As Long
    If Not wsTL Is Nothing Then
        cDiv = ColIndexExact(wsTL, "Division")
        cVeh = ColIndexExact(wsTL, "Vehiculo")
        cTipo = ColIndexExact(wsTL, "Tipo")
        cIni = ColIndexExact(wsTL, "Inicio")
        cKm = ColIndexExact(wsTL, "Km")
        cCliCol = GetClienteSiteVisitColumn(wsTL)

        Dim missingTLCols As String
        If cDiv = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Division"
        If cVeh = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Vehiculo"
        If cTipo = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Tipo"
        If cIni = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Inicio"
        If cKm = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Km"
        If cCliCol = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Cliente / SiteVisit"
        If Len(missingTLCols) > 0 Then
            WarnMissingColumns wsTL.name, missingTLCols
            Set wsTL = Nothing
        End If
    End If

    ' 1) TL Match por cliente y por (div|veh|dd)
    Dim tlMatchCli As Object: Set tlMatchCli = CreateObject("Scripting.Dictionary"): tlMatchCli.CompareMode = vbTextCompare
    Dim tlMatchByKey As Object: Set tlMatchByKey = CreateObject("Scripting.Dictionary")
    If Not wsTL Is Nothing Then
        Dim lr As Long: lr = wsTL.Cells(wsTL.rows.count, 1).End(xlUp).row
        Dim r As Long, tp As String, km As Double, divn As String, veh0 As String, dd0 As Double, cli0 As String, key0 As String
        For r = 2 To lr
            tp = LCase$(Trim$(CStr(wsTL.Cells(r, cTipo).value)))
            If tp = "match" Then
                If IsNumeric(wsTL.Cells(r, cKm).value) Then
                    km = CDbl(wsTL.Cells(r, cKm).value)
                    If km > 0 Then
                        If cCliCol > 0 Then cli0 = Trim$(CStr(wsTL.Cells(r, cCliCol).value)) Else cli0 = ""
                        If Len(cli0) > 0 Then
                            If tlMatchCli.Exists(cli0) Then tlMatchCli(cli0) = CDbl(tlMatchCli(cli0)) + km Else tlMatchCli.Add cli0, km
                        End If
                        divn = CStr(wsTL.Cells(r, cDiv).value)
                        veh0 = CStr(wsTL.Cells(r, cVeh).value)
                        If IsDate(wsTL.Cells(r, cIni).value) Then dd0 = Int(CDbl(wsTL.Cells(r, cIni).value)) Else dd0 = 0
                        If dd0 > 0 Then
                            key0 = divn & "|" & veh0 & "|" & CStr(dd0)
                            If tlMatchByKey.Exists(key0) Then tlMatchByKey(key0) = CDbl(tlMatchByKey(key0)) + km Else tlMatchByKey.Add key0, km
                        End If
                    End If
                End If
            End If
        Next r
    End If

    ' 2) Uniones en BD por veh|dd|cliente
    Dim mapVD As Object: Set mapVD = BuildVehDayClientMap(dicLleg, dDiasVistos)
    Dim unionSec As Object: Set unionSec = CreateObject("Scripting.Dictionary")   ' veh|dd -> dict cli->seg
    Dim unionTot As Object: Set unionTot = CreateObject("Scripting.Dictionary")   ' veh|dd -> total seg

    If wsBD Is Nothing Then On Error Resume Next: Set wsBD = ThisWorkbook.Worksheets("Base de datos"): On Error GoTo 0
    If Not wsBD Is Nothing Then
        Dim hdr As Long: hdr = FindHeaderRowBD(wsBD)
        Dim cB_Veh As Long, cB_Cli As Long, cB_FS As Long, cB_FF As Long, cB_H1 As Long, cB_H2 As Long
        cB_Veh = FindColAnyInRow(wsBD, hdr, Array("K_Carro", "Unidad", "Carro", "Vehiculo", "Vehículo"))
        cB_Cli = FindColAnyInRow(wsBD, hdr, Array("C_Cliente", "Cliente"))
        cB_FS = FindColAnyInRow(wsBD, hdr, Array("fecha_inicial", "Fecha Inicio", "F Servicio", "Fecha", "F_Servicio", "FServicio"))
        cB_FF = FindColAnyInRow(wsBD, hdr, Array("fecha_final", "Fecha Fin", "F FServicio", "F_FServicio", "FFServicio", "Fecha")): If cB_FF = 0 Then cB_FF = cB_FS
        cB_H1 = FindColAnyInRow(wsBD, hdr, Array("hora_inicial", "Hora Inicio", "HHMM1", "HoraInicial", "Inicio", "HI"))
        cB_H2 = FindColAnyInRow(wsBD, hdr, Array("hora_final", "Hora Fin", "HHMM2", "HoraFinal", "Fin", "HF"))

        If cB_Veh * cB_FS * cB_FF * cB_H1 * cB_H2 > 0 Then
            Dim lrB As Long: lrB = wsBD.Cells(wsBD.rows.count, cB_Veh).End(xlUp).row
            Dim rB As Long, vVeh As String, vCli As String
            Dim d1 As Double, d2 As Double, h1 As Long, h2 As Long
            Dim diaIni As Long, diaFin As Long, ddB As Long
            Dim segIni As Long, segFin As Long
            Dim intervals As Object: Set intervals = CreateObject("Scripting.Dictionary")
            Dim keyVD0 As String

            For rB = hdr + 1 To lrB
                vVeh = Trim$(CStr(wsBD.Cells(rB, cB_Veh).value))
                If Len(vVeh) = 0 Then GoTo NextRB

                If cB_Cli > 0 Then vCli = Trim$(CStr(wsBD.Cells(rB, cB_Cli).value)) Else vCli = ""
                d1 = DateOnlyEx2(wsBD.Cells(rB, cB_FS).value, "DMY")
                d2 = DateOnlyEx2(wsBD.Cells(rB, cB_FF).value, "DMY")
                If d1 = 0 And d2 > 0 Then d1 = d2
                If d2 = 0 And d1 > 0 Then d2 = d1
                If d1 = 0 Then GoTo NextRB

                h1 = TimeToSecEx(wsBD.Cells(rB, cB_H1).value)
                h2 = TimeToSecEx(wsBD.Cells(rB, cB_H2).value)
                If h2 < h1 And d2 = d1 Then d2 = d1 + 1

                diaIni = CLng(d1): diaFin = CLng(d2): If diaFin < diaIni Then diaFin = diaIni

                For ddB = diaIni To diaFin
                    If Not (dDiasVistos Is Nothing) Then If Not dDiasVistos.Exists(CStr(ddB)) Then GoTo NextDD
                    If ddB = diaIni Then segIni = h1 Else segIni = 0
                    If ddB = diaFin Then segFin = h2 Else segFin = 86400
                    If segFin < segIni Then segFin = segIni
                    If segFin - segIni <= 0 Then GoTo NextDD

                    If Len(vCli) = 0 Then
                        keyVD0 = vVeh & "|" & CStr(ddB)
                        If mapVD.Exists(keyVD0) Then vCli = CStr(mapVD(keyVD0)) Else vCli = ""
                        If Len(vCli) = 0 Then GoTo NextDD
                    End If

                    Dim keyIC As String: keyIC = vVeh & "|" & CStr(ddB) & "|" & vCli
                    Dim colI1 As Collection
                    If intervals.Exists(keyIC) Then
                        Set colI1 = intervals(keyIC)
                    Else
                        Set colI1 = New Collection
                        intervals.Add keyIC, colI1
                    End If
                    colI1.Add Array(segIni, segFin)
NextDD:
                Next ddB
NextRB:
            Next rB

            ' Uniones por cliente
            Dim kIC As Variant, parts() As String, vehD As String, ddOnly As Long, cliName As String
            For Each kIC In intervals.keys
                parts = Split(CStr(kIC), "|")
                vehD = parts(0): ddOnly = CLng(parts(1)): cliName = parts(2)
                Dim colI2 As Collection: Set colI2 = intervals(kIC)

                Dim n As Long: n = colI2.count
                Dim arrI() As Double: ReDim arrI(1 To n, 1 To 2)
                Dim ii As Long
                For ii = 1 To n
                    arrI(ii, 1) = CDbl(colI2(ii)(0))
                    arrI(ii, 2) = CDbl(colI2(ii)(1))
                Next ii

                Dim i1 As Long, j1 As Long, t1 As Double, t2 As Double
                For i1 = 1 To n - 1
                    For j1 = i1 + 1 To n
                        If arrI(j1, 1) < arrI(i1, 1) Then
                            t1 = arrI(i1, 1): t2 = arrI(i1, 2)
                            arrI(i1, 1) = arrI(j1, 1): arrI(i1, 2) = arrI(j1, 2)
                            arrI(j1, 1) = t1: arrI(j1, 2) = t2
                        End If
                    Next j1
                Next i1

                Dim curS As Double, curE As Double, first As Boolean, uni As Double
                first = True: uni = 0
                For ii = 1 To n
                    If first Then
                        curS = arrI(ii, 1): curE = arrI(ii, 2): first = False
                    Else
                        If arrI(ii, 1) <= curE Then
                            If arrI(ii, 2) > curE Then curE = arrI(ii, 2)
                        Else
                            uni = uni + (curE - curS)
                            curS = arrI(ii, 1): curE = arrI(ii, 2)
                        End If
                    End If
                Next ii
                If Not first Then uni = uni + (curE - curS)

                Dim keyVD2 As String: keyVD2 = vehD & "|" & CStr(ddOnly)
                Dim dCli As Object
                If unionSec.Exists(keyVD2) Then
                    Set dCli = unionSec(keyVD2)
                Else
                    Set dCli = CreateObject("Scripting.Dictionary"): dCli.CompareMode = vbTextCompare
                    unionSec.Add keyVD2, dCli
                End If
                If dCli.Exists(cliName) Then dCli(cliName) = CDbl(dCli(cliName)) + uni Else dCli.Add cliName, uni

                If unionTot.Exists(keyVD2) Then
                    unionTot(keyVD2) = CDbl(unionTot(keyVD2)) + uni
                Else
                    unionTot.Add keyVD2, uni
                End If
            Next kIC
        End If
    End If

    ' 3) Salida inicia con TL-Match por cliente
    Dim clientKey As Variant
    For Each clientKey In tlMatchCli.keys
        dicOut(clientKey) = CDbl(tlMatchCli(clientKey))
    Next clientKey

    ' 4) Reparto de residuo por dicAgg
    Dim kAgg As Variant, arrK() As String, div1 As String, veh1 As String, dd1 As Long
    Dim vAgg As Variant, kmTot As Double, kmMatchTL_Key As Double, resid As Double
    For Each kAgg In dicAgg.keys
        arrK = Split(CStr(kAgg), "|")
        If UBound(arrK) = 2 Then
            div1 = arrK(0): veh1 = arrK(1): dd1 = CLng(CDbl(arrK(2)))
            vAgg = dicAgg(kAgg): kmTot = CDbl(vAgg(0))
            If tlMatchByKey.Exists(CStr(kAgg)) Then kmMatchTL_Key = CDbl(tlMatchByKey(CStr(kAgg))) Else kmMatchTL_Key = 0
            resid = kmTot - kmMatchTL_Key: If resid <= 0 Then GoTo NextAgg

            Dim keyVD3 As String: keyVD3 = veh1 & "|" & CStr(dd1)
            Dim dCli2 As Object, totU As Double, cKey As Variant, share As Double

            If unionSec.Exists(keyVD3) Then
                Set dCli2 = unionSec(keyVD3): totU = 0
                If unionTot.Exists(keyVD3) Then totU = CDbl(unionTot(keyVD3))
                If totU > 0 Then
                    For Each cKey In dCli2.keys
                        share = resid * (CDbl(dCli2(cKey)) / totU)
                        If dicOut.Exists(cKey) Then dicOut(cKey) = CDbl(dicOut(cKey)) + share Else dicOut.Add cKey, share
                    Next cKey
                    GoTo NextAgg
                End If
            End If

            ' Fallback: dominante por día
            Dim dom As String: dom = ""
            If mapVD.Exists(keyVD3) Then dom = CStr(mapVD(keyVD3))
            If Len(dom) > 0 Then
                If dicOut.Exists(dom) Then dicOut(dom) = CDbl(dicOut(dom)) + resid Else dicOut.Add dom, resid
            End If
        End If
NextAgg:
    Next kAgg

    Set SumarKMRealizadosPorCliente_TLPrimero = dicOut
End Function

Public Function SumarKMCotizadosPorClienteDesdeCotizacion(ByVal wsCot As Worksheet) As Object
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare

    Dim cCli As Long: cCli = ColIndexExact(wsCot, "C_Cliente")
    If cCli = 0 Then
        ' Si el encabezado tuviera variaciones:
        cCli = FindColAnyInRow(wsCot, 1, Array("C_Cliente", "Cliente"))
        If cCli = 0 Then Set SumarKMCotizadosPorClienteDesdeCotizacion = dic: Exit Function
    End If

    Dim lr As Long: lr = wsCot.Cells(wsCot.rows.count, cCli).End(xlUp).row
    Dim r As Long, cli As String, c As Long, s As Double

    For r = 2 To lr
        cli = Trim$(CStr(wsCot.Cells(r, cCli).value))
        If Len(cli) = 0 Then GoTo NextR

        s = 0
        ' Columnas W..AA = 23..27
        For c = 23 To 27
            s = s + KmsToDouble(wsCot.Cells(r, c).value)
        Next c

        If s <> 0 Then
            If dic.Exists(cli) Then
                dic(cli) = CDbl(dic(cli)) + s
            Else
                dic.Add cli, CDbl(s)
            End If
        End If
NextR:
    Next r

    Set SumarKMCotizadosPorClienteDesdeCotizacion = dic
End Function

Public Function SumarKMRealizadosDesdeBDPorCliente(ByVal wsBD As Worksheet, _
                                                    ByRef dDiasVistos As Object, _
                                                    ByRef dicLleg As Object, _
                                                    Optional ByRef dicAgg As Object = Nothing) As Object
    Dim dicCli As Object: Set dicCli = CreateObject("Scripting.Dictionary")
    dicCli.CompareMode = vbTextCompare

    Dim durVehDiaTot As Object: Set durVehDiaTot = CreateObject("Scripting.Dictionary") ' veh|dd -> seg
    Dim durVehDiaCli As Object: Set durVehDiaCli = CreateObject("Scripting.Dictionary") ' veh|dd|cli -> seg
    Dim divVehDia As Object: Set divVehDia = CreateObject("Scripting.Dictionary")       ' veh|dd -> División
    Dim coveredVD As Object: Set coveredVD = CreateObject("Scripting.Dictionary")       ' veh|dd asignados vía BD

    If wsBD Is Nothing Then On Error Resume Next: Set wsBD = ThisWorkbook.Worksheets("Base de datos"): On Error GoTo 0
    If wsBD Is Nothing Then Set SumarKMRealizadosDesdeBDPorCliente = dicCli: Exit Function

    Dim hdr As Long: hdr = FindHeaderRowBD(wsBD)

    Dim cVeh As Long, cCli As Long, cDiv As Long, cFS As Long, cFFS As Long, cH1 As Long, cH2 As Long
    cVeh = FindColAnyInRow(wsBD, hdr, Array("K_Carro", "Unidad", "Carro", "Vehiculo", "Vehículo"))
    cCli = FindColAnyInRow(wsBD, hdr, Array("C_Cliente", "Cliente"))
    cDiv = FindColAnyInRow(wsBD, hdr, Array("Division", "División", "Div"))
    cFS = FindColAnyInRow(wsBD, hdr, Array("fecha_inicial", "Fecha Inicio", "F Servicio", "Fecha", "F_Servicio", "FServicio"))
    cFFS = FindColAnyInRow(wsBD, hdr, Array("fecha_final", "Fecha Fin", "F FServicio", "F_FServicio", "FFServicio", "Fecha"))
    If cFFS = 0 Then cFFS = cFS
    cH1 = FindColAnyInRow(wsBD, hdr, Array("hora_inicial", "Hora Inicio", "HHMM1", "HoraInicial", "Inicio", "HI"))
    cH2 = FindColAnyInRow(wsBD, hdr, Array("hora_final", "Hora Fin", "HHMM2", "HoraFinal", "Fin", "HF"))

    ' Si la BD no es utilizable, cae al fallback directo por dicAgg
    If cVeh * cFS * cFFS * cH1 * cH2 = 0 Then GoTo FALLBACK_AGG

    Dim lr As Long: lr = wsBD.Cells(wsBD.rows.count, cVeh).End(xlUp).row
    If lr <= hdr Then GoTo FALLBACK_AGG

    Dim mapVD As Object: Set mapVD = BuildVehDayClientMap(dicLleg, dDiasVistos)

    Dim r As Long, veh As String, cli As String, divn As String
    Dim d1 As Double, d2 As Double, h1 As Long, h2 As Long
    Dim diaIni As Long, diaFin As Long, dd As Long
    Dim segIni As Long, segFin As Long, durD As Long
    Dim keyVD As String, kTot As String, kCliKey As String

    For r = hdr + 1 To lr
        veh = Trim$(CStr(wsBD.Cells(r, cVeh).value))
        If Len(veh) = 0 Then GoTo NextR

        If cCli > 0 Then cli = Trim$(CStr(wsBD.Cells(r, cCli).value)) Else cli = ""
        If cDiv > 0 Then divn = Trim$(CStr(wsBD.Cells(r, cDiv).value)) Else divn = ""

        d1 = DateOnlyEx2(wsBD.Cells(r, cFS).value, "DMY")
        d2 = DateOnlyEx2(wsBD.Cells(r, cFFS).value, "DMY")
        If d1 = 0 And d2 > 0 Then d1 = d2
        If d2 = 0 And d1 > 0 Then d2 = d1
        If d1 = 0 Then GoTo NextR

        h1 = TimeToSecEx(wsBD.Cells(r, cH1).value)
        h2 = TimeToSecEx(wsBD.Cells(r, cH2).value)
        If h2 < h1 And d2 = d1 Then d2 = d1 + 1

        diaIni = CLng(d1): diaFin = CLng(d2)
        If diaFin < diaIni Then diaFin = diaIni

        For dd = diaIni To diaFin
            If Not (dDiasVistos Is Nothing) Then If Not dDiasVistos.Exists(CStr(dd)) Then GoTo NextDD

            If dd = diaIni Then segIni = h1 Else segIni = 0
            If dd = diaFin Then segFin = h2 Else segFin = 86400
            If segFin < segIni Then segFin = segIni
            durD = segFin - segIni
            If durD <= 0 Then GoTo NextDD

            If Len(cli) = 0 Then
                keyVD = veh & "|" & CStr(dd)
                If mapVD.Exists(keyVD) Then cli = CStr(mapVD(keyVD)) Else GoTo NextDD
            End If

            If Len(divn) > 0 Then If Not divVehDia.Exists(veh & "|" & CStr(dd)) Then divVehDia.Add veh & "|" & CStr(dd), divn

            kTot = veh & "|" & CStr(dd)
            kCliKey = kTot & "|" & cli

            If durVehDiaTot.Exists(kTot) Then durVehDiaTot(kTot) = CLng(durVehDiaTot(kTot)) + durD Else durVehDiaTot.Add kTot, CLng(durD)
            If durVehDiaCli.Exists(kCliKey) Then durVehDiaCli(kCliKey) = CLng(durVehDiaCli(kCliKey)) + durD Else durVehDiaCli.Add kCliKey, CLng(durD)
NextDD:
        Next dd
NextR:
    Next r

    ' Reparto proporcional con dicAgg por cada veh|dd que SI tenemos en BD
    If Not (dicAgg Is Nothing) And durVehDiaTot.count > 0 Then
        Dim kVehDia As Variant, parts() As String, vehVD As String
        Dim ddOnly As Double, divVD As String, keyAgg As String, vAgg As Variant
        Dim totDur As Double, share As Double, kmTot As Double, kmMatch As Double
        Dim kCli As Variant, pref As String, cliName As String, vCliArr As Variant
        Dim kAgg As Variant, arrK() As String

        For Each kVehDia In durVehDiaTot.keys
            parts = Split(CStr(kVehDia), "|") ' veh|dd
            vehVD = CStr(parts(0))
            ddOnly = CDbl(parts(1))

            If divVehDia.Exists(CStr(kVehDia)) Then divVD = CStr(divVehDia(CStr(kVehDia))) Else divVD = ""

            If Len(divVD) = 0 Then
                ' inferir división desde dicAgg si la BD no la trae
                For Each kAgg In dicAgg.keys
                    arrK = Split(CStr(kAgg), "|") ' div|veh|dd
                    If UBound(arrK) = 2 Then
                        If arrK(1) = vehVD And CDbl(arrK(2)) = ddOnly Then divVD = arrK(0): Exit For
                    End If
                Next kAgg
            End If
            If Len(divVD) = 0 Then GoTo NextVD

            keyAgg = divVD & "|" & vehVD & "|" & CStr(ddOnly)
            If Not dicAgg.Exists(keyAgg) Then GoTo NextVD

            vAgg = dicAgg(keyAgg): kmTot = CDbl(vAgg(0)): kmMatch = CDbl(vAgg(1))
            totDur = CDbl(durVehDiaTot(CStr(kVehDia)))
            If totDur <= 0 Or kmTot <= 0 Then GoTo NextVD

            pref = CStr(kVehDia) & "|" ' veh|dd|
            For Each kCli In durVehDiaCli.keys
                If left$(CStr(kCli), Len(pref)) = pref Then
                    share = CDbl(durVehDiaCli(kCli)) / totDur
                    cliName = mid$(CStr(kCli), Len(pref) + 1)
                    If dicCli.Exists(cliName) Then vCliArr = dicCli(cliName) Else vCliArr = Array(0#, 0#)
                    vCliArr(0) = CDbl(vCliArr(0)) + kmTot * share
                    vCliArr(1) = CDbl(vCliArr(1)) + kmMatch * share
                    dicCli(cliName) = vCliArr
                End If
            Next kCli

            coveredVD(kVehDia) = True
NextVD:
        Next kVehDia
    End If

    ' *** COMPLETAR LO QUE LA BD NO CUBRIÓ: asigna 100% por cliente dominante (mapVD) ***
    If Not (dicAgg Is Nothing) Then
        Dim mapVD2 As Object: Set mapVD2 = BuildVehDayClientMap(dicLleg, dDiasVistos)
        Dim k As Variant, p() As String, vehF As String, ddF As Long, cliF As String
        Dim v As Variant, vCli As Variant, vdKey As String

        For Each k In dicAgg.keys
            p = Split(CStr(k), "|") ' div|veh|dd
            If UBound(p) = 2 Then
                vehF = p(1): ddF = CLng(CDbl(p(2)))
                If (dDiasVistos Is Nothing) Or dDiasVistos.Exists(CStr(ddF)) Then
                    vdKey = vehF & "|" & CStr(ddF)
                    If Not coveredVD.Exists(vdKey) Then
                        If mapVD2.Exists(vdKey) Then
                            cliF = CStr(mapVD2(vdKey))
                            v = dicAgg(k)
                            If dicCli.Exists(cliF) Then vCli = dicCli(cliF) Else vCli = Array(0#, 0#)
                            vCli(0) = CDbl(vCli(0)) + CDbl(v(0)) ' km Tot
                            vCli(1) = CDbl(vCli(1)) + CDbl(v(1)) ' km Match
                            dicCli(cliF) = vCli
                        End If
                    End If
                End If
            End If
        Next k
    End If

    Set SumarKMRealizadosDesdeBDPorCliente = dicCli
    Exit Function

FALLBACK_AGG:
    ' Fallback completo: asigna 100% por cliente dominante por día (si no pudimos usar BD)
    If Not (dicAgg Is Nothing) Then
        Dim mapVD3 As Object: Set mapVD3 = BuildVehDayClientMap(dicLleg, dDiasVistos)
        Dim kk As Variant, PP() As String, vehX As String, ddX As Long, cliX As String
        Dim vv As Variant, vvCli As Variant
        For Each kk In dicAgg.keys
            PP = Split(CStr(kk), "|")
            If UBound(PP) = 2 Then
                vehX = PP(1): ddX = CLng(CDbl(PP(2)))
                If (dDiasVistos Is Nothing) Or dDiasVistos.Exists(CStr(ddX)) Then
                    keyVD = vehX & "|" & CStr(ddX)
                    If mapVD3.Exists(keyVD) Then
                        cliX = CStr(mapVD3(keyVD))
                        vv = dicAgg(kk)
                        If dicCli.Exists(cliX) Then vvCli = dicCli(cliX) Else vvCli = Array(0#, 0#)
                        vvCli(0) = CDbl(vvCli(0)) + CDbl(vv(0))
                        vvCli(1) = CDbl(vvCli(1)) + CDbl(vv(1))
                        dicCli(cliX) = vvCli
                    End If
                End If
            End If
        Next kk
    End If

    Set SumarKMRealizadosDesdeBDPorCliente = dicCli
End Function

Public Function BuildVehDayClientMap(ByRef dLleg As Object, ByRef dDiasVistos As Object) As Object
    Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
    Dim best As Object: Set best = CreateObject("Scripting.Dictionary") ' veh|dd -> bestAbs
    Dim k As Variant, col As Collection, i As Long
    Dim w As Variant, key As String, dd As Long, absLleg As Double, cli As String
    If dLleg Is Nothing Then Set BuildVehDayClientMap = map: Exit Function

    For Each k In dLleg.keys
        Set col = dLleg(k)
        For i = 1 To col.count
            w = col(i)
            absLleg = CDbl(w(0))
            cli = CStr(w(1))
            dd = CLng(w(6))
            If (dDiasVistos Is Nothing) Or dDiasVistos.Exists(CStr(dd)) Then
                key = CStr(k) & "|" & CStr(dd)
                If Not best.Exists(key) Then
                    best.Add key, absLleg
                    map.Add key, cli
                ElseIf absLleg < CDbl(best(key)) Then
                    best(key) = absLleg
                    map(key) = cli
                End If
            End If
        Next i
    Next k
    Set BuildVehDayClientMap = map
End Function

Public Function SumarKMMatchDesdeTLPorCliente(ByVal wsTL As Worksheet) As Object
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare

    If wsTL Is Nothing Then On Error Resume Next: Set wsTL = ThisWorkbook.Worksheets(HOJA_TL): On Error GoTo 0
    If wsTL Is Nothing Then Set SumarKMMatchDesdeTLPorCliente = dic: Exit Function

    Dim cTipo As Long, cKm As Long, cCli As Long
    cTipo = ColIndexExact(wsTL, "Tipo")
    cKm = ColIndexExact(wsTL, "Km")
    cCli = GetClienteSiteVisitColumn(wsTL)   ' <- clave correcta en tu TL

    Dim missingTLCols As String
    If cTipo = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Tipo"
    If cKm = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Km"
    If cCli = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Cliente / SiteVisit"
    If Len(missingTLCols) > 0 Then
        WarnMissingColumns wsTL.name, missingTLCols
        Set SumarKMMatchDesdeTLPorCliente = dic
        Exit Function
    End If

    Dim lr As Long, r As Long, tipo As String, cli As String, km As Double
    lr = wsTL.Cells(wsTL.rows.count, cTipo).End(xlUp).row

    For r = 2 To lr
        tipo = LCase$(Trim$(CStr(wsTL.Cells(r, cTipo).value)))
        If tipo = "match" Then
            cli = Trim$(CStr(wsTL.Cells(r, cCli).value))
            If Len(cli) > 0 Then
                If IsNumeric(wsTL.Cells(r, cKm).value) Then
                    km = CDbl(wsTL.Cells(r, cKm).value)
                    If km > 0 Then
                        If dic.Exists(cli) Then
                            dic(cli) = CDbl(dic(cli)) + km
                        Else
                            dic.Add cli, km
                        End If
                    End If
                End If
            End If
        End If
    Next r

    Set SumarKMMatchDesdeTLPorCliente = dic
End Function

Public Sub CrearResumenGlobal(ByVal wb As Workbook, ByRef dicTot As Object, ByRef dicKVcats As Object)
    Dim wsG As Worksheet
    On Error Resume Next
    Set wsG = wb.Worksheets(HOJA_RES_GLOBAL)
    On Error GoTo 0
    If wsG Is Nothing Then
        Set wsG = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.count))
        On Error Resume Next: wsG.name = HOJA_RES_GLOBAL: On Error GoTo 0
    Else
        wsG.Cells.Clear
    End If

    Dim totalKm As Double, totalMatch As Double, totalVacios As Double
    Dim gP As Double, gD As Double, gT As Double, gO As Double
    Dim k As Variant, v As Variant

    If Not dicTot Is Nothing Then
        For Each k In dicTot.keys
            v = dicTot(k)
            totalKm = totalKm + CDbl(v(0))
            totalMatch = totalMatch + CDbl(v(1))
        Next k
    End If
    totalVacios = Application.Max(0, totalKm - totalMatch)

    If Not dicKVcats Is Nothing Then
        For Each k In dicKVcats.keys
            v = dicKVcats(k)
            gP = gP + CDbl(v(0))
            gD = gD + CDbl(v(1))
            gT = gT + CDbl(v(2))
            gO = gO + CDbl(v(3))
        Next k
    End If

    Dim sKG As Double, diffG As Double
    sKG = gP + gD + gT + gO
    diffG = totalVacios - sKG
    If Abs(diffG) > 0.000001 Then gO = gO + diffG

    wsG.Range("A1:B1").value = Array("Tipo de Km", "Valor")
    wsG.Range("A2").value = "Match":  wsG.Range("B2").value = totalMatch
    wsG.Range("A3").value = "Patio":  wsG.Range("B3").value = gP
    wsG.Range("A4").value = "Taller": wsG.Range("B4").value = gT
    wsG.Range("A5").value = "Diesel": wsG.Range("B5").value = gD
    wsG.Range("A6").value = "Otros":  wsG.Range("B6").value = gO

    wsG.Columns("A:B").AutoFit

    Dim lo As ListObject
    On Error Resume Next
    wsG.ListObjects("tblResumenGlobal").Unlist
    On Error GoTo 0
    Set lo = wsG.ListObjects.Add(xlSrcRange, wsG.Range("A1:B6"), , xlYes)
    lo.name = "tblResumenGlobal"
    On Error Resume Next
    lo.TableStyle = "TableStyleMedium2"
    On Error GoTo 0
End Sub

Public Function SumarKMCliente_TLxBD(ByVal wsTL As Worksheet, ByVal wsBD As Worksheet, _
                                      ByRef dDiasVistos As Object, ByRef dLleg As Object, _
                                      Optional ByRef dicSoporte As Object) As Object
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary"): dic.CompareMode = vbTextCompare
    If wsTL Is Nothing Then On Error Resume Next: Set wsTL = ThisWorkbook.Worksheets(HOJA_TL): On Error GoTo 0
    If wsTL Is Nothing Then Set SumarKMCliente_TLxBD = dic: Exit Function

    If dicSoporte Is Nothing Then
        Set dicSoporte = CreateObject("Scripting.Dictionary")
        dicSoporte.CompareMode = vbTextCompare
    End If

    Dim cVeh As Long, cTipo As Long, cIni As Long, cFin As Long, cKm As Long, cCliTL As Long
    cVeh = ColIndexExact(wsTL, "Vehiculo")
    cTipo = ColIndexExact(wsTL, "Tipo")
    cIni = ColIndexExact(wsTL, "Inicio")
    cFin = ColIndexExact(wsTL, "Fin")
    cKm = ColIndexExact(wsTL, "Km")
    cCliTL = GetClienteSiteVisitColumn(wsTL)

    Dim missingTLCols As String
    If cVeh = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Vehiculo"
    If cTipo = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Tipo"
    If cIni = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Inicio"
    If cFin = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Fin"
    If cKm = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Km"
    If cCliTL = 0 Then missingTLCols = missingTLCols & vbCrLf & " - Cliente / SiteVisit"
    If Len(missingTLCols) > 0 Then
        WarnMissingColumns wsTL.name, missingTLCols
        Set SumarKMCliente_TLxBD = dic
        Exit Function
    End If

    Dim bdWins As Object: Set bdWins = BuildBDWinsFromSheet(wsBD)

    Dim lr As Long: lr = wsTL.Cells(wsTL.rows.count, cVeh).End(xlUp).row
    Dim r As Long, veh As String, tipo As String, cliTL As String
    Dim dIni As Double, dFin As Double, km As Double
    Dim absS As Double, absE As Double, dur As Double

    For r = 2 To lr
        veh = Trim$(CStr(wsTL.Cells(r, cVeh).value)): If Len(veh) = 0 Then GoTo NextR
        tipo = LCase$(Trim$(CStr(wsTL.Cells(r, cTipo).value)))
        If Not IsDate(wsTL.Cells(r, cIni).value) Or Not IsDate(wsTL.Cells(r, cFin).value) Then GoTo NextR
        dIni = CDbl(CDate(wsTL.Cells(r, cIni).value))
        dFin = CDbl(CDate(wsTL.Cells(r, cFin).value))
        km = 0: If IsNumeric(wsTL.Cells(r, cKm).value) Then km = CDbl(wsTL.Cells(r, cKm).value)
        If km <= 0 Then GoTo NextR

        absS = dIni * 86400#
        absE = dFin * 86400#
        dur = absE - absS: If dur <= 0 Then GoTo NextR

        If tipo = "match" Then
            cliTL = Trim$(CStr(wsTL.Cells(r, cCliTL).value))
            If Len(cliTL) > 0 Then
                If dic.Exists(cliTL) Then dic(cliTL) = CDbl(dic(cliTL)) + km Else dic.Add cliTL, km
            End If
        Else
            Dim shareTot As Double: shareTot = 0
            Dim tmpCli As Object: Set tmpCli = CreateObject("Scripting.Dictionary"): tmpCli.CompareMode = vbTextCompare

            If bdWins.Exists(veh) Then
                Dim i As Long, w As Variant, os As Double, oe As Double, ol As Double, cliBD As String
                For i = 1 To bdWins(veh).count
                    w = bdWins(veh)(i) ' [absS, absE, cli]
                    os = IIf(absS > CDbl(w(0)), absS, CDbl(w(0)))
                    oe = IIf(absE < CDbl(w(1)), absE, CDbl(w(1)))
                    ol = oe - os
                    If ol > 0 Then
                        cliBD = Trim$(CStr(w(2)))
                        If Len(cliBD) > 0 Then
                            If tmpCli.Exists(cliBD) Then tmpCli(cliBD) = CDbl(tmpCli(cliBD)) + ol Else tmpCli.Add cliBD, ol
                            shareTot = shareTot + ol
                        End If
                    End If
                Next i
            End If

            If shareTot > 0 Then
                Dim k As Variant, frac As Double
                For Each k In tmpCli.keys
                    frac = CDbl(tmpCli(k)) / shareTot
                    If dic.Exists(CStr(k)) Then dic(CStr(k)) = CDbl(dic(CStr(k))) + km * frac Else dic.Add CStr(k), km * frac
                    If Not dicSoporte.Exists(CStr(k)) Then dicSoporte.Add CStr(k), "BD"
                Next k
            Else
                Dim dd As Long: dd = CLng(Int(dIni))
                Dim cliDom As String: cliDom = ClienteForVehDia(veh, dd)
                If Len(cliDom) > 0 Then
                    If dic.Exists(cliDom) Then dic(cliDom) = CDbl(dic(cliDom)) + km Else dic.Add cliDom, km
                    If Not dicSoporte.Exists(cliDom) Then dicSoporte.Add cliDom, "LLEGADAS"
                End If
            End If
        End If
NextR:
    Next r

    Set SumarKMCliente_TLxBD = dic
End Function

Public Function BuildBDWinsFromSheet(ByVal wsBD As Worksheet) As Object
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare
    If wsBD Is Nothing Then Set BuildBDWinsFromSheet = dic: Exit Function

    Dim hdr As Long: hdr = FindHeaderRowBD(wsBD)

    Dim cVeh As Long, cCli As Long, cFS As Long, cFF As Long, cH1 As Long, cH2 As Long
    cVeh = FindColAnyInRow(wsBD, hdr, Array("K_Carro", "Unidad", "Carro", "Vehiculo", "Vehículo"))
    cCli = FindColAnyInRow(wsBD, hdr, Array("C_Cliente", "Cliente"))
    cFS = FindColAnyInRow(wsBD, hdr, Array("fecha_inicial", "Fecha Inicio", "F Servicio", "Fecha", "F_Servicio", "FServicio"))
    cFF = FindColAnyInRow(wsBD, hdr, Array("fecha_final", "Fecha Fin", "F FServicio", "F_FServicio", "FFServicio", "Fecha")): If cFF = 0 Then cFF = cFS
    cH1 = FindColAnyInRow(wsBD, hdr, Array("hora_inicial", "Hora Inicio", "HHMM1", "HoraInicial", "Inicio", "HI"))
    cH2 = FindColAnyInRow(wsBD, hdr, Array("hora_final", "Hora Fin", "HHMM2", "HoraFinal", "Fin", "HF"))

    If cVeh * cFS * cH1 * cH2 = 0 Then Set BuildBDWinsFromSheet = dic: Exit Function

    Dim lr As Long: lr = wsBD.Cells(wsBD.rows.count, cVeh).End(xlUp).row
    Dim r As Long, veh As String, cli As String
    Dim d1 As Double, d2 As Double, h1 As Long, h2 As Long
    Dim absS As Double, absE As Double
    Dim col As Collection

    For r = hdr + 1 To lr
        veh = Trim$(CStr(wsBD.Cells(r, cVeh).value))
        If Len(veh) = 0 Then GoTo NextR

        If cCli > 0 Then cli = Trim$(CStr(wsBD.Cells(r, cCli).value)) Else cli = ""

        d1 = DateOnlyEx2(wsBD.Cells(r, cFS).value, "DMY")
        d2 = DateOnlyEx2(wsBD.Cells(r, cFF).value, "DMY")
        If d1 = 0 And d2 > 0 Then d1 = d2
        If d2 = 0 And d1 > 0 Then d2 = d1
        If d1 = 0 Then GoTo NextR

        h1 = TimeToSecEx(wsBD.Cells(r, cH1).value)
        h2 = TimeToSecEx(wsBD.Cells(r, cH2).value)
        If h2 < h1 And d2 = d1 Then d2 = d1 + 1

        absS = d1 * 86400# + h1
        absE = d2 * 86400# + h2
        If absE <= absS Then GoTo NextR

        If dic.Exists(veh) Then
            Set col = dic(veh)
        Else
            Set col = New Collection
            dic.Add veh, col
        End If
        col.Add Array(absS, absE, cli)
NextR:
    Next r

    Set BuildBDWinsFromSheet = dic
End Function


