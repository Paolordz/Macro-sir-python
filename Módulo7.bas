Attribute VB_Name = "Módulo7"
Option Explicit

Public Const TOLERANCIA_MIN As Long = 0
Public Const TL_GAP_MIN As Long = 0

Public Sub RebalanceLineaTiempoContraAgg(ByVal wsTL As Worksheet, ByRef dicAgg As Object)
    Dim lr As Long, r As Long
    Dim key As String, div_ As String, veh_ As String
    Dim dd As Double, expKm As Double, gotKm As Double, diff As Double
    Dim rowsByKey As Object, sums As Object
    Dim vals As Variant, k As Variant
    Dim lastRowIdx As Long, kmCell As Range, newVal As Double
    Dim TOL As Double: TOL = 0.0005 ' ~0.5 m

    lr = wsTL.Cells(wsTL.rows.count, 1).End(xlUp).row
    If lr < 2 Then Exit Sub

    Set rowsByKey = CreateObject("Scripting.Dictionary")
    Set sums = CreateObject("Scripting.Dictionary")

    ' Indexar filas por Division|Vehiculo|Fecha(dd de Inicio)
    For r = 2 To lr
        If Len(wsTL.Cells(r, 1).value) = 0 Then GoTo NextR
        div_ = CStr(wsTL.Cells(r, 1).value)
        veh_ = CStr(wsTL.Cells(r, 2).value)
        If IsDate(wsTL.Cells(r, 4).value) Then
            dd = Int(CDbl(wsTL.Cells(r, 4).value))
        Else
            GoTo NextR
        End If
        key = div_ & "|" & veh_ & "|" & CStr(dd)

        If Not rowsByKey.Exists(key) Then rowsByKey.Add key, New Collection
        rowsByKey(key).Add r

        gotKm = 0
        If IsNumeric(wsTL.Cells(r, 6).value) Then gotKm = CDbl(wsTL.Cells(r, 6).value)
        If sums.Exists(key) Then sums(key) = CDbl(sums(key)) + gotKm Else sums.Add key, gotKm
NextR:
    Next r

    ' Ajustar a KmTot diarios (dicAgg(k).vals(0))
    For Each k In rowsByKey.keys
        If Not dicAgg.Exists(k) Then GoTo NextK
        vals = dicAgg(k)                 ' vals(0)=KmTot
        expKm = 0
        On Error Resume Next: expKm = CDbl(vals(0)): On Error GoTo 0
        gotKm = 0
        If sums.Exists(k) Then gotKm = CDbl(sums(k))
        diff = expKm - gotKm
        If Abs(diff) <= TOL Then GoTo NextK

        lastRowIdx = rowsByKey(k)(rowsByKey(k).count)
        Set kmCell = wsTL.Cells(lastRowIdx, 6)
        newVal = CDbl(IIf(IsNumeric(kmCell.value), kmCell.value, 0)) + diff

        If newVal < 0 Then
            Call RedistribuirProporcional(wsTL, rowsByKey(k), expKm)
        Else
            kmCell.value = newVal
        End If
NextK:
    Next k
End Sub

Private Sub RedistribuirProporcional(ByVal wsTL As Worksheet, ByVal rows As Collection, ByVal objetivo As Double)
    Dim i As Long, suma As Double, v() As Double, r As Long
    Dim factor As Double, resid As Double

    ReDim v(1 To rows.count)
    For i = 1 To rows.count
        r = rows(i)
        v(i) = CDbl(IIf(IsNumeric(wsTL.Cells(r, 6).value), wsTL.Cells(r, 6).value, 0))
        suma = suma + v(i)
    Next i

    If suma <= 0 Then
        wsTL.Cells(rows(rows.count), 6).value = objetivo
        Exit Sub
    End If

    factor = objetivo / suma
    resid = objetivo
    For i = 1 To rows.count - 1
        v(i) = v(i) * factor
        wsTL.Cells(rows(i), 6).value = v(i)
        resid = resid - v(i)
    Next i
    wsTL.Cells(rows(rows.count), 6).value = resid
End Sub

Public Sub RebalanceKMVaciosContraAgg(ByVal wsKM As Worksheet, ByRef dicAgg As Object)
    Dim lr As Long, r As Long
    Dim sumKM As Object, lastRowByKey As Object
    Dim key As String, div_ As String, veh_ As String, dd As Double
    Dim arrAgg As Variant, expVacios As Double, gotVacios As Double, diff As Double
    Dim k As Variant
    Const TOL As Double = 0.0005

    lr = wsKM.Cells(wsKM.rows.count, 1).End(xlUp).row
    If lr < 2 Then Exit Sub

    Set sumKM = CreateObject("Scripting.Dictionary")
    Set lastRowByKey = CreateObject("Scripting.Dictionary")

    For r = 2 To lr
        If Len(wsKM.Cells(r, 1).value) = 0 Then GoTo NextR
        div_ = CStr(wsKM.Cells(r, 1).value)
        veh_ = CStr(wsKM.Cells(r, 2).value)
        If IsDate(wsKM.Cells(r, 3).value) Then
            dd = Int(CDbl(wsKM.Cells(r, 3).value))
        Else
            GoTo NextR
        End If
        key = div_ & "|" & veh_ & "|" & CStr(dd)
        If Not sumKM.Exists(key) Then sumKM.Add key, 0#
        If Not lastRowByKey.Exists(key) Then lastRowByKey.Add key, r Else lastRowByKey(key) = r
        If IsNumeric(wsKM.Cells(r, 6).value) Then
            sumKM(key) = CDbl(sumKM(key)) + CDbl(wsKM.Cells(r, 6).value)
        End If
NextR:
    Next r

    For Each k In sumKM.keys
        If dicAgg.Exists(k) Then
            arrAgg = dicAgg(k)                       ' (0)=KmTot, (1)=KmMatch
            expVacios = Application.Max(0, CDbl(arrAgg(0)) - CDbl(arrAgg(1)))
            gotVacios = CDbl(sumKM(k))
            diff = expVacios - gotVacios
            If Abs(diff) > TOL Then
                r = CLng(lastRowByKey(k))
                wsKM.Cells(r, 6).value = CDbl(IIf(IsNumeric(wsKM.Cells(r, 6).value), wsKM.Cells(r, 6).value, 0)) + diff
                If wsKM.Cells(r, 6).value < 0 Then wsKM.Cells(r, 6).value = 0
            End If
        End If
    Next k
End Sub

Public Function FindArrivalDayForMovement(ByVal veh As String, ByVal absStart As Double, ByVal absEnd As Double, ByRef dayOut As Long) As Boolean
    Dim col As Collection, i As Long, w As Variant, t As Double, best As Double, found As Boolean
    dayOut = -1
    If dicLlegadas Is Nothing Then Exit Function
    If Not dicLlegadas.Exists(veh) Then Exit Function

    Set col = dicLlegadas(veh)
    best = 1E+99: found = False
    For i = 1 To col.count
        w = col(i)
        t = CDbl(w(0))
        If t >= absStart And t <= absEnd Then
            If t < best Then best = t: found = True
        End If
    Next i
    If found Then
        dayOut = CLng(Int(best / 86400#))
        FindArrivalDayForMovement = True
    End If
End Function

Public Sub BuildVehEndTimes(ByRef dicAut As Object, ByRef dicVehEnd As Object)
    Dim k As Variant, arr() As String, veh As String, d As Double, col As Collection, i As Long
    Dim ar As Variant, finA As Long, absEnd As Double
    For Each k In dicAut.keys
        arr = Split(CStr(k), "|")
        veh = arr(0)
        d = CDbl(arr(1))
        Set col = dicAut(k)
        For i = 1 To col.count
            ar = col(i)
            finA = CLng(ar(2))
            absEnd = d * 86400 + finA
            If Not dicVehEnd.Exists(veh) Then dicVehEnd.Add veh, New Collection
            dicVehEnd(veh).Add absEnd
        Next i
    Next k
End Sub

Public Sub LoadVisitasWindowsFromSheet(ByVal wsV As Worksheet, ByRef dicVehEnds As Object, ByRef dicWins As Object)
    Dim hdr As Long
    Dim cUnidad As Long, cFL As Long, cHL As Long, cFSal As Long, cHSal As Long, cDur As Long, cCat As Long, cSitio As Long
    Dim lastRow As Long, r As Long
    Dim veh As String, dLleg As Double, hLleg As Long, dSal As Double, hSal As Long, absLleg As Double, absSal As Double
    Dim durSec As Long, catRaw As String, cat As String, sitio As String

    ' Detecta encabezado; si no, forzar D..H con datos desde fila 8
    hdr = FindHeaderRowVisitas(wsV, 8)
    cUnidad = FindColAnyInRow(wsV, hdr, Array("Unidad", "Económico", "Economico", "No Economico", "NoEconomico"))
    cFL = FindColAnyInRow(wsV, hdr, Array("Fecha Llegada", "FechaLlegada", "Fecha Arribo", "Fecha", "F Llegada"))
    cHL = FindColAnyInRow(wsV, hdr, Array("Hora Llegada", "HoraLlegada", "Hora Arribo", "Hora"))
    cFSal = FindColAnyInRow(wsV, hdr, Array("Fecha Salida", "FechaSalida", "F Salida", "Fecha Fin"))
    cHSal = FindColAnyInRow(wsV, hdr, Array("Hora Salida", "HoraSalida", "H Salida", "Hora Fin"))
    cDur = FindColAnyInRow(wsV, hdr, Array("Tiempo de Visita", "TiempoVisita", "Duración", "Duracion"))
    cCat = FindColAnyInRow(wsV, hdr, Array("Categoría", "Categoria", "Categoría Visita", "Categoria Visita", "Tipo", "Tipo Visita"))
    cSitio = FindColAnyInRow(wsV, hdr, Array("Sitio", "Lugar", "Ubicacion", "Ubicación", "Punto", "Destino"))

    If cUnidad = 0 And cFL = 0 And cHL = 0 And cFSal = 0 And cHSal = 0 And cDur = 0 Then
        hdr = 7: cFL = 4: cHL = 5: cFSal = 6: cHSal = 7: cDur = 8
        If cUnidad = 0 Then cUnidad = 1 ' si no hay, toma col 1 como unidad
    End If
    If cUnidad = 0 Or cFL = 0 Or cHL = 0 Then
        MsgBox "Visitas: faltan columnas mínimas (Unidad, Fecha/Hora Llegada).", vbExclamation
        Exit Sub
    End If

    lastRow = wsV.Cells(wsV.rows.count, cFL).End(xlUp).row
    For r = hdr + 1 To lastRow
        veh = Trim$(CStr(wsV.Cells(r, cUnidad).value))
        If Len(veh) = 0 Then GoTo NextR

        dLleg = DateOnlyEx2(wsV.Cells(r, cFL).value, "DMY")
        hLleg = TimeToSecEx(wsV.Cells(r, cHL).value)
        If dLleg = 0 Then GoTo NextR
        absLleg = dLleg * 86400# + hLleg

        ' Salida y/o duración
        dSal = 0: hSal = 0: durSec = 0
        If cFSal > 0 Then dSal = DateOnlyEx2(wsV.Cells(r, cFSal).value, "DMY")
        If cHSal > 0 Then hSal = TimeToSecEx(wsV.Cells(r, cHSal).value)
        If cDur > 0 Then durSec = TimeToSecEx(wsV.Cells(r, cDur).value)

        If dSal > 0 Or hSal > 0 Then
            If dSal = 0 Then dSal = dLleg
            absSal = dSal * 86400# + hSal
            If absSal < absLleg Then absSal = absLleg
        ElseIf durSec > 0 Then
            absSal = absLleg + durSec
        Else
            absSal = absLleg ' sin salida ni duración
        End If

        ' Duración efectiva y filtro < 60 s
        If absSal - absLleg < 60# Then GoTo NextR

        ' Categoría y sitio
        If cCat > 0 Then catRaw = CStr(wsV.Cells(r, cCat).value) Else catRaw = ""
        cat = CatNormalize(catRaw)
        If cSitio > 0 Then sitio = CStr(wsV.Cells(r, cSitio).value) Else sitio = ""
        If cat = "Otros" Or cat = "" Then
            Dim cat2 As String: cat2 = GuessCatFromSite(sitio)
            If cat2 <> "Otros" Then cat = cat2
        End If

        If Not dicWins.Exists(veh) Then dicWins.Add veh, New Collection
        dicWins(veh).Add Array(absLleg, absSal, cat, sitio)
NextR:
    Next r
End Sub

Public Sub LoadVisitasWindows(ByVal visitasPath As String, ByRef dicVehEnds As Object, ByRef dicWins As Object)
    Dim wbV As Workbook, wsV As Worksheet
    Dim hdr As Long
    Dim cUnidad As Long, cFL As Long, cHL As Long, cCat As Long, cSitio As Long
    Dim lastRow As Long, r As Long
    Dim veh As String
    Dim dLleg As Double, hLleg As Long
    Dim absLleg As Double
    Dim catRaw As String, cat As String
    Dim sitio As String
    Dim startAbs As Double
    Dim col As Collection
    Dim mx As Double, j As Long, e As Double

    Set wbV = Workbooks.Open(fileName:=visitasPath, ReadOnly:=True)
    On Error Resume Next
    Set wsV = wbV.Worksheets("Visitas")
    On Error GoTo 0
    If wsV Is Nothing Then Set wsV = wbV.Worksheets(1)

    hdr = FindHeaderRowVisitas(wsV, 8)

    cUnidad = FindColAnyInRow(wsV, hdr, Array("Unidad", "Económico", "Economico", "No Economico", "NoEconomico"))
    cFL = FindColAnyInRow(wsV, hdr, Array("Fecha Llegada", "FechaLlegada", "Fecha Arribo", "Fecha", "F Llegada"))
    cHL = FindColAnyInRow(wsV, hdr, Array("Hora Llegada", "HoraLlegada", "Hora Arribo", "Hora"))
    cCat = FindColAnyInRow(wsV, hdr, Array("Categoría", "Categoria", "Categoría Visita", "Categoria Visita", "Tipo", "Tipo Visita"))
    cSitio = FindColAnyInRow(wsV, hdr, Array("Sitio", "Lugar", "Ubicacion", "Ubicación", "Punto", "Destino"))

    If cUnidad = 0 Or cFL = 0 Or cHL = 0 Or cCat = 0 Then
        wbV.Close False
        MsgBox "No se hallaron columnas mínimas en el archivo de Visitas (Unidad, Fecha Llegada, Hora Llegada, Categoría).", vbExclamation
        Exit Sub
    End If

    lastRow = wsV.Cells(wsV.rows.count, cUnidad).End(xlUp).row
    For r = hdr + 1 To lastRow
        veh = Trim$(CStr(wsV.Cells(r, cUnidad).value))
        If veh = "" Then GoTo NextR

        dLleg = DateOnlyEx2(wsV.Cells(r, cFL).value, "DMY")
        hLleg = TimeToSecEx(wsV.Cells(r, cHL).value)
        If dLleg = 0 Then GoTo NextR

        absLleg = dLleg * 86400 + hLleg

        catRaw = CStr(wsV.Cells(r, cCat).value)
        cat = CatNormalize(catRaw)

        If cSitio > 0 Then sitio = CStr(wsV.Cells(r, cSitio).value) Else sitio = ""

        If cat = "Otros" Or cat = "" Then
            Dim cat2 As String: cat2 = GuessCatFromSite(sitio)
            If cat2 <> "Otros" Then cat = cat2
        End If

        startAbs = -1
        If dicVehEnds.Exists(veh) Then
            Set col = dicVehEnds(veh)
            mx = -1
            For j = 1 To col.count
                e = CDbl(col(j))
                If e <= absLleg And e > mx Then mx = e
            Next j
            If mx >= 0 Then startAbs = mx
        End If
        If startAbs < 0 Then
            startAbs = absLleg - 12# * 3600#
            If startAbs < 0# Then startAbs = 0#
        End If

        If Not dicWins.Exists(veh) Then dicWins.Add veh, New Collection
        dicWins(veh).Add Array(startAbs, absLleg, cat, sitio)
NextR:
    Next r

    wbV.Close SaveChanges:=False
End Sub

Public Sub LoadCargasFromSheet(ByVal wsC As Worksheet, ByRef dic As Object)
    Dim hdr As Long, lastRow As Long, r As Long
    Dim veh As String, d As Double, divc As String

    hdr = 1
    Dim cUnidad As Long: cUnidad = 6
    Dim cFecha  As Long: cFecha = 12

    Dim cDiv As Long: cDiv = 0
    On Error Resume Next
    cDiv = FindColAnyInRow(wsC, hdr, Array("Division", "División", "Div", "Sede", "Plaza"))
    On Error GoTo 0

    lastRow = wsC.Cells(wsC.rows.count, cUnidad).End(xlUp).row
    For r = hdr + 1 To lastRow
        veh = Trim$(CStr(wsC.Cells(r, cUnidad).value))
        If Len(veh) = 0 Then GoTo NextR

        d = DateOnlyEx2(wsC.Cells(r, cFecha).value, "DMY")
        If d = 0 Then GoTo NextR

        If cDiv > 0 Then
            divc = Trim$(CStr(wsC.Cells(r, cDiv).value))
        Else
            divc = ""
        End If

        dic(veh & "|" & CStr(d)) = divc
NextR:
    Next r
End Sub

Public Function CargasDivisionFor(ByVal veh As String, ByVal dd As Double) As String
    Dim key As String: key = veh & "|" & CStr(dd)
    If Not dicCargas Is Nothing Then
        If dicCargas.Exists(key) Then
            CargasDivisionFor = CStr(dicCargas(key))
            Exit Function
        End If
    End If
    CargasDivisionFor = ""
End Function

Public Sub AttributeNoMatchToVisitCats(ByVal divisionName As String, ByVal veh As String, ByVal dd As Double, _
                                        ByVal absStart As Double, ByVal absEnd As Double, _
                                        ByVal kmsDia As Double, ByVal durD As Long, _
                                        ByVal srcTag As String, ByVal wsKM As Worksheet, _
                                        ByRef catPick As String)

    Dim overlaps() As Double, cats() As String, sites() As String, ovCount As Long
    Dim maxShare As Double, maxCat As String, maxSite As String
    Dim unionLen As Double
    Dim i As Long, k As Long, j As Long, ii As Long
    Dim s As Double, e As Double, curS As Double, curE As Double
    Dim first As Boolean
    Dim segStart As Double, segEnd As Double, ov As Double, share As Double
    Dim sumP As Double, sumD As Double, sumT As Double, sumO As Double
    Dim totalAssigned As Double, scaleFactor As Double, resid As Double
    Dim targetCat As String, key As String, vals As Variant, rM As Long
    Dim idx() As Long, tmp As Long
    Dim col As Collection, w As Variant
    Dim winS As Double, winE As Double
    Dim wCat As String, wSite As String
    Dim ovStart As Double, ovEnd As Double

    ovCount = 0
    maxShare = -1
    maxCat = "": maxSite = ""
    unionLen = 0
    sumP = 0: sumD = 0: sumT = 0: sumO = 0

    If Not dicVisitWins Is Nothing Then
        If dicVisitWins.Exists(veh) Then
            Set col = dicVisitWins(veh)
            For j = 1 To col.count
                w = col(j)
                winS = CDbl(w(0)): winE = CDbl(w(1))
                wCat = CStr(w(2))
                If UBound(w) >= 3 Then wSite = CStr(w(3)) Else wSite = ""

                ovStart = absStart: If winS > ovStart Then ovStart = winS
                ovEnd = absEnd:     If winE < ovEnd Then ovEnd = winE
                If ovEnd > ovStart Then
                    If ovCount = 0 Then
                        ReDim overlaps(0 To 1): ReDim cats(0 To 0): ReDim sites(0 To 0)
                    Else
                        ReDim Preserve overlaps(0 To 2 * ovCount + 1)
                        ReDim Preserve cats(0 To ovCount)
                        ReDim Preserve sites(0 To ovCount)
                    End If
                    overlaps(2 * ovCount) = ovStart
                    overlaps(2 * ovCount + 1) = ovEnd
                    cats(ovCount) = wCat
                    sites(ovCount) = wSite
                    ovCount = ovCount + 1
                End If
            Next j
        End If
    End If

    If ovCount > 0 Then
        ReDim idx(0 To ovCount - 1)
        For i = 0 To ovCount - 1: idx(i) = i: Next i
        For i = 0 To ovCount - 2
            For k = i + 1 To ovCount - 1
                If overlaps(2 * idx(k)) < overlaps(2 * idx(i)) Then
                    tmp = idx(i): idx(i) = idx(k): idx(k) = tmp
                End If
            Next k
        Next i

        first = True
        For i = 0 To ovCount - 1
            ii = idx(i)
            s = overlaps(2 * ii)
            e = overlaps(2 * ii + 1)
            If first Then
                curS = s: curE = e: first = False
            Else
                If s <= curE Then
                    If e > curE Then curE = e
                Else
                    unionLen = unionLen + (curE - curS)
                    curS = s: curE = e
                End If
            End If
        Next i
        If Not first Then unionLen = unionLen + (curE - curS)
    End If

    If unionLen > 0 And durD > 0 And kmsDia > 0 Then
        For i = 0 To ovCount - 1
            segStart = overlaps(2 * i): segEnd = overlaps(2 * i + 1)
            ov = segEnd - segStart: If ov < 0 Then ov = 0
            share = kmsDia * (ov / unionLen)
            Select Case cats(i)
                Case "Patio":  sumP = sumP + share
                Case "Diesel": sumD = sumD + share
                Case "Taller": sumT = sumT + share
                Case Else:     sumO = sumO + share
            End Select
            If share > maxShare Then
                maxShare = share: maxCat = cats(i): maxSite = sites(i)
            End If
        Next i
    Else
        sumO = kmsDia: maxCat = "Otros": maxSite = ""
    End If

    totalAssigned = sumP + sumD + sumT + sumO
    If maxCat <> "" Then targetCat = maxCat Else targetCat = "Otros"

    If totalAssigned < 0.000000001 Then
        sumP = 0: sumD = 0: sumT = 0: sumO = kmsDia
    Else
        scaleFactor = kmsDia / totalAssigned
        sumP = sumP * scaleFactor
        sumD = sumD * scaleFactor
        sumT = sumT * scaleFactor
        sumO = sumO * scaleFactor
    End If

    resid = kmsDia - (sumP + sumD + sumT + sumO)
    If Abs(resid) > 0.000001 Then
        Select Case targetCat
            Case "Patio":  sumP = sumP + resid
            Case "Diesel": sumD = sumD + resid
            Case "Taller": sumT = sumT + resid
            Case Else:     sumO = sumO + resid
        End Select
    End If

    If sumP < 0 Then sumP = 0
    If sumD < 0 Then sumD = 0
    If sumT < 0 Then sumT = 0
    If sumO < 0 Then sumO = 0

    resid = kmsDia - (sumP + sumD + sumT + sumO)
    If Abs(resid) > 0.000000001 Then
        Select Case targetCat
            Case "Patio":  sumP = sumP + resid
            Case "Diesel": sumD = sumD + resid
            Case "Taller": sumT = sumT + resid
            Case Else:     sumO = sumO + resid
        End Select
    End If

    catPick = IIf(maxCat = "", "Otros", maxCat)

    ' Mantener acumulados por categoría (para Resumen_Agregado / _Total / Global)
    key = divisionName & "|" & veh & "|" & CStr(dd)
    If dicKVcats.Exists(key) Then
        vals = dicKVcats(key)
        vals(0) = CDbl(vals(0)) + sumP
        vals(1) = CDbl(vals(1)) + sumD
        vals(2) = CDbl(vals(2)) + sumT
        vals(3) = CDbl(vals(3)) + sumO
        dicKVcats(key) = vals
    Else
        vals = Array(sumP, sumD, sumT, sumO)
        dicKVcats.Add key, vals
    End If

    ' Salida a hoja KM vacíos deshabilitada: sólo escribir si se pasó una hoja válida.
    If Not wsKM Is Nothing Then
        rM = wsKM.Cells(wsKM.rows.count, 1).End(xlUp).row + 1
        wsKM.Cells(rM, 1).value = divisionName
        wsKM.Cells(rM, 2).value = veh
        wsKM.Cells(rM, 3).value = dd: wsKM.Cells(rM, 3).NumberFormat = "yyyy-mm-dd"
        wsKM.Cells(rM, 4).value = (absStart - dd * 86400) / 86400
        wsKM.Cells(rM, 5).value = (absEnd - dd * 86400) / 86400
        wsKM.Cells(rM, 6).value = Round(kmsDia, 6)
        wsKM.Cells(rM, 7).value = "Sin solape ±" & (TOLERANCIA_MIN) & "m"
        wsKM.Cells(rM, 8).value = srcTag
        wsKM.Cells(rM, 9).value = catPick
        wsKM.Cells(rM, 10).value = maxSite
    End If
End Sub

Public Sub WriteTimelineForDaySegment(ByVal divisionName As String, ByVal veh As String, ByVal dd As Double, _
                                       ByVal segIniD As Long, ByVal segFinD As Long, _
                                       ByVal kmsDia As Double, ByVal durD As Long, _
                                       ByRef dicAut As Object, ByVal tolSec As Long, _
                                       ByRef dicVisitWins As Object, ByVal wsTL As Worksheet)

    If durD <= 0 Or kmsDia <= 0 Then Exit Sub

    Dim absStart As Double, absEnd As Double
    absStart = dd * 86400# + segIniD
    absEnd = dd * 86400# + segFinD

    Dim EPS As Double: EPS = 0.5

    ' === 1) Ventanas autorizadas (Cotización) con cliente (prioridad absoluta)
    Dim authS() As Double, authE() As Double, authC() As String, authId() As String, na As Long
    Dim keyD As String: keyD = veh & "|" & CStr(dd)
    Dim colA As Collection, arA As Variant
    Dim sLoc As Long, eLoc As Long
    Dim authAbsStart As Double, authAbsEnd As Double
    Dim i As Long, k As Long

    If Not dicAut Is Nothing Then
        If dicAut.Exists(keyD) Then
            Set colA = dicAut(keyD)
            For i = 1 To colA.count
                arA = colA(i)                      ' (row, h1, h2, clienteCot)
                sLoc = CLng(arA(1)) - tolSec: If sLoc < 0 Then sLoc = 0
                eLoc = CLng(arA(2)) + tolSec: If eLoc > 86400 Then eLoc = 86400

                authAbsStart = dd * 86400# + sLoc
                authAbsEnd = dd * 86400# + eLoc

                If authAbsEnd > absStart And authAbsStart < absEnd Then
                    If authAbsStart < absStart Then authAbsStart = absStart
                    If authAbsEnd > absEnd Then authAbsEnd = absEnd

                    ReDim Preserve authS(0 To na): ReDim Preserve authE(0 To na): ReDim Preserve authC(0 To na): ReDim Preserve authId(0 To na)
                    authS(na) = authAbsStart
                    authE(na) = authAbsEnd
                    If UBound(arA) >= 3 Then authC(na) = CStr(arA(3)) Else authC(na) = ""
                    If UBound(arA) >= 4 Then authId(na) = CStr(arA(4)) Else authId(na) = ""
                    na = na + 1
                End If
            Next i
        End If
    End If

    ' Ordenar por inicio
    If na > 1 Then
        Dim idxA() As Long, tmpI As Long
        ReDim idxA(0 To na - 1)
        For i = 0 To na - 1: idxA(i) = i: Next i
        For i = 0 To na - 2
            For k = i + 1 To na - 1
                If authS(idxA(k)) < authS(idxA(i)) Then tmpI = idxA(i): idxA(i) = idxA(k): idxA(k) = tmpI
            Next k
        Next i
        Dim s2() As Double, e2() As Double, c2() As String, id2() As String
        ReDim s2(0 To na - 1): ReDim e2(0 To na - 1): ReDim c2(0 To na - 1): ReDim id2(0 To na - 1)
        For i = 0 To na - 1
            s2(i) = authS(idxA(i))
            e2(i) = authE(idxA(i))
            c2(i) = authC(idxA(i))
            id2(i) = authId(idxA(i))
        Next i
        For i = 0 To na - 1
            authS(i) = s2(i): authE(i) = e2(i): authC(i) = c2(i)
            authId(i) = id2(i)
        Next i
    End If

    ' Unir solapes de autorizadas SOLO si el cliente es el mismo
    Dim uS() As Double, uE() As Double, uC() As String, uId() As String, nu As Long
    Dim curS As Double, curE As Double, curC As String, curId As String
    Dim first As Boolean
    If na > 0 Then
        curS = authS(0): curE = authE(0): curC = authC(0): curId = authId(0): first = False
        Dim j As Long
        For i = 1 To na - 1
            If authS(i) <= curE + EPS And authC(i) = curC And authId(i) = curId Then
                If authE(i) > curE Then curE = authE(i)
            Else
                ReDim Preserve uS(0 To nu): ReDim Preserve uE(0 To nu): ReDim Preserve uC(0 To nu): ReDim Preserve uId(0 To nu)
                uS(nu) = curS: uE(nu) = curE: uC(nu) = curC: uId(nu) = curId: nu = nu + 1
                curS = authS(i): curE = authE(i): curC = authC(i): curId = authId(i)
            End If
        Next i
        ReDim Preserve uS(0 To nu): ReDim Preserve uE(0 To nu): ReDim Preserve uC(0 To nu): ReDim Preserve uId(0 To nu)
        uS(nu) = curS: uE(nu) = curE: uC(nu) = curC: uId(nu) = curId: nu = nu + 1
    End If

    ' === 2) Ventanas de Visitas (para NO-Match). Guardar categoría y sitio.
    Dim vs() As Double, vE() As Double, vCat() As String, vSite() As String, nv As Long
    Dim cW As Collection, w As Variant
    Dim visAbsStart As Double, visAbsEnd As Double
    If Not dicVisitWins Is Nothing Then
        If dicVisitWins.Exists(veh) Then
            Set cW = dicVisitWins(veh)
            For i = 1 To cW.count
                w = cW(i) ' [startAbs, endAbs, cat, sitio]
                visAbsStart = CDbl(w(0)): visAbsEnd = CDbl(w(1))
                If visAbsEnd > absStart And visAbsStart < absEnd Then
                    If visAbsStart < absStart Then visAbsStart = absStart
                    If visAbsEnd > absEnd Then visAbsEnd = absEnd
                    ReDim Preserve vs(0 To nv): ReDim Preserve vE(0 To nv)
                    ReDim Preserve vCat(0 To nv): ReDim Preserve vSite(0 To nv)
                    vs(nv) = visAbsStart: vE(nv) = visAbsEnd
                    vCat(nv) = CStr(w(2))
                    If UBound(w) >= 3 Then vSite(nv) = CStr(w(3)) Else vSite(nv) = ""
                    nv = nv + 1
                End If
            Next i
        End If
    End If

    ' === 3) Segmentar: bloques "Match" (uS/uE/uC) y huecos "NM" entre ellos
    Dim segS() As Double, segE() As Double, segT() As String, segCli() As String, segSite() As String, segSvcId() As String, ns As Long
    Dim lastEnd As Double: lastEnd = absStart

    If nu = 0 Then
        ReDim segS(0): ReDim segE(0): ReDim segT(0): ReDim segCli(0): ReDim segSite(0): ReDim segSvcId(0)
        segS(0) = absStart: segE(0) = absEnd: segT(0) = "NM": segCli(0) = "": segSite(0) = "": segSvcId(0) = "": ns = 1
    Else
        For i = 0 To nu - 1
            If uS(i) > lastEnd + EPS Then
                ReDim Preserve segS(0 To ns): ReDim Preserve segE(0 To ns)
                ReDim Preserve segT(0 To ns): ReDim Preserve segCli(0 To ns): ReDim Preserve segSite(0 To ns): ReDim Preserve segSvcId(0 To ns)
                segS(ns) = lastEnd: segE(ns) = uS(i): segT(ns) = "NM": segCli(ns) = "": segSite(ns) = "": segSvcId(ns) = "": ns = ns + 1
            End If
            ReDim Preserve segS(0 To ns): ReDim Preserve segE(0 To ns)
            ReDim Preserve segT(0 To ns): ReDim Preserve segCli(0 To ns): ReDim Preserve segSite(0 To ns): ReDim Preserve segSvcId(0 To ns)
            segS(ns) = uS(i): segE(ns) = uE(i): segT(ns) = "Match": segCli(ns) = uC(i): segSite(ns) = "": segSvcId(ns) = uId(i): ns = ns + 1
            lastEnd = uE(i)
        Next i
        If absEnd > lastEnd + EPS Then
            ReDim Preserve segS(0 To ns): ReDim Preserve segE(0 To ns)
            ReDim Preserve segT(0 To ns): ReDim Preserve segCli(0 To ns): ReDim Preserve segSite(0 To ns): ReDim Preserve segSvcId(0 To ns)
            segS(ns) = lastEnd: segE(ns) = absEnd: segT(ns) = "NM": segCli(ns) = "": segSite(ns) = "": segSvcId(ns) = "": ns = ns + 1
        End If
    End If

    ' === 4) Etiquetar SOLO los "NM" con categoría dominante por Visitas y guardar SiteVisit
    Dim os As Double, oe As Double, ol As Double, best As Double, catPick As String, sitePick As String
    Dim m As Long
    For i = 0 To ns - 1
        If segT(i) = "NM" Then
            best = -1: catPick = "Otros": sitePick = ""
            For m = 0 To nv - 1
                os = segS(i): If vs(m) > os Then os = vs(m)
                oe = segE(i): If vE(m) < oe Then oe = vE(m)
                ol = oe - os
                If ol > best + EPS Then
                    Select Case vCat(m)
                        Case "Patio", "Diesel", "Taller": catPick = vCat(m)
                        Case Else: catPick = "Otros"
                    End Select
                    sitePick = vSite(m)
                    best = ol
                End If
            Next m
            segT(i) = catPick
            segSite(i) = sitePick ' H mostrará SiteVisit para NO-Match
        End If
    Next i

    ' === 5) Propagar identificadores de servicio a segmentos "NM" vecinos ===
    For i = 0 To ns - 1
        If segT(i) <> "Match" Then
            If Len(segSvcId(i)) = 0 Then
                For j = i + 1 To ns - 1
                    If Len(segSvcId(j)) > 0 Then
                        segSvcId(i) = segSvcId(j)
                        Exit For
                    End If
                Next j
                If Len(segSvcId(i)) = 0 Then
                    For j = i - 1 To 0 Step -1
                        If Len(segSvcId(j)) > 0 Then
                            segSvcId(i) = segSvcId(j)
                            Exit For
                        End If
                    Next j
                End If
            End If
        End If
    Next i

    ' === 6) Escribir TL. Precedencia: Match > Visita. H = Cliente si Match; H = SiteVisit si NO-Match.
    Dim gl As Double, accKm As Double, shareKm As Double, minutos As Long, rTL As Long
    accKm = 0
    For i = 0 To ns - 1
        gl = segE(i) - segS(i)
        If gl <= EPS Then GoTo NextSeg
        If i <> ns - 1 Then
            shareKm = kmsDia * (gl / CDbl(durD))
        Else
            shareKm = kmsDia - accKm
        End If
        If shareKm < 0 Then shareKm = 0
        minutos = CLng((gl / 60#) + 0.5)

        rTL = wsTL.Cells(wsTL.rows.count, 1).End(xlUp).row + 1
        wsTL.Cells(rTL, 1).value = divisionName
        wsTL.Cells(rTL, 2).value = veh
        wsTL.Cells(rTL, 3).value = segT(i)                                    ' "Match" o categoría
        wsTL.Cells(rTL, 4).value = dd + (segS(i) - dd * 86400#) / 86400#
        wsTL.Cells(rTL, 5).value = dd + (segE(i) - dd * 86400#) / 86400#
        wsTL.Cells(rTL, 6).value = shareKm
        wsTL.Cells(rTL, 7).value = minutos
        If segT(i) = "Match" Then
            wsTL.Cells(rTL, 8).value = segCli(i)                               ' Cliente de Cotización
        Else
            wsTL.Cells(rTL, 8).value = segSite(i)                              ' SiteVisit para NO-Match
        End If
        wsTL.Cells(rTL, 9).value = segSvcId(i)

        accKm = accKm + shareKm
NextSeg:
    Next i
End Sub

Public Sub GetKVCats(ByRef d As Object, ByVal div_ As String, ByVal veh_ As String, ByVal fecha_ As Double, _
                      ByRef patio As Double, ByRef diesel As Double, ByRef taller As Double, ByRef otros As Double)
    Dim key As String: key = div_ & "|" & veh_ & "|" & CStr(fecha_)
    Dim v As Variant
    If d.Exists(key) Then
        v = d(key)
        patio = CDbl(v(0)): diesel = CDbl(v(1)): taller = CDbl(v(2)): otros = CDbl(v(3))
    Else
        patio = 0: diesel = 0: taller = 0: otros = 0
    End If
End Sub

Public Sub SumKVCatsRange(ByRef d As Object, ByVal div_ As String, ByVal veh_ As String, ByVal dMin As Double, ByVal dMax As Double, _
                           ByRef patio As Double, ByRef diesel As Double, ByRef taller As Double, ByRef otros As Double)
    Dim k As Variant, arr() As String, f As Double, v As Variant
    patio = 0: diesel = 0: taller = 0: otros = 0
    For Each k In d.keys
        arr = Split(CStr(k), "|")
        If arr(0) = div_ And arr(1) = veh_ Then
            f = CDbl(arr(2))
            If f >= dMin And f <= dMax Then
                v = d(k)
                patio = patio + CDbl(v(0))
                diesel = diesel + CDbl(v(1))
                taller = taller + CDbl(v(2))
                otros = otros + CDbl(v(3))
            End If
        End If
    Next k
End Sub

Public Function ClienteForVehDia(ByVal veh As String, ByVal dd As Long) As String
    Dim col As Collection, i As Long, w As Variant, bestAbs As Double
    ClienteForVehDia = ""
    If dicLlegadas Is Nothing Then Exit Function
    If Not dicLlegadas.Exists(veh) Then Exit Function
    Set col = dicLlegadas(veh)
    bestAbs = 1E+99
    For i = 1 To col.count
        w = col(i) ' [absLleg, Cliente, División, DiffMin, Estado, TipoServ, dayInt]
        If CLng(w(6)) = dd Then
            If CDbl(w(0)) < bestAbs Then
                bestAbs = CDbl(w(0))
                ClienteForVehDia = CStr(w(1))
            End If
        End If
    Next i
End Function

