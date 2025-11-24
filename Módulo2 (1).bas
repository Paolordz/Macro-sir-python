Attribute VB_Name = "Módulo2"
Option Explicit

Public seenRows As Object
Public totalFiles As Long
Public totalSheets As Long
Public totalFilesExpected As Long
Public currentFileIndex As Long

Public dicVehEndTimes As Object
Public dicVisitWins As Object
Public dicKVcats As Object
Public dicCargas As Object
Public dicCargaAssigned As Object
Public dicLlegadas As Object
Public diasVistos As Object

Public Sub ProcesarMes_SIR()
    On Error GoTo FINALLY

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsCot As Worksheet
    Dim wsRes As Worksheet, wsTot As Worksheet, wsTL As Worksheet
    Dim wsKM As Worksheet

    Dim cVeh As Long, cFS As Long, cFFS As Long, cHIni As Long, cHFin As Long
    Dim cCliCot As Long, cTipoCorto As Long

    Dim cW As Long, cX As Long, cY As Long, cZ As Long, cAA As Long

    Dim folderPath As String

    Dim kvP As Double, kvD As Double, kvT As Double, kvO As Double
    Dim kvP_tot As Double, kvD_tot As Double, kvT_tot As Double, kvO_tot As Double

    Dim dicAut As Object
    Dim dicKMCotizados As Object

    Dim lastRowCot As Long
    Dim i As Long, d1 As Double, d2 As Double, h1 As Long, h2 As Long, veh As String
    Dim d As Long, key As String, rec As Variant
    Dim dicAgg As Object
    Dim tolSec As Long
    Dim processedRows As Long, matchedRows As Long, noMatchRows As Long, processedSheets As Long
    Dim nmCount As Long
    Dim r As Long
    Dim dicTot As Object
    Dim arr() As String, div_ As String, veh_ As String, fecha_ As Double
    Dim vals As Variant, kmTot As Double, kmMatch As Double, kmVacios As Double, pct As Double
    Dim keyTot As String, t As Variant
    Dim rt As Long, kt As Variant, a() As String, dTot As String, vTot As String
    Dim tt As Variant, totKm As Double, totMatch As Double, dMin As Double, dMax As Double, totVacios As Double

    Dim prevCalc As XlCalculation
    Dim prevScreenUpdating As Boolean
    Dim prevStatusBar As Variant
    Dim appStateCaptured As Boolean

    prevCalc = Application.Calculation
    prevScreenUpdating = Application.ScreenUpdating
    prevStatusBar = Application.StatusBar
    appStateCaptured = True

    SetStatus "Iniciando..."
    LogStart wb
    Set seenRows = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set wsCot = wb.Worksheets(HOJA_COT)
    On Error GoTo 0
    If wsCot Is Nothing Then
        MsgBox "No encuentro la hoja '" & HOJA_COT & "'.", vbCritical
        GoTo FINALLY
    End If

    cVeh = ColIndexExact(wsCot, "K_Carro")
    cFS = ColIndexExact(wsCot, "F_Servicio")
    cFFS = ColIndexExact(wsCot, "F_FServicio")
    cHIni = ColIndexExact(wsCot, "HoraInicial")
    cHFin = ColIndexExact(wsCot, "HoraFinal")
    cCliCot = ColIndexExact(wsCot, "C_Cliente")
    cTipoCorto = ColIndexExact(wsCot, "D_Servicio_Tipo_Corto")
    If cTipoCorto = 0 Then cTipoCorto = wsCot.Range("V1").Column
    If cVeh * cFS * cFFS * cHIni * cHFin = 0 Then
        MsgBox "Faltan columnas en 'Cotizacion' (K_Carro, F_Servicio, F_FServicio, HoraInicial, HoraFinal).", vbCritical
        GoTo FINALLY
    End If

    ' Posiciones fijas para KM cotizados por cliente (W:X:Y:Z:AA)
    cW = wsCot.Range("W1").Column
    cX = wsCot.Range("X1").Column
    cY = wsCot.Range("Y1").Column
    cZ = wsCot.Range("Z1").Column
    cAA = wsCot.Range("AA1").Column

    ' BD: llegadas
    Set dicLlegadas = CreateObject("Scripting.Dictionary")
    Dim wsBD As Worksheet
    On Error Resume Next
    Set wsBD = wb.Worksheets("Base de datos")
    On Error GoTo 0
    If Not wsBD Is Nothing Then
        LoadBaseDatosFromSheet wsBD, dicLlegadas
    End If

    folderPath = GetFolderPath()
    If Len(folderPath) = 0 Then GoTo FINALLY
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    totalFilesExpected = CountFilesRecursively(folderPath)
    currentFileIndex = 0
    SetStatus "Archivos a procesar: " & totalFilesExpected

    On Error Resume Next
    Set wsRes = wb.Worksheets(HOJA_RES)
    Set wsTot = wb.Worksheets(HOJA_RES_TOT)
    Set wsTL = wb.Worksheets(HOJA_TL)
    On Error GoTo 0

    If wsRes Is Nothing Then Set wsRes = wb.Worksheets.Add: wsRes.name = HOJA_RES Else wsRes.Cells.Clear
    If wsTot Is Nothing Then Set wsTot = wb.Worksheets.Add: wsTot.name = HOJA_RES_TOT Else wsTot.Cells.Clear
    If wsTL Is Nothing Then Set wsTL = wb.Worksheets.Add: wsTL.name = HOJA_TL Else wsTL.Cells.Clear

    wsRes.Range("A1:K1").value = Array("Division", "Vehiculo", "Fecha", "Km Totales", "Km Match", "Km Vacios", "KV Patio", "KV Diesel", "KV Taller", "KV Otros", "% Eficiencia")
    wsTot.Range("A1:L1").value = Array("Division", "Vehiculo", "Fecha Inicio", "Fecha Fin", "Km Totales", "Km Match", "Km Vacios", "KV Patio", "KV Diesel", "KV Taller", "KV Otros", "% Eficiencia")
    wsTL.Range("A1:I1").value = Array("Division", "Vehiculo", "Tipo", "Inicio", "Fin", "Km", "Min", "Cliente / SiteVisit", "Servicio_Id")
    wsTL.Columns(4).NumberFormat = "yyyy-mm-dd hh:mm": wsTL.Columns(5).NumberFormat = "yyyy-mm-dd hh:mm"
    wsTL.Columns(9).NumberFormat = "@"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    SetStatus "Cargando ventanas desde Cotizacion..."
    Set dicAut = CreateObject("Scripting.Dictionary")
    Set dicKMCotizados = CreateObject("Scripting.Dictionary"): dicKMCotizados.CompareMode = vbTextCompare

    lastRowCot = wsCot.Cells(wsCot.rows.count, cVeh).End(xlUp).row
    For i = 2 To lastRowCot
        veh = Trim$(CStr(wsCot.Cells(i, cVeh).value))
        If Len(veh) = 0 Then GoTo NextCot

        d1 = DateOnlyEx2(wsCot.Cells(i, cFS).value, COT_DATE_ORDER)
        d2 = DateOnlyEx2(wsCot.Cells(i, cFFS).value, COT_DATE_ORDER)
        If d1 = 0 And d2 > 0 Then d1 = d2
        If d2 = 0 And d1 > 0 Then d2 = d1
        If d1 = 0 Then GoTo NextCot

        h1 = TimeToSecEx(wsCot.Cells(i, cHIni).value)
        h2 = TimeToSecEx(wsCot.Cells(i, cHFin).value)

        ' Cliente y tipo corto de esta fila
        Dim cliCot As String, tipoCortoRow As String, cliCotMapped As String
        If cCliCot > 0 Then cliCot = Trim$(CStr(wsCot.Cells(i, cCliCot).value)) Else cliCot = ""
        If cTipoCorto > 0 Then tipoCortoRow = CStr(wsCot.Cells(i, cTipoCorto).value) Else tipoCortoRow = ""
        cliCotMapped = MapClienteNull(cliCot, tipoCortoRow)

        ' Suma KM cotizados W:AA por cliente mapeado
        Dim kmsCotRow As Double, valC As Double, cols As Variant, idx As Long
        kmsCotRow = 0
        cols = Array(cW, cX, cY, cZ, cAA)
        For idx = LBound(cols) To UBound(cols)
            If IsNumeric(wsCot.Cells(i, cols(idx)).value) Then
                valC = CDbl(wsCot.Cells(i, cols(idx)).value)
                If valC > 0 Then kmsCotRow = kmsCotRow + valC
            End If
        Next idx
        If Len(cliCotMapped) > 0 And kmsCotRow > 0 Then
            If dicKMCotizados.Exists(cliCotMapped) Then
                dicKMCotizados(cliCotMapped) = CDbl(dicKMCotizados(cliCotMapped)) + kmsCotRow
            Else
                dicKMCotizados.Add cliCotMapped, kmsCotRow
            End If
        End If

        ' Ventanas autorizadas por día con cliente mapeado
        For d = d1 To d2
            key = veh & "|" & CStr(d)
            If Not dicAut.Exists(key) Then dicAut.Add key, New Collection
            rec = Array(i, h1, h2, cliCotMapped)
            dicAut(key).Add rec
        Next d
NextCot:
        If (i Mod 500) = 0 Then SetStatus "Cargando Cotizacion... " & i - 1 & " / " & (lastRowCot - 1)
    Next i

    Call AttachServicioIdsToDicAut(dicAut)
    Set dicVehEndTimes = CreateObject("Scripting.Dictionary")
    Call BuildVehEndTimes(dicAut, dicVehEndTimes)

    Dim wsLocalV As Worksheet
    On Error Resume Next: Set wsLocalV = ThisWorkbook.Worksheets("Visitas"): On Error GoTo 0
    Set dicVisitWins = CreateObject("Scripting.Dictionary")
    If Not wsLocalV Is Nothing Then
        Call LoadVisitasWindowsFromSheet(wsLocalV, dicVehEndTimes, dicVisitWins)
    Else
        Dim visitasPath As String: visitasPath = GetVisitasPath()
        If Len(visitasPath) > 0 Then
            Call LoadVisitasWindows(visitasPath, dicVehEndTimes, dicVisitWins)
        End If
    End If

    ' Cargas diesel opcionales
    Dim wsLocalC As Worksheet
    Set dicCargas = CreateObject("Scripting.Dictionary")
    Set dicCargaAssigned = CreateObject("Scripting.Dictionary")
    On Error Resume Next: Set wsLocalC = ThisWorkbook.Worksheets("CargaDiesel"): On Error GoTo 0
    If Not wsLocalC Is Nothing Then
        Call LoadCargasFromSheet(wsLocalC, dicCargas)
    End If

    Set dicAgg = CreateObject("Scripting.Dictionary")
    tolSec = TOLERANCIA_MIN * 60
    processedRows = 0: matchedRows = 0: noMatchRows = 0: processedSheets = 0
    nmCount = 0
    Set dicKVcats = CreateObject("Scripting.Dictionary")

    ' wsKM como Nothing para no generar hoja de KM vacíos
    Call ProcesarCarpetaRecursiva(folderPath, dicAut, dicAgg, tolSec, wsKM, nmCount, processedRows, matchedRows, noMatchRows, processedSheets, wsTL)

    SetStatus "Volcando resultados..."
    Set dicTot = CreateObject("Scripting.Dictionary")
    r = 2

    Dim k2 As Variant
    For Each k2 In dicAgg.keys
        arr = Split(CStr(k2), "|")
        div_ = arr(0): veh_ = arr(1): fecha_ = CDbl(arr(2))
        vals = dicAgg(k2)
        kmTot = CDbl(vals(0))
        kmMatch = CDbl(vals(1))
        If kmTot < 0 Then kmTot = 0
        If kmMatch < 0 Then kmMatch = 0
        If kmMatch > kmTot Then kmMatch = kmTot
        kmVacios = kmTot - kmMatch: If kmVacios < 0 Then kmVacios = 0
        If kmTot > 0 Then pct = kmMatch / kmTot Else pct = 0
        If pct < 0 Then pct = 0
        If pct > 1 Then pct = 1

        With wsRes
            .Cells(r, 1).value = div_
            .Cells(r, 2).value = veh_
            .Cells(r, 3).value = fecha_: .Cells(r, 3).NumberFormat = "yyyy-mm-dd"
            .Cells(r, 4).value = kmTot
            .Cells(r, 5).value = kmMatch
            .Cells(r, 6).value = kmVacios
        End With

        Call GetKVCats(dicKVcats, div_, veh_, fecha_, kvP, kvD, kvT, kvO)
        Dim sKV As Double, diffKV As Double
        sKV = kvP + kvD + kvT + kvO
        diffKV = kmVacios - sKV
        If Abs(diffKV) > 0.000001 Then kvO = kvO + diffKV

        With wsRes
            .Cells(r, 7).value = kvP
            .Cells(r, 8).value = kvD
            .Cells(r, 9).value = kvT
            .Cells(r, 10).value = kvO
            .Cells(r, 11).value = pct
        End With
        r = r + 1

        keyTot = div_ & "|" & veh_
        If dicTot.Exists(keyTot) Then
            t = dicTot(keyTot)
            t(0) = CDbl(t(0)) + kmTot
            t(1) = CDbl(t(1)) + kmMatch
            If fecha_ < CDbl(t(2)) Then t(2) = fecha_
            If fecha_ > CDbl(t(3)) Then t(3) = fecha_
            dicTot(keyTot) = t
        Else
            t = Array(CDbl(kmTot), CDbl(kmMatch), CDbl(fecha_), CDbl(fecha_))
            dicTot.Add keyTot, t
        End If
    Next k2

    rt = 2
    For Each kt In dicTot.keys
        a = Split(CStr(kt), "|")
        dTot = a(0): vTot = a(1)
        tt = dicTot(kt)
        totKm = CDbl(tt(0))
        totMatch = CDbl(tt(1))
        dMin = CDbl(tt(2))
        dMax = CDbl(tt(3))
        totVacios = totKm - totMatch: If totVacios < 0 Then totVacios = 0

        With wsTot
            .Cells(rt, 1).value = dTot
            .Cells(rt, 2).value = vTot
            .Cells(rt, 3).value = dMin: .Cells(rt, 3).NumberFormat = "yyyy-mm-dd"
            .Cells(rt, 4).value = dMax: .Cells(rt, 4).NumberFormat = "yyyy-mm-dd"
            .Cells(rt, 5).value = totKm
            .Cells(rt, 6).value = totMatch
            .Cells(rt, 7).value = totVacios
        End With

        Call SumKVCatsRange(dicKVcats, dTot, vTot, dMin, dMax, kvP_tot, kvD_tot, kvT_tot, kvO_tot)
        Dim sKVt As Double, diffKVt As Double
        sKVt = kvP_tot + kvD_tot + kvT_tot + kvO_tot
        diffKVt = totVacios - sKVt
        If Abs(diffKVt) > 0.000001 Then kvO_tot = kvO_tot + diffKVt

        With wsTot
            .Cells(rt, 8).value = kvP_tot
            .Cells(rt, 9).value = kvD_tot
            .Cells(rt, 10).value = kvT_tot
            .Cells(rt, 11).value = kvO_tot
            If totKm > 0 Then .Cells(rt, 12).value = totMatch / totKm Else .Cells(rt, 12).value = 0
        End With
        rt = rt + 1
    Next kt

    Call CrearResumenGlobal(wb, dicTot, dicKVcats)

    Call RebalanceLineaTiempoContraAgg(wsTL, dicAgg)

    Set diasVistos = CreateObject("Scripting.Dictionary")
    Dim kx As Variant, ddOnly As Double
    For Each kx In dicAgg.keys
        arr = Split(CStr(kx), "|")
        ddOnly = CDbl(arr(2))
        If Not diasVistos.Exists(CStr(ddOnly)) Then diasVistos.Add CStr(ddOnly), True
    Next kx

    Dim dicKMCot As Object
    Set dicKMCot = SumarKMCotizadosPorClienteDesdeCotizacion(wsCot)

    Call CrearResumenClienteFiltrado(wb, dicLlegadas, diasVistos, dicAgg, dicKMCotizados, wsBD)

    wsRes.Columns.AutoFit: If r > 2 Then wsRes.Range("K2:K" & r - 1).NumberFormat = "0.0%": wsRes.Range("A1").AutoFilter
    wsTot.Columns.AutoFit: If rt > 2 Then wsTot.Range("L2:L" & rt - 1).NumberFormat = "0.0%": wsTot.Range("A1").AutoFilter
    wsTL.Columns.AutoFit: If wsTL.Cells(wsTL.rows.count, 1).End(xlUp).row > 1 Then wsTL.Range("A1").AutoFilter

    MsgBox "Proceso finalizado." & vbCrLf & _
           "Archivos: " & totalFiles & " | Hojas: " & totalSheets & vbCrLf & _
           "Hojas procesadas: " & processedSheets & vbCrLf & _
           "Filas leídas: " & processedRows & vbCrLf & _
           "Días con solape (match): " & matchedRows & vbCrLf & _
           "Días sin solape (no-match): " & noMatchRows & vbCrLf & _
           "Revisa '" & HOJA_RES & "', '" & HOJA_RES_TOT & "', '" & HOJA_RES_GLOBAL & "' y '" & HOJA_TL & "'.", vbInformation
FINALLY:
    On Error Resume Next
    If appStateCaptured Then
        Application.Calculation = prevCalc
        Application.ScreenUpdating = prevScreenUpdating
        Application.StatusBar = prevStatusBar
    Else
        Application.StatusBar = False
    End If
End Sub

Private Sub AttachServicioIdsToDicAut(ByRef dicAut As Object)
    Dim key As Variant
    Dim parts() As String
    Dim veh As String
    Dim fechaD As Double
    Dim col As Collection
    Dim n As Long
    Dim arr() As Variant
    Dim i As Long, j As Long
    Dim tmp As Variant
    For Each key In dicAut.keys
        parts = Split(CStr(key), "|")
        If UBound(parts) < 1 Then GoTo NextKey
        veh = parts(0)
        On Error Resume Next
        fechaD = CDbl(parts(1))
        On Error GoTo 0
        Set col = dicAut(key)
        n = col.count
        If n = 0 Then GoTo NextKey
        ReDim arr(1 To n)
        For i = 1 To n
            arr(i) = col(i)
        Next i
        For i = 1 To n - 1
            For j = i + 1 To n
                Dim recI As Variant, recJ As Variant
                Dim hi As Long, hJ As Long
                Dim rowI As Long, rowJ As Long
                recI = arr(i)
                recJ = arr(j)
                hi = CLng(recI(1))
                hJ = CLng(recJ(1))
                rowI = CLng(recI(0))
                rowJ = CLng(recJ(0))
                If hi > hJ Or (hi = hJ And rowI > rowJ) Then
                    tmp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = tmp
                End If
            Next j
        Next i
        Dim newCol As Collection
        Set newCol = New Collection
        For i = 1 To n
            Dim rec As Variant
            rec = arr(i)
            If UBound(rec) < 4 Then ReDim Preserve rec(0 To 4)
            rec(4) = ServicioIdFromComponents(veh, fechaD, i)

            newCol.Add rec
        Next i
        Set dicAut(key) = newCol
NextKey:
    Next key
End Sub

