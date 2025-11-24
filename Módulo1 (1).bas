Attribute VB_Name = "Módulo1"
Option Explicit

' ======== TIPO DE DATOS ========
Public Type MovimientoData
    FechaHora As Date
    cliente As String
    LatD As Double
    LonE As Double
    InicioLat As Double
    InicioLon As Double
    FinLat As Double
    FinLon As Double
End Type

' ================== ENTRADA PRINCIPAL ==================
Public Sub ProcesarTramosMovimiento()
    On Error GoTo ErrorHandler

    Dim wsCot As Worksheet, wsOut As Worksheet
    Dim lastRow As Long, r As Long
    Dim grupos As Object, cols As Object
    Dim clave As String

    Set wsCot = GetWorksheet("Cotizacion")
    If wsCot Is Nothing Then Exit Sub

    Set wsOut = GetOrCreateWorksheet("Analizis Lineal")
    Set cols = FindRequiredColumns(wsCot)
    If cols Is Nothing Then Exit Sub

    lastRow = wsCot.Cells(wsCot.rows.count, CLng(cols("Unidad"))).End(xlUp).row
    Set grupos = CreateObject("Scripting.Dictionary")

    For r = 2 To lastRow
        Dim mov As MovimientoData
        Dim unidad As String
        Dim fServ As Date, hIni As Date, hFin As Date

        If ReadMovimientoData(wsCot, r, cols, mov, unidad, fServ, hIni, hFin) Then
            clave = unidad & "|" & Format$(fServ, "yyyy-mm-dd")
            If Not grupos.Exists(clave) Then grupos.Add clave, New Collection

            Dim dataArray(0 To 13) As Variant
            dataArray(0) = fServ          ' fecha
            dataArray(1) = hIni           ' hora inicio
            dataArray(2) = hFin           ' hora fin
            dataArray(3) = mov.cliente
            dataArray(4) = mov.LatD
            dataArray(5) = mov.LonE
            dataArray(6) = mov.InicioLat
            dataArray(7) = mov.InicioLon
            dataArray(8) = mov.FinLat
            dataArray(9) = mov.FinLon
            dataArray(10) = unidad
            dataArray(11) = r            ' fila original en Cotización
            dataArray(12) = HasTimeValue(wsCot.Cells(r, CLng(cols("HoraInicial"))).value)
            dataArray(13) = HasTimeValue(wsCot.Cells(r, CLng(cols("HoraFinal"))).value)
            grupos(clave).Add dataArray
        End If
    Next r

    PrepareAnalizisLinealSheet wsOut
    ProcessAllGroups grupos, wsOut
    FormatAnalizisLinealSheet wsOut
    Exit Sub

ErrorHandler:
    MsgBox "Error en procesamiento: " & Err.Description, vbCritical, "Error"
End Sub

' ================== DISTANCIAS ==================
Private Function EquirectangularKM(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double) As Double
    Const r As Double = 6371
    Dim PI As Double: PI = 4 * Atn(1)
    Dim phi1 As Double, phi2 As Double, lambda1 As Double, lambda2 As Double
    Dim x As Double, y As Double
    phi1 = lat1 * PI / 180#: phi2 = lat2 * PI / 180#
    lambda1 = lon1 * PI / 180#: lambda2 = lon2 * PI / 180#
    x = (lambda2 - lambda1) * Cos((phi1 + phi2) / 2)
    y = (phi2 - phi1)
    EquirectangularKM = Round(r * Sqr(x * x + y * y), 3)
End Function

Private Function HaversineKM(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double) As Double
    Dim PI As Double: PI = 4 * Atn(1)
    Dim r As Double: r = 6371
    Dim dLat As Double, dLon As Double, lat1Rad As Double, lat2Rad As Double
    Dim a As Double, c As Double
    dLat = (lat2 - lat1) * PI / 180#
    dLon = (lon2 - lon1) * PI / 180#
    lat1Rad = lat1 * PI / 180#
    lat2Rad = lat2 * PI / 180#
    a = Sin(dLat / 2) ^ 2 + Cos(lat1Rad) * Cos(lat2Rad) * Sin(dLon / 2) ^ 2
    c = 2 * Application.WorksheetFunction.Atan2(Sqr(a), Sqr(1 - a))
    HaversineKM = Round(r * c, 3)
End Function

' ================== HOJAS ==================
Private Function GetWorksheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = ThisWorkbook.Worksheets(sheetName)
    If GetWorksheet Is Nothing Then
        MsgBox "No se encontró la hoja '" & sheetName & "'", vbCritical, "Error"
    End If
    On Error GoTo 0
End Function

Private Function GetOrCreateWorksheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateWorksheet Is Nothing Then
        Set GetOrCreateWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        GetOrCreateWorksheet.name = sheetName
    End If
End Function

Private Sub PrepareAnalizisLinealSheet(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow > 1 Then ws.Range("A2:Z" & lastRow).Delete xlShiftUp
    ws.Range("A1:N1").value = Array( _
        "Fecha_Servicio", "Hora_Inicial", "Hora_Fin", "Vehiculo", _
        "Km", "Tiempo", "Tipo", "Cliente / SiteVisit", _
        "Desde_Lat", "Desde_Lon", "Hasta_Lat", "Hasta_Lon", _
        "Servicio_Id", "Fila_Cotizacion")
End Sub

Private Sub FormatAnalizisLinealSheet(ws As Worksheet)
    With ws
        .Columns("A:A").NumberFormat = "yyyy-mm-dd"
        .Columns("B:C").NumberFormat = "hh:mm"
        .Columns("D:D").NumberFormat = "0"          ' Vehículo entero
        .Columns("E:E").NumberFormat = "0.000"
        .Columns("F:F").NumberFormat = "0.00"       ' Tiempo con 2 decimales
        .Columns("I:L").NumberFormat = "0.000000"
        .Columns("M:M").NumberFormat = "@"
        .Columns("N:N").NumberFormat = "0"
        .Columns.AutoFit
        Dim lastRow As Long: lastRow = .Cells(.rows.count, 1).End(xlUp).row
        If lastRow > 1 Then .Range("A1:N" & lastRow).AutoFilter
    End With
End Sub

' ================== ENCABEZADOS ==================
Private Function HeaderKey(ByVal s As String) As String
    s = Replace$(CStr(s), Chr(160), " ")
    s = LCase$(Trim$(s))
    s = Replace$(s, "á", "a"): s = Replace$(s, "é", "e")
    s = Replace$(s, "í", "i"): s = Replace$(s, "ó", "o")
    s = Replace$(s, "ú", "u"): s = Replace$(s, "ü", "u")
    s = Replace$(s, "ñ", "n")
    s = Replace$(s, "/", ""): s = Replace$(s, "\", "")
    s = Replace$(s, "-", ""): s = Replace$(s, "_", "")
    s = Replace$(s, " ", "")
    HeaderKey = s
End Function

Private Function FindColumnByKey(ws As Worksheet, targetKey As String, Optional maxHeaderRow As Long = 3) As Long
    Dim lastCol As Long, r As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For r = 1 To maxHeaderRow
        For c = 1 To lastCol
            If HeaderKey(ws.Cells(r, c).value) = targetKey Then
                FindColumnByKey = c
                Exit Function
            End If
        Next c
    Next r
    FindColumnByKey = 0
End Function

Private Function FindColumnAny(ws As Worksheet, keys() As String, Optional fallbackCol As Long = 0) As Long
    Dim k As Long, c As Long
    For k = LBound(keys) To UBound(keys)
        c = FindColumnByKey(ws, keys(k))
        If c > 0 Then FindColumnAny = c: Exit Function
    Next k
    FindColumnAny = fallbackCol
End Function

Private Function FindRequiredColumns(ws As Worksheet) As Object
    Dim cols As Object: Set cols = CreateObject("Scripting.Dictionary")
    Dim kLatD(0 To 4) As String
    Dim kLonE(0 To 4) As String
    Dim kInicioLat(0 To 4) As String
    Dim kInicioLon(0 To 4) As String
    Dim kFinLat(0 To 4) As String
    Dim kFinLon(0 To 4) As String

    kLatD(0) = HeaderKey("LatD"): kLatD(1) = HeaderKey("Lat_D"): kLatD(2) = HeaderKey("LatitudDestino")
    kLatD(3) = HeaderKey("DestinoLat"): kLatD(4) = HeaderKey("Latitud Destino")
    kLonE(0) = HeaderKey("LonE"): kLonE(1) = HeaderKey("Lon_E"): kLonE(2) = HeaderKey("LongitudDestino")
    kLonE(3) = HeaderKey("DestinoLon"): kLonE(4) = HeaderKey("Longitud Destino")
    kInicioLat(0) = HeaderKey("InicioLat"): kInicioLat(1) = HeaderKey("Inicio_Lat")
    kInicioLat(2) = HeaderKey("LatitudInicio"): kInicioLat(3) = HeaderKey("InicioLatitud")
    kInicioLat(4) = HeaderKey("Latitud Inicio")
    kInicioLon(0) = HeaderKey("InicioLon"): kInicioLon(1) = HeaderKey("Inicio_Lon")
    kInicioLon(2) = HeaderKey("LongitudInicio"): kInicioLon(3) = HeaderKey("InicioLongitud")
    kInicioLon(4) = HeaderKey("Longitud Inicio")
    kFinLat(0) = HeaderKey("FinLat"): kFinLat(1) = HeaderKey("Fin_Lat"): kFinLat(2) = HeaderKey("LatitudFin")
    kFinLat(3) = HeaderKey("FinLatitud"): kFinLat(4) = HeaderKey("Latitud Fin")
    kFinLon(0) = HeaderKey("FinLon"): kFinLon(1) = HeaderKey("Fin_Lon"): kFinLon(2) = HeaderKey("LongitudFin")
    kFinLon(3) = HeaderKey("FinLongitud"): kFinLon(4) = HeaderKey("Longitud Fin")

    Dim latDCol As Long, lonECol As Long
    Dim inicioLatCol As Long, inicioLonCol As Long
    Dim finLatCol As Long, finLonCol As Long

    latDCol = FindColumnAny(ws, kLatD, 4)
    lonECol = FindColumnAny(ws, kLonE, 5)
    inicioLatCol = FindColumnAny(ws, kInicioLat, 13)
    inicioLonCol = FindColumnAny(ws, kInicioLon, 14)
    finLatCol = FindColumnAny(ws, kFinLat, 15)
    finLonCol = FindColumnAny(ws, kFinLon, 16)

    If latDCol = 0 Then
        MsgBox "Falta columna de Latitud Destino (LatD).", vbCritical, "Error"
        Set FindRequiredColumns = Nothing
        Exit Function
    End If
    If lonECol = 0 Then
        MsgBox "Falta columna de Longitud Destino (LonE).", vbCritical, "Error"
        Set FindRequiredColumns = Nothing
        Exit Function
    End If
    If inicioLatCol = 0 Then
        MsgBox "Falta columna de Latitud Inicio.", vbCritical, "Error"
        Set FindRequiredColumns = Nothing
        Exit Function
    End If
    If inicioLonCol = 0 Then
        MsgBox "Falta columna de Longitud Inicio.", vbCritical, "Error"
        Set FindRequiredColumns = Nothing
        Exit Function
    End If
    If finLatCol = 0 Then
        MsgBox "Falta columna de Latitud Fin.", vbCritical, "Error"
        Set FindRequiredColumns = Nothing
        Exit Function
    End If
    If finLonCol = 0 Then
        MsgBox "Falta columna de Longitud Fin.", vbCritical, "Error"
        Set FindRequiredColumns = Nothing
        Exit Function
    End If

    cols("LatD") = latDCol
    cols("LonE") = lonECol
    cols("InicioLat") = inicioLatCol
    cols("InicioLon") = inicioLonCol
    cols("FinLat") = finLatCol
    cols("FinLon") = finLonCol

    Dim fServicioCol As Long, horaInicialCol As Long, fFServicioCol As Long, horaFinalCol As Long
    Dim kFS() As String: kFS = Split("fservicio|fechaservicio|fecha", "|")
    Dim kHI() As String: kHI = Split("horainicial|horainicio|hora|horallegada|inicioservicio", "|")
    Dim kFFS() As String: kFFS = Split("ffservicio|fechafinservicio|fechafin", "|")
    Dim kHF() As String: kHF = Split("horafinal|horafin|salidaservicio", "|")

    fServicioCol = FindColumnAny(ws, kFS, 17)
    horaInicialCol = FindColumnAny(ws, kHI, 18)
    fFServicioCol = FindColumnAny(ws, kFFS, 19)
    horaFinalCol = FindColumnAny(ws, kHF, 20)

    Dim unidadCol As Long
    unidadCol = FindColumnByKey(ws, "kcarro")
    If unidadCol = 0 Then
        unidadCol = FindColumnByKey(ws, "unidad")
        If unidadCol = 0 Then
            unidadCol = FindColumnByKey(ws, "vehiculo")
            If unidadCol = 0 Then unidadCol = FindColumnByKey(ws, "vehiculoid")
        End If
    End If
    If unidadCol = 0 Then
        MsgBox "Falta columna de Unidad.", vbCritical, "Error"
        Set FindRequiredColumns = Nothing
        Exit Function
    End If

    Dim clienteCol As Long
    clienteCol = FindColumnByKey(ws, "ccliente")
    If clienteCol = 0 And LCase$(ws.name) = "cotizacion" Then clienteCol = 8

    cols("Unidad") = unidadCol
    If clienteCol <> 0 Then cols("Cliente") = clienteCol
    cols("F_Servicio") = fServicioCol
    cols("HoraInicial") = horaInicialCol
    cols("F_FServicio") = fFServicioCol
    cols("HoraFinal") = horaFinalCol
    cols("UsarFechaHoraEspecificas") = True

    Set FindRequiredColumns = cols
End Function

' ================== PARSEO HORA ==================
Private Function ParseHora(ByVal v As Variant) As Double
    On Error GoTo EH
    If IsDate(v) Then
        ParseHora = CDbl(timeValue(CDate(v)))
        Exit Function
    End If
    If IsNumeric(v) Then
        Dim d As Double: d = CDbl(v)
        ParseHora = d - Fix(d)
        If ParseHora < 0 Then ParseHora = 0
        Exit Function
    End If
    Dim s As String: s = Trim$(CStr(v))
    If s = "" Then ParseHora = 0: Exit Function
    s = Replace$(s, ".", ":")
    Dim parts() As String: parts = Split(s, ":")
    Dim hh As Long, mm As Long, ss As Long
    hh = Val(parts(0))
    If UBound(parts) >= 1 Then mm = Val(parts(1))
    If UBound(parts) >= 2 Then ss = Val(parts(2))
    If hh < 0 Or hh > 23 Or mm < 0 Or mm > 59 Or ss < 0 Or ss > 59 Then
        ParseHora = 0
    Else
        ParseHora = TimeSerial(hh, mm, ss)
    End If
    Exit Function
EH:
    ParseHora = 0
End Function

Private Function HasTimeValue(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then
        HasTimeValue = False
    ElseIf VarType(v) = vbString Then
        HasTimeValue = Len(Trim$(CStr(v))) > 0
    Else
        HasTimeValue = True
    End If
End Function

' ================== LECTURA ==================
Private Function ReadMovimientoData(ws As Worksheet, row As Long, cols As Object, _
                                  ByRef mov As MovimientoData, ByRef unidad As String, _
                                  ByRef fServicio As Date, ByRef horaInicial As Date, _
                                  ByRef horaFinal As Date) As Boolean
    On Error GoTo ReadError

    Dim latDCol As Long, lonECol As Long
    Dim inicioLatCol As Long, inicioLonCol As Long
    Dim finLatCol As Long, finLonCol As Long

    If Not cols.Exists("LatD") Then GoTo ReadError Else latDCol = CLng(cols("LatD"))
    If Not cols.Exists("LonE") Then GoTo ReadError Else lonECol = CLng(cols("LonE"))
    If Not cols.Exists("InicioLat") Then GoTo ReadError Else inicioLatCol = CLng(cols("InicioLat"))
    If Not cols.Exists("InicioLon") Then GoTo ReadError Else inicioLonCol = CLng(cols("InicioLon"))
    If Not cols.Exists("FinLat") Then GoTo ReadError Else finLatCol = CLng(cols("FinLat"))
    If Not cols.Exists("FinLon") Then GoTo ReadError Else finLonCol = CLng(cols("FinLon"))

    unidad = Trim$(CStr(ws.Cells(row, CLng(cols("Unidad"))).value))
    If Len(unidad) = 0 Then GoTo ReadError

    If IsDate(ws.Cells(row, CLng(cols("F_Servicio"))).value) Then
        fServicio = dateValue(CDate(ws.Cells(row, CLng(cols("F_Servicio"))).value))
    Else
        GoTo ReadError
    End If

    horaInicial = ParseHora(ws.Cells(row, CLng(cols("HoraInicial"))).value)
    horaFinal = ParseHora(ws.Cells(row, CLng(cols("HoraFinal"))).value)

    If Not IsValidCoordinateCell(ws.Cells(row, latDCol)) Then GoTo ReadError
    If Not IsValidCoordinateCell(ws.Cells(row, lonECol)) Then GoTo ReadError
    If Not IsValidCoordinateCell(ws.Cells(row, inicioLatCol)) Then GoTo ReadError
    If Not IsValidCoordinateCell(ws.Cells(row, inicioLonCol)) Then GoTo ReadError
    If Not IsValidCoordinateCell(ws.Cells(row, finLatCol)) Then GoTo ReadError
    If Not IsValidCoordinateCell(ws.Cells(row, finLonCol)) Then GoTo ReadError

    mov.LatD = CDbl(ws.Cells(row, latDCol).value)
    mov.LonE = CDbl(ws.Cells(row, lonECol).value)
    mov.InicioLat = CDbl(ws.Cells(row, inicioLatCol).value)
    mov.InicioLon = CDbl(ws.Cells(row, inicioLonCol).value)
    mov.FinLat = CDbl(ws.Cells(row, finLatCol).value)
    mov.FinLon = CDbl(ws.Cells(row, finLonCol).value)

    If cols.Exists("Cliente") Then
        mov.cliente = Trim$(CStr(ws.Cells(row, CLng(cols("Cliente"))).value))
    Else
        mov.cliente = vbNullString
    End If

    mov.FechaHora = fServicio + horaInicial

    ReadMovimientoData = True
    Exit Function
ReadError:
    ReadMovimientoData = False
End Function

Private Function IsValidCoordinateCell(cell As Range) As Boolean
    On Error GoTo InvalidCell
    If IsEmpty(cell.value) Or IsError(cell.value) Then
        IsValidCoordinateCell = False
        Exit Function
    End If
    If Not IsNumeric(cell.value) Then
        IsValidCoordinateCell = False
        Exit Function
    End If
    Dim coord As Double: coord = CDbl(cell.value)
    IsValidCoordinateCell = True
    Exit Function
InvalidCell:
    IsValidCoordinateCell = False
End Function

' ================== PROCESAMIENTO ==================
Private Sub ProcessAllGroups(grupos As Object, wsOut As Worksheet)
    Dim clave As Variant
    Dim NextRow As Long: NextRow = 2
    For Each clave In grupos.keys
        Dim grupo As Collection: Set grupo = grupos(clave)
        If grupo.count > 0 Then NextRow = ProcessSingleGroup(grupo, wsOut, NextRow)
    Next clave
End Sub

Private Function ProcessSingleGroup(grupo As Collection, wsOut As Worksheet, startRow As Long) As Long
    SortGroupByTime grupo

    Dim currentRow As Long: currentRow = startRow
    Dim i As Long
    For i = 1 To grupo.count
        Dim data As Variant: data = grupo(i)
        Dim mov As MovimientoData, unidad As String
        Dim fServ As Date, hIni As Date, hFin As Date
        Dim servicioId As String
        Dim cotRow As Long

        fServ = CDate(data(0))
        hIni = data(1)
        hFin = data(2)
        mov.cliente = CStr(data(3))
        mov.LatD = CDbl(data(4))
        mov.LonE = CDbl(data(5))
        mov.InicioLat = CDbl(data(6))
        mov.InicioLon = CDbl(data(7))
        mov.FinLat = CDbl(data(8))
        mov.FinLon = CDbl(data(9))
        unidad = CStr(data(10))
        cotRow = CLng(data(11))

        servicioId = ServicioIdFormComponents(unidad, fServ, i)
        servicioId = ServicioIdFromComponents(unidad, fServ, i)

        If i = 1 Then
            currentRow = WriteTramoToSheet(wsOut, currentRow, fServ, hIni, hFin, unidad, _
                        mov.LatD, mov.LonE, mov.InicioLat, mov.InicioLon, _
                        "Inicio", BuildClienteInfo("Inicio", mov.cliente), servicioId, cotRow)
        End If

        If i > 1 Then
            Dim prevData As Variant: prevData = grupo(i - 1)
            Dim prevFinLat As Double, prevFinLon As Double
            prevFinLat = CDbl(prevData(8))
            prevFinLon = CDbl(prevData(9))
            Dim prevDataUBound As Long: prevDataUBound = UBound(prevData)
            Dim dataUBound As Long: dataUBound = UBound(data)
            Dim hasPrevHoraFin As Boolean
            Dim hasCurrHoraIni As Boolean
            If prevDataUBound >= 13 Then hasPrevHoraFin = CBool(prevData(13))
            If dataUBound >= 12 Then hasCurrHoraIni = CBool(data(12))

            Dim prevHoraFinVal As Double
            Dim currHoraIniVal As Double
            If hasPrevHoraFin Then prevHoraFinVal = NzTime(prevData(2))
            If hasCurrHoraIni Then currHoraIniVal = NzTime(data(1))

            If hasPrevHoraFin Or hasCurrHoraIni Then
                If Not hasPrevHoraFin Then prevHoraFinVal = currHoraIniVal
                If Not hasCurrHoraIni Then currHoraIniVal = prevHoraFinVal

                Dim engancheFecha As Date
                Dim engancheHoraIni As Date
                Dim engancheHoraFin As Date
                engancheFecha = dateValue(CDate(prevData(0)))
                engancheHoraIni = CDate(prevHoraFinVal)
                engancheHoraFin = CDate(currHoraIniVal)

                currentRow = WriteTramoToSheet(wsOut, currentRow, engancheFecha, engancheHoraIni, engancheHoraFin, unidad, _
                            prevFinLat, prevFinLon, mov.InicioLat, mov.InicioLon, _
                            "Enganche", BuildClienteInfo("Enganche", mov.cliente), servicioId, cotRow)
            End If
        End If

        If i = grupo.count Then
            currentRow = WriteTramoToSheet(wsOut, currentRow, fServ, hIni, hFin, unidad, _
                        mov.FinLat, mov.FinLon, mov.LatD, mov.LonE, _
                        "Fin", BuildClienteInfo("Fin", mov.cliente), servicioId, cotRow)
        End If
    Next i

    ProcessSingleGroup = currentRow
End Function

Private Function BuildClienteInfo(ByVal tipo As String, ByVal cliente As String) As String
    If Len(Trim$(cliente)) > 0 Then
        BuildClienteInfo = tipo & "(" & cliente & ")"
    Else
        BuildClienteInfo = tipo
    End If
End Function

Private Function WriteTramoToSheet(ws As Worksheet, row As Long, _
                                 fServ As Date, hIni As Date, hFin As Date, _
                                 unidad As String, lat1 As Double, lon1 As Double, _
                                 lat2 As Double, lon2 As Double, tipo As String, _
                                 clienteInfo As String, servicioId As String, cotRow As Long) As Long
    Dim km As Double, tiempo As Double
    km = EquirectangularKM(lat1, lon1, lat2, lon2)
    If km < 0.01 Then
        WriteTramoToSheet = row
        Exit Function
    End If
    tiempo = (km / 35) * 60

    With ws
        .Cells(row, 1).value = dateValue(fServ)                       ' A Fecha
        If hIni = 0 Then .Cells(row, 2).ClearContents Else .Cells(row, 2).value = timeValue(hIni) ' B Hora Inicio
        If hFin = 0 Then .Cells(row, 3).ClearContents Else .Cells(row, 3).value = timeValue(hFin) ' C Hora Fin
        .Cells(row, 4).value = UnidadFmt(unidad)                      ' D Vehículo entero si aplica
        .Cells(row, 5).value = km                                     ' E Km
        .Cells(row, 6).value = tiempo                                 ' F Tiempo
        .Cells(row, 7).value = tipo                                   ' G Tipo
        .Cells(row, 8).value = clienteInfo                            ' H Cliente/SiteVisit
        .Cells(row, 9).value = lat1                                   ' I Desde_Lat
        .Cells(row, 10).value = lon1                                  ' J Desde_Lon
        .Cells(row, 11).value = lat2                                  ' K Hasta_Lat
        .Cells(row, 12).value = lon2                                  ' L Hasta_Lon
        .Cells(row, 13).value = servicioId                            ' M Servicio_Id
        .Cells(row, 14).value = cotRow                                 ' N Fila Cotización
    End With
    WriteTramoToSheet = row + 1
End Function

' ================== ORDENAMIENTO ==================
Private Sub SortGroupByTime(grupo As Collection)
    Dim n As Long: n = grupo.count
    If n <= 1 Then Exit Sub

    Dim arr() As Variant, i As Long, j As Long, tmp As Variant
    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = grupo(i)
    Next i

    For i = 1 To n - 1
        For j = i + 1 To n
            Dim tI As Date, tJ As Date
            Dim rowI As Long, rowJ As Long
            tI = CDate(arr(i)(0)) + NzTime(arr(i)(1))
            tJ = CDate(arr(j)(0)) + NzTime(arr(j)(1))
            rowI = CLng(arr(i)(11))
            rowJ = CLng(arr(j)(11))
            If tI > tJ Or (Abs(tI - tJ) < (1# / 864000#) And rowI > rowJ) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i

    For i = n To 1 Step -1
        grupo.Remove i
    Next i
    For i = 1 To n
        grupo.Add arr(i)
    Next i
End Sub

Private Function NzTime(v As Variant) As Double
    If IsDate(v) Then
        NzTime = CDbl(timeValue(CDate(v)))
    ElseIf IsNumeric(v) Then
        NzTime = CDbl(v) - Fix(CDbl(v))
    Else
        NzTime = 0
    End If
End Function

' ================== FORMATO VEHÍCULO ==================
Private Function UnidadFmt(ByVal u As Variant) As Variant
    Dim vehKey As String

    vehKey = NormalizeVehiculoKey(u)

    If Len(vehKey) = 0 Then
        UnidadFmt = vbNullString
    ElseIf IsNumeric(vehKey) Then
        UnidadFmt = CLng(CDbl(vehKey))   ' quita decimales .000
    Else
        UnidadFmt = vehKey
    End If
End Function

