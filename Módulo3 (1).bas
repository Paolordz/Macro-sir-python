Attribute VB_Name = "Módulo3"
Option Explicit

Private Const HEADER_ROW As Long = 1
Private Const MAX_GAP_MINUTES As Double = 10#
Private Const SENTINEL_TIME As Double = -1#

Private Const COL_TL_DIVISION As Long = 1
Private Const COL_TL_VEHICULO As Long = 2
Private Const COL_TL_TIPO As Long = 3
Private Const COL_TL_INICIO As Long = 4
Private Const COL_TL_FIN As Long = 5
Private Const COL_TL_KM As Long = 6
Private Const COL_TL_MIN As Long = 7
Private Const COL_TL_CLIENTE_SITEVISIT As Long = 8

Private Const COL_AL_FECHA As Long = 1
Private Const COL_AL_HORA_INICIO As Long = 2
Private Const COL_AL_HORA_FIN As Long = 3
Private Const COL_AL_VEHICULO As Long = 4
Private Const COL_AL_TIPO As Long = 7

Private Const TYPE_INICIO As String = "INICIO"
Private Const TYPE_FIN As String = "FIN"
Private Const TYPE_ENGANCHE As String = "ENGANCHE"
Private Const TYPE_OTROS As String = "OTROS"
Public Sub GenerarGanttConsolidado()
    Dim wsTimeline As Worksheet
    Dim wsAnalisis As Worksheet
    Dim wsGantt As Worksheet
    Dim prevScreen As Boolean
    Dim prevCalc As XlCalculation
    Dim tlData As Variant
    Dim alData As Variant
    Dim lastRowTL As Long
    Dim lastColTL As Long
    Dim lastRowAL As Long
    Dim lastColAL As Long
    Dim timelineByVeh As Object
    Dim eventsByVeh As Object
    Dim eventsProcessed As Long
    Dim eventsReclassified As Long
    Dim consolidatedRows As Collection
    Dim gantData() As Variant
    Dim msg As String
    Dim fmtInicio As String
    Dim fmtFin As String

    Set wsTimeline = ResolveSheet(Array("Línea_Tiempo", "Linea_Tiempo", "Línea Tiempo", "Linea Tiempo"))
    Set wsAnalisis = ResolveSheet(Array("Análisis Lineal", "Analisis Lineal", "Análisis_Lineal", "Analizis Lineal"))
    Set wsGantt = ResolveSheet(Array("Gantt_consolidado", "Gant_consolidado", "Gantt Consolidado"))

    If wsTimeline Is Nothing Then
        MsgBox "No se encontró la hoja 'Línea_Tiempo'.", vbCritical
        Exit Sub
    End If
    If wsAnalisis Is Nothing Then
        MsgBox "No se encontró la hoja 'Análisis Lineal'.", vbCritical
        Exit Sub
    End If
    If wsGantt Is Nothing Then
        MsgBox "No se encontró la hoja 'Gantt_consolidado'.", vbCritical
        Exit Sub
    End If

    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    lastRowTL = wsTimeline.Cells(wsTimeline.rows.count, 1).End(xlUp).row
    lastColTL = wsTimeline.Cells(HEADER_ROW, wsTimeline.Columns.count).End(xlToLeft).Column
    If lastRowTL <= HEADER_ROW Then
        MsgBox "La hoja 'Línea_Tiempo' no contiene datos para procesar.", vbExclamation
        GoTo CleanExit
    End If

    lastRowAL = wsAnalisis.Cells(wsAnalisis.rows.count, 1).End(xlUp).row
    lastColAL = wsAnalisis.Cells(HEADER_ROW, wsAnalisis.Columns.count).End(xlToLeft).Column
    If lastRowAL <= HEADER_ROW Then
        MsgBox "La hoja 'Análisis Lineal' no contiene datos para procesar.", vbExclamation
        GoTo CleanExit
    End If

    tlData = wsTimeline.Cells(HEADER_ROW, 1).Resize(lastRowTL, lastColTL).value
    alData = wsAnalisis.Cells(HEADER_ROW, 1).Resize(lastRowAL, lastColAL).value

    Set timelineByVeh = CreateObject("Scripting.Dictionary")
    BuildTimelineIndex timelineByVeh, tlData

    Set eventsByVeh = CreateObject("Scripting.Dictionary")
    BuildEventsIndex eventsByVeh, alData

    If eventsByVeh.count > 0 Then
        eventsProcessed = ProcessEvents(eventsByVeh, timelineByVeh, tlData, eventsReclassified)
    End If

    wsTimeline.Cells(HEADER_ROW, 1).Resize(UBound(tlData, 1), UBound(tlData, 2)).value = tlData

    Set consolidatedRows = ConsolidateTimeline(timelineByVeh, tlData, lastColTL)
    If consolidatedRows Is Nothing Then GoTo CleanExit

    ReDim gantData(1 To consolidatedRows.count + 1, 1 To lastColTL)

    Dim c As Long
    For c = 1 To lastColTL
        gantData(1, c) = tlData(HEADER_ROW, c)
    Next c

    Dim i As Long
    For i = 1 To consolidatedRows.count
        Dim rowArray As Variant
        rowArray = consolidatedRows(i)
        For c = 1 To lastColTL
            gantData(i + 1, c) = rowArray(c)
        Next c
    Next i

    wsGantt.Cells.ClearContents
    wsGantt.Cells(HEADER_ROW, 1).Resize(UBound(gantData, 1), UBound(gantData, 2)).value = gantData

    fmtInicio = wsTimeline.Columns(COL_TL_INICIO).NumberFormat
    fmtFin = wsTimeline.Columns(COL_TL_FIN).NumberFormat

    With wsGantt
        .Columns(COL_TL_INICIO).NumberFormat = fmtInicio
        .Columns(COL_TL_FIN).NumberFormat = fmtFin
    End With

    msg = "Proceso completado." & vbCrLf & _
          "Eventos analizados: " & eventsProcessed & vbCrLf & _
          "Eventos reclasificados: " & eventsReclassified & vbCrLf & _
          "Registros consolidados: " & consolidatedRows.count
    MsgBox msg, vbInformation

CleanExit:
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    Exit Sub

CleanFail:
    MsgBox "Se produjo un error al generar el Gantt consolidado: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
Private Function ResolveSheet(ByVal names As Variant) As Worksheet
    Dim i As Long
    For i = LBound(names) To UBound(names)
        On Error Resume Next
        Set ResolveSheet = ThisWorkbook.Worksheets(CStr(names(i)))
        On Error GoTo 0
        If Not ResolveSheet Is Nothing Then Exit Function
    Next i
End Function

Private Sub BuildTimelineIndex(ByRef timelineByVeh As Object, ByRef tlData As Variant)
    Dim lastRow As Long
    Dim r As Long
    Dim vehiculo As String
    Dim rowInfo As Object
    Dim clienteSiteVisit As String
    Dim startVal As Double
    Dim endVal As Double

    lastRow = UBound(tlData, 1)

    For r = HEADER_ROW + 1 To lastRow
        vehiculo = Trim$(NzString(tlData(r, COL_TL_VEHICULO)))
        startVal = ToDateTimeValue(tlData(r, COL_TL_INICIO))
        endVal = ToDateTimeValue(tlData(r, COL_TL_FIN))
        If startVal >= 0# And endVal >= 0# And endVal < startVal Then
            endVal = startVal
        End If

        Set rowInfo = CreateObject("Scripting.Dictionary")
        rowInfo("Index") = r
        rowInfo("Vehiculo") = vehiculo
        rowInfo("Tipo") = NzString(tlData(r, COL_TL_TIPO))
        clienteSiteVisit = EnsureClienteSiteVisitValue(tlData, r)
        rowInfo("ClienteSiteVisit") = clienteSiteVisit
        rowInfo("Inicio") = startVal
        rowInfo("Fin") = endVal

        If Not timelineByVeh.Exists(vehiculo) Then
            Set timelineByVeh(vehiculo) = New Collection
        End If
        timelineByVeh(vehiculo).Add rowInfo
    Next r

    ValidateClienteSiteVisitColumn tlData
End Sub
Private Sub BuildEventsIndex(ByRef eventsByVeh As Object, ByRef alData As Variant)
    Dim lastRow As Long
    Dim r As Long
    Dim vehiculo As String
    Dim tipoNorm As String
    Dim fechaVal As Double
    Dim horaIni As Double
    Dim horaFin As Double
    Dim startVal As Double
    Dim endVal As Double
    Dim eventInfo As Object

    lastRow = UBound(alData, 1)

    For r = HEADER_ROW + 1 To lastRow
        vehiculo = Trim$(NzString(alData(r, COL_AL_VEHICULO)))
        tipoNorm = NormalizeTipo(alData(r, COL_AL_TIPO))
        If Len(vehiculo) = 0 Or Len(tipoNorm) = 0 Then GoTo NextRow

        fechaVal = ToDateTimeValue(alData(r, COL_AL_FECHA))
        horaIni = ToDateTimeValue(alData(r, COL_AL_HORA_INICIO))
        horaFin = ToDateTimeValue(alData(r, COL_AL_HORA_FIN))

        startVal = CombineDateAndTime(fechaVal, horaIni)
        endVal = CombineDateAndTime(fechaVal, horaFin)
        If startVal < 0# Then GoTo NextRow
        If endVal >= 0# And endVal < startVal Then endVal = endVal + 1#

        If (tipoNorm = TYPE_FIN Or tipoNorm = TYPE_ENGANCHE) And endVal < 0# Then GoTo NextRow

        Set eventInfo = CreateObject("Scripting.Dictionary")
        eventInfo("Vehiculo") = vehiculo
        eventInfo("Tipo") = tipoNorm
        eventInfo("TipoTexto") = DisplayTipo(tipoNorm)
        eventInfo("Start") = startVal
        eventInfo("End") = endVal
        eventInfo("Fecha") = Fix(startVal)

        If Not eventsByVeh.Exists(vehiculo) Then
            Set eventsByVeh(vehiculo) = New Collection
        End If
        eventsByVeh(vehiculo).Add eventInfo

NextRow:
    Next r
End Sub

Private Function ProcessEvents(ByRef eventsByVeh As Object, ByRef timelineByVeh As Object, ByRef tlData As Variant, ByRef reclassified As Long) As Long
    Dim vehKey As Variant
    Dim eventsArray As Variant
    Dim evtIndex As Long
    Dim evtInfo As Object
    Dim nextStart As Double
    Dim rowsCollection As Collection
    Dim targetPosition As Long
    Dim eventCount As Long

    reclassified = 0
    For Each vehKey In eventsByVeh.keys
        eventsArray = EventsToSortedArray(eventsByVeh(vehKey))
        If IsEmpty(eventsArray) Then GoTo NextVehicle
        If timelineByVeh.Exists(vehKey) Then
            Set rowsCollection = timelineByVeh(vehKey)
        Else
            Set rowsCollection = Nothing
        End If

        For evtIndex = LBound(eventsArray) To UBound(eventsArray)
            Set evtInfo = eventsArray(evtIndex)
            eventCount = eventCount + 1
            If rowsCollection Is Nothing Then GoTo nextEvent
            If evtInfo("Tipo") = TYPE_ENGANCHE Then
                If evtIndex < UBound(eventsArray) Then
                    nextStart = NzDouble(eventsArray(evtIndex + 1)("Start"))
                Else
                    nextStart = evtInfo("End") + 1#
                End If
            Else
                nextStart = SENTINEL_TIME
            End If
            targetPosition = FindTimelineMatch(rowsCollection, evtInfo, nextStart)
            If targetPosition > 0 Then
                Dim eventStart As Double
                Dim eventEnd As Double
                Dim tolerance As Double
                Dim baseInfo As Object
                Dim forwardIdx As Long
                Dim backwardIdx As Long
                Dim rowInfo As Object
                Dim referenceEnd As Double
                Dim referenceStart As Double
                Dim lastForwardInfo As Object
                Dim lastBackwardInfo As Object

                eventStart = NzDouble(evtInfo("Start"))
                eventEnd = NzDouble(evtInfo("End"))
                If eventStart < 0# Then GoTo nextEvent
                If eventEnd < 0# Or eventEnd < eventStart Then eventEnd = eventStart

                tolerance = (MAX_GAP_MINUTES + 0.0001) / 1440#

                Set baseInfo = rowsCollection(targetPosition)
                ApplyEventToRow baseInfo, evtInfo("TipoTexto"), tlData, reclassified

                referenceEnd = NzDouble(baseInfo("Fin"))
                Set lastForwardInfo = baseInfo
                For forwardIdx = targetPosition + 1 To rowsCollection.count
                    Set rowInfo = rowsCollection(forwardIdx)
                    If Not ShouldReassignRow(lastForwardInfo, rowInfo, eventStart, eventEnd, tolerance, referenceEnd, True) Then Exit For
                    ApplyEventToRow rowInfo, evtInfo("TipoTexto"), tlData, reclassified
                    Set lastForwardInfo = rowInfo
                    If NzDouble(rowInfo("Fin")) >= 0# Then referenceEnd = NzDouble(rowInfo("Fin"))
                Next forwardIdx

                referenceStart = NzDouble(baseInfo("Inicio"))
                Set lastBackwardInfo = baseInfo
                For backwardIdx = targetPosition - 1 To 1 Step -1
                    Set rowInfo = rowsCollection(backwardIdx)
                    If Not ShouldReassignRow(lastBackwardInfo, rowInfo, eventStart, eventEnd, tolerance, referenceStart, False) Then Exit For
                    ApplyEventToRow rowInfo, evtInfo("TipoTexto"), tlData, reclassified
                    Set lastBackwardInfo = rowInfo
                    If NzDouble(rowInfo("Inicio")) >= 0# Then referenceStart = NzDouble(rowInfo("Inicio"))
                Next backwardIdx
            End If
nextEvent:
        Next evtIndex
NextVehicle:
    Next vehKey

    ProcessEvents = eventCount
End Function

Private Sub ApplyEventToRow(ByVal rowInfo As Object, ByVal newTipo As String, ByRef tlData As Variant, ByRef reclassified As Long)
    Dim oldTipoNorm As String
    Dim newTipoNorm As String
    Dim rowIdx As Long
    Dim clienteSiteVisit As String

    rowIdx = rowInfo("Index")
    oldTipoNorm = NormalizeText(NzString(rowInfo("Tipo")))
    newTipoNorm = NormalizeText(NzString(newTipo))

    If oldTipoNorm = TYPE_OTROS And newTipoNorm <> TYPE_OTROS Then
        reclassified = reclassified + 1
    End If

    rowInfo("Tipo") = newTipo
    tlData(rowIdx, COL_TL_TIPO) = newTipo

    clienteSiteVisit = EnsureClienteSiteVisitValue(tlData, rowIdx, newTipo)
    rowInfo("ClienteSiteVisit") = clienteSiteVisit
End Sub

Private Function EnsureClienteSiteVisitValue(ByRef tlData As Variant, ByVal rowIndex As Long, Optional ByVal preferredValue As String = vbNullString) As String
    Dim resolved As String
    Dim trimmedPreferred As String
    Dim canonicalPreferred As String
    Dim canonicalTipo As String

    resolved = Trim$(NzString(tlData(rowIndex, COL_TL_CLIENTE_SITEVISIT)))

    trimmedPreferred = Trim$(preferredValue)
    If Len(trimmedPreferred) > 0 Then
        canonicalTipo = ResolveCanonicalTipo(trimmedPreferred)
        If Len(canonicalTipo) > 0 Then
            canonicalPreferred = DisplayTipo(canonicalTipo)
        Else
            canonicalPreferred = trimmedPreferred
        End If
    End If

    If Len(Trim$(canonicalPreferred)) > 0 Then
        resolved = canonicalPreferred
    Else
        canonicalTipo = ResolveCanonicalTipo(resolved)
        If Len(canonicalTipo) > 0 Then
            resolved = DisplayTipo(canonicalTipo)
        End If

        If Len(resolved) = 0 Then
            resolved = ComposeClienteSiteVisitFallback(tlData, rowIndex)
        End If
    End If

    If Len(resolved) = 0 Then
        resolved = "Sin Cliente / SiteVisit"
    End If

    tlData(rowIndex, COL_TL_CLIENTE_SITEVISIT) = resolved
    EnsureClienteSiteVisitValue = resolved
End Function

Private Function ComposeClienteSiteVisitFallback(ByRef tlData As Variant, ByVal rowIndex As Long) As String
    Dim parts As Object
    Dim clienteCol As Long
    Dim siteVisitCol As Long
    Dim tipoRaw As String
    Dim tipoDisplay As String
    Dim separatorPos As Long
    Dim canonicalTipo As String

    tipoRaw = NzString(tlData(rowIndex, COL_TL_TIPO))
    canonicalTipo = ResolveCanonicalTipo(tipoRaw)
    If Len(canonicalTipo) > 0 Then
        ComposeClienteSiteVisitFallback = DisplayTipo(canonicalTipo)
        Exit Function
    End If

    Set parts = CreateObject("Scripting.Dictionary")

    tipoDisplay = Trim$(tipoRaw)
    separatorPos = InStr(tipoDisplay, " - ")
    If separatorPos > 0 Then
        tipoDisplay = left$(tipoDisplay, separatorPos - 1)
    End If

    If Len(Trim$(tipoDisplay)) > 0 Then
        AddFallbackPart parts, tipoDisplay
    End If

    AddFallbackPart parts, tlData(rowIndex, COL_TL_DIVISION)

    clienteCol = GetTimelineColumnByAliases(tlData, "CLIENTE", Array("Cliente", "Cliente Nombre", "Nombre Cliente"))
    If clienteCol > 0 Then
        AddFallbackPart parts, tlData(rowIndex, clienteCol)
    End If

    siteVisitCol = GetTimelineColumnByAliases(tlData, "SITEVISIT", Array("SiteVisit", "Site Visit", "Visita", "Visita Sitio"))
    If siteVisitCol > 0 Then
        AddFallbackPart parts, tlData(rowIndex, siteVisitCol)
    End If

    AddFallbackPart parts, tlData(rowIndex, COL_TL_VEHICULO)

    ComposeClienteSiteVisitFallback = JoinFallbackParts(parts)
End Function

Private Function ResolveCanonicalTipo(ByVal tipoValue As Variant) As String
    Dim tipoNorm As String
    Dim normalizedText As String
    Dim leadingToken As String

    tipoNorm = NormalizeTipo(tipoValue)
    If Len(tipoNorm) > 0 Then
        ResolveCanonicalTipo = tipoNorm
        Exit Function
    End If

    normalizedText = NormalizeText(NzString(tipoValue))
    If Len(normalizedText) = 0 Then Exit Function

    leadingToken = ExtractLeadingTipoToken(normalizedText)
    Select Case leadingToken
        Case TYPE_INICIO, TYPE_FIN, TYPE_ENGANCHE, TYPE_OTROS
            ResolveCanonicalTipo = leadingToken
    End Select
End Function

Private Function ExtractLeadingTipoToken(ByVal value As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String

    For i = 1 To Len(value)
        ch = mid$(value, i, 1)
        If ch >= "A" And ch <= "Z" Then
            result = result & ch
        ElseIf Len(result) > 0 Then
            Exit For
        End If
    Next i

    ExtractLeadingTipoToken = result
End Function

Private Sub AddFallbackPart(ByVal parts As Object, ByVal value As Variant)
    Dim trimmed As String
    Dim normalized As String

    trimmed = Trim$(NzString(value))
    If Len(trimmed) = 0 Then Exit Sub

    normalized = NormalizeText(trimmed)
    If Len(normalized) = 0 Then Exit Sub

    If Not parts.Exists(normalized) Then
        parts.Add normalized, trimmed
    End If
End Sub

Private Function JoinFallbackParts(ByVal parts As Object) As String
    Dim result As String
    Dim key As Variant

    For Each key In parts.keys
        If Len(result) > 0 Then result = result & " - "
        result = result & parts(key)
    Next key

    JoinFallbackParts = result
End Function

Private Function GetTimelineColumnByAliases(ByRef tlData As Variant, ByVal cacheKey As String, ByVal aliases As Variant) As Long
    Static cache As Object
    Dim normalizedKey As String
    Dim lastCol As Long
    Dim c As Long
    Dim aliasValue As Variant
    Dim headerValue As String

    normalizedKey = NormalizeText(cacheKey)

    If cache Is Nothing Then
        Set cache = CreateObject("Scripting.Dictionary")
    End If

    If cache.Exists(normalizedKey) Then
        GetTimelineColumnByAliases = cache(normalizedKey)
        Exit Function
    End If

    lastCol = UBound(tlData, 2)
    For c = 1 To lastCol
        headerValue = NormalizeText(NzString(tlData(HEADER_ROW, c)))
        For Each aliasValue In aliases
            If headerValue = NormalizeText(CStr(aliasValue)) Then
                cache.Add normalizedKey, c
                GetTimelineColumnByAliases = c
                Exit Function
            End If
        Next aliasValue
    Next c

    cache.Add normalizedKey, 0
    GetTimelineColumnByAliases = 0
End Function

Private Sub ValidateClienteSiteVisitColumn(ByRef tlData As Variant)
    Dim lastRow As Long
    Dim r As Long
    Dim blanks As Long

    lastRow = UBound(tlData, 1)
    For r = HEADER_ROW + 1 To lastRow
        If Len(Trim$(NzString(tlData(r, COL_TL_CLIENTE_SITEVISIT)))) = 0 Then
            blanks = blanks + 1
        End If
    Next r

    If blanks > 0 Then
        Debug.Print "Advertencia: " & blanks & " filas sin Cliente / SiteVisit asignado"
    End If
    Debug.Assert blanks = 0
End Sub

Private Function ShouldReassignRow(ByVal currentInfo As Object, ByVal candidateInfo As Object, ByVal evtStart As Double, ByVal evtEnd As Double, ByVal tolerance As Double, ByVal referenceValue As Double, ByVal isForward As Boolean) As Boolean
    Dim candidateStart As Double
    Dim candidateEnd As Double
    Dim gap As Double

    If Not IsOtrosRow(candidateInfo) Then Exit Function
    If Not RowOverlapsEvent(candidateInfo, evtStart, evtEnd) Then Exit Function

    If Not currentInfo Is Nothing Then
        If NormalizeText(NzString(candidateInfo("Vehiculo"))) <> NormalizeText(NzString(currentInfo("Vehiculo"))) Then Exit Function
    End If

    candidateStart = NzDouble(candidateInfo("Inicio"))
    candidateEnd = NzDouble(candidateInfo("Fin"))

    If isForward Then
        If referenceValue >= 0# And candidateStart >= 0# Then
            gap = candidateStart - referenceValue
            If gap > tolerance Then Exit Function
        End If
    Else
        If referenceValue >= 0# And candidateEnd >= 0# Then
            gap = referenceValue - candidateEnd
            If gap > tolerance Then Exit Function
        End If
    End If

    ShouldReassignRow = True
End Function

Private Function RowOverlapsEvent(ByVal rowInfo As Object, ByVal evtStart As Double, ByVal evtEnd As Double) As Boolean
    Dim rowStart As Double
    Dim rowEnd As Double

    rowStart = NzDouble(rowInfo("Inicio"))
    rowEnd = NzDouble(rowInfo("Fin"))

    If rowStart < 0# And rowEnd >= 0# Then rowStart = rowEnd
    If rowEnd < 0# And rowStart >= 0# Then rowEnd = rowStart
    If rowStart < 0# And rowEnd < 0# Then Exit Function

    If evtEnd < evtStart Then evtEnd = evtStart

    If rowStart <= evtEnd + 0.000001 And rowEnd >= evtStart - 0.000001 Then
        RowOverlapsEvent = True
    End If
End Function
Private Function EventsToSortedArray(ByVal eventsColl As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long
    Dim j As Long
    Dim tmp As Object

    If eventsColl Is Nothing Then
        EventsToSortedArray = Empty
        Exit Function
    End If
    If eventsColl.count = 0 Then
        EventsToSortedArray = Empty
        Exit Function
    End If

    ReDim arr(1 To eventsColl.count)
    For i = 1 To eventsColl.count
        Set arr(i) = eventsColl(i)
    Next i

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If NzDouble(arr(j)("Start")) < NzDouble(arr(i)("Start")) Then
                Set tmp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = tmp
            End If
        Next j
    Next i

    EventsToSortedArray = arr
End Function

Private Function RowsToSortedArray(ByVal rowsColl As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long
    Dim j As Long
    Dim tmp As Object

    If rowsColl Is Nothing Then
        RowsToSortedArray = Empty
        Exit Function
    End If
    If rowsColl.count = 0 Then
        RowsToSortedArray = Empty
        Exit Function
    End If

    ReDim arr(1 To rowsColl.count)
    For i = 1 To rowsColl.count
        Set arr(i) = rowsColl(i)
    Next i

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If NzDouble(arr(j)("Inicio")) < NzDouble(arr(i)("Inicio")) Then
                Set tmp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = tmp
            End If
        Next j
    Next i

    RowsToSortedArray = arr
End Function

Private Function FindTimelineMatch(ByVal rowsCollection As Collection, ByVal evtInfo As Object, ByVal nextStart As Double) As Long
    Dim tipo As String
    Dim i As Long
    Dim rowInfo As Object
    Dim bestIdx As Long
    Dim bestDelta As Double
    Dim evtStart As Double
    Dim evtEnd As Double
    Dim evtDate As Double
    Dim rowStart As Double
    Dim rowEnd As Double
    Dim delta As Double

    tipo = evtInfo("Tipo")
    evtStart = NzDouble(evtInfo("Start"))
    evtEnd = NzDouble(evtInfo("End"))
    evtDate = NzDouble(evtInfo("Fecha"))
    bestDelta = 1000000000000#

    Select Case tipo
        Case TYPE_INICIO
            For i = 1 To rowsCollection.count
                Set rowInfo = rowsCollection(i)
                rowStart = NzDouble(rowInfo("Inicio"))
                rowEnd = NzDouble(rowInfo("Fin"))
                If Not IsOtrosRow(rowInfo) Then GoTo NextRowInicio
                If rowStart < 0# Or rowEnd < 0# Then GoTo NextRowInicio
                If Fix(rowStart) <> evtDate Then GoTo NextRowInicio
                If rowEnd <= evtStart + 0.000001 Then
                    delta = evtStart - rowEnd
                    If delta < 0# Then delta = 0#
                    If delta < bestDelta Then
                        bestDelta = delta
                        bestIdx = i
                    End If
                End If
NextRowInicio:
            Next i
        Case TYPE_FIN
            For i = 1 To rowsCollection.count
                Set rowInfo = rowsCollection(i)
                rowStart = NzDouble(rowInfo("Inicio"))
                If Not IsOtrosRow(rowInfo) Then GoTo NextRowFin
                If rowStart < 0# Then GoTo NextRowFin
                If Fix(rowStart) <> evtDate Then GoTo NextRowFin
                If rowStart >= evtEnd - 0.000001 Then
                    delta = rowStart - evtEnd
                    If delta < 0# Then delta = 0#
                    If delta < bestDelta Then
                        bestDelta = delta
                        bestIdx = i
                    End If
                End If
NextRowFin:
            Next i
        Case TYPE_ENGANCHE
            If evtEnd < 0# Then Exit Function
            Dim upperLimit As Double
            If nextStart > SENTINEL_TIME Then
                upperLimit = nextStart
            Else
                upperLimit = evtEnd + 1#
            End If
            For i = 1 To rowsCollection.count
                Set rowInfo = rowsCollection(i)
                rowStart = NzDouble(rowInfo("Inicio"))
                rowEnd = NzDouble(rowInfo("Fin"))
                If Not IsOtrosRow(rowInfo) Then GoTo NextRowEng
                If rowStart < 0# Or rowEnd < 0# Then GoTo NextRowEng
                If rowEnd >= evtStart - 0.000001 And rowStart <= evtEnd + 0.000001 And rowStart <= upperLimit + 0.000001 Then
                    If evtStart < rowStart Then
                        delta = rowStart - evtStart
                    ElseIf evtStart > rowEnd Then
                        delta = evtStart - rowEnd
                    Else
                        delta = 0#
                    End If
                    If delta < bestDelta Then
                        bestDelta = delta
                        bestIdx = i
                    End If
                End If
NextRowEng:
            Next i
    End Select

    FindTimelineMatch = bestIdx
End Function
Private Function IsOtrosRow(ByVal rowInfo As Object) As Boolean
    IsOtrosRow = (UCase$(NormalizeText(NzString(rowInfo("Tipo")))) = TYPE_OTROS)
End Function

Private Function ConsolidateTimeline(ByRef timelineByVeh As Object, ByRef tlData As Variant, ByVal lastColTL As Long) As Collection
    Dim vehKey As Variant
    Dim groups As New Collection
    Dim rowsArray As Variant
    Dim idx As Long
    Dim rowInfo As Object
    Dim nextInfo As Object
    Dim currentType As String
    Dim currentVeh As String
    Dim currentStart As Double
    Dim currentEnd As Double
    Dim currentMinutes As Double
    Dim currentKm As Double
    Dim firstRowIdx As Long
    Dim gap As Double
    Dim tolerance As Double

    tolerance = (MAX_GAP_MINUTES + 0.0001) / 1440#

    For Each vehKey In timelineByVeh.keys
        rowsArray = RowsToSortedArray(timelineByVeh(vehKey))
        If IsEmpty(rowsArray) Then GoTo NextVehicle
        idx = LBound(rowsArray)
        Do While idx <= UBound(rowsArray)
            Set rowInfo = rowsArray(idx)
            currentType = NzString(tlData(rowInfo("Index"), COL_TL_TIPO))
            currentVeh = NzString(tlData(rowInfo("Index"), COL_TL_VEHICULO))
            currentStart = NzDouble(rowInfo("Inicio"))
            currentEnd = NzDouble(rowInfo("Fin"))
            currentMinutes = GetRowMinutes(rowInfo, tlData)
            currentKm = NzDouble(tlData(rowInfo("Index"), COL_TL_KM))
            firstRowIdx = rowInfo("Index")
            idx = idx + 1
            Do While idx <= UBound(rowsArray)
                Set nextInfo = rowsArray(idx)
                If NormalizeText(NzString(tlData(nextInfo("Index"), COL_TL_TIPO))) <> NormalizeText(currentType) Then Exit Do
                If NormalizeText(NzString(tlData(nextInfo("Index"), COL_TL_VEHICULO))) <> NormalizeText(currentVeh) Then Exit Do
                If currentStart < 0# Or currentEnd < 0# Then Exit Do
                If NzDouble(nextInfo("Inicio")) < 0# Then Exit Do
                gap = NzDouble(nextInfo("Inicio")) - currentEnd
                If gap > tolerance Then Exit Do
                If NzDouble(nextInfo("Fin")) > currentEnd Then currentEnd = NzDouble(nextInfo("Fin"))
                currentMinutes = currentMinutes + GetRowMinutes(nextInfo, tlData)
                currentKm = currentKm + NzDouble(tlData(nextInfo("Index"), COL_TL_KM))
                idx = idx + 1
            Loop
            Dim newRow() As Variant
            ReDim newRow(1 To lastColTL)
            Dim c As Long
            For c = 1 To lastColTL
                newRow(c) = tlData(firstRowIdx, c)
            Next c
            If currentStart >= 0# Then newRow(COL_TL_INICIO) = currentStart
            If currentEnd >= 0# Then newRow(COL_TL_FIN) = currentEnd
            newRow(COL_TL_MIN) = currentMinutes
            newRow(COL_TL_KM) = currentKm
            groups.Add newRow
        Loop
NextVehicle:
    Next vehKey

    Set ConsolidateTimeline = groups
End Function
Private Function GetRowMinutes(ByVal rowInfo As Object, ByRef tlData As Variant) As Double
    Dim rowIdx As Long
    Dim value As Variant
    Dim calculated As Double

    rowIdx = rowInfo("Index")
    value = tlData(rowIdx, COL_TL_MIN)
    If IsNumeric(value) Then
        calculated = CDbl(value)
    Else
        calculated = 0#
    End If
    If calculated <= 0# Then
        calculated = MinutesBetween(NzDouble(rowInfo("Inicio")), NzDouble(rowInfo("Fin")))
    End If
    GetRowMinutes = calculated
End Function

Private Function MinutesBetween(ByVal startVal As Double, ByVal endVal As Double) As Double
    If startVal < 0# Or endVal < 0# Then
        MinutesBetween = 0#
    Else
        MinutesBetween = (endVal - startVal) * 1440#
        If MinutesBetween < 0# Then MinutesBetween = 0#
    End If
End Function

Private Function ToDateTimeValue(ByVal value As Variant) As Double
    On Error GoTo Fail
    If IsDate(value) Then
        ToDateTimeValue = CDbl(CDate(value))
        Exit Function
    End If
    If IsNumeric(value) Then
        ToDateTimeValue = CDbl(value)
        Exit Function
    End If
Fail:
    ToDateTimeValue = SENTINEL_TIME
End Function

Private Function CombineDateAndTime(ByVal dateValue As Double, ByVal timeValue As Double) As Double
    If dateValue < 0# Then
        CombineDateAndTime = SENTINEL_TIME
    ElseIf timeValue < 0# Then
        CombineDateAndTime = SENTINEL_TIME
    Else
        CombineDateAndTime = Fix(dateValue) + (timeValue - Fix(timeValue))
    End If
End Function
Private Function NormalizeTipo(ByVal value As Variant) As String
    Dim textValue As String
    textValue = NormalizeText(NzString(value))
    Select Case textValue
        Case TYPE_INICIO, "INICIO SERVICIO"
            NormalizeTipo = TYPE_INICIO
        Case TYPE_FIN, "FIN SERVICIO"
            NormalizeTipo = TYPE_FIN
        Case TYPE_ENGANCHE
            NormalizeTipo = TYPE_ENGANCHE
        Case TYPE_OTROS
            NormalizeTipo = TYPE_OTROS
        Case Else
            NormalizeTipo = vbNullString
    End Select
End Function

Private Function DisplayTipo(ByVal tipo As String) As String
    Select Case tipo
        Case TYPE_INICIO
            DisplayTipo = "Inicio"
        Case TYPE_FIN
            DisplayTipo = "Fin"
        Case TYPE_ENGANCHE
            DisplayTipo = "Enganche"
        Case TYPE_OTROS
            DisplayTipo = "Otros"
        Case Else
            DisplayTipo = tipo
    End Select
End Function

Private Function NormalizeText(ByVal value As String) As String
    NormalizeText = UCase$(ReplaceAccents(Trim$(value)))
End Function

Private Function ReplaceAccents(ByVal value As String) As String
    value = Replace(value, "Á", "A")
    value = Replace(value, "É", "E")
    value = Replace(value, "Í", "I")
    value = Replace(value, "Ó", "O")
    value = Replace(value, "Ú", "U")
    value = Replace(value, "Ü", "U")
    value = Replace(value, "á", "a")
    value = Replace(value, "é", "e")
    value = Replace(value, "í", "i")
    value = Replace(value, "ó", "o")
    value = Replace(value, "ú", "u")
    value = Replace(value, "ü", "u")
    value = Replace(value, "Ñ", "N")
    value = Replace(value, "ñ", "n")
    ReplaceAccents = value
End Function

Public Sub DebugClienteSiteVisitSamples()
    Dim sampleData(1 To 5, 1 To 8) As Variant
    Dim result As String

    sampleData(1, COL_TL_TIPO) = "Tipo"
    sampleData(1, COL_TL_DIVISION) = "División"
    sampleData(1, COL_TL_CLIENTE_SITEVISIT) = "Cliente/SiteVisit"
    sampleData(1, COL_TL_VEHICULO) = "Vehículo"

    sampleData(2, COL_TL_TIPO) = "OTROS"
    sampleData(2, COL_TL_DIVISION) = "Division A"
    sampleData(2, COL_TL_CLIENTE_SITEVISIT) = "Inicio - e - 192 - Division A"
    sampleData(2, COL_TL_VEHICULO) = "VH-01"

    result = EnsureClienteSiteVisitValue(sampleData, 2, "Inicio")
    Debug.Print "Fila 2 reclasificada a Inicio: ", result

    sampleData(3, COL_TL_TIPO) = "FIN - e - 192"
    sampleData(3, COL_TL_DIVISION) = "Division B"
    sampleData(3, COL_TL_VEHICULO) = "VH-02"

    result = EnsureClienteSiteVisitValue(sampleData, 3, vbNullString)
    Debug.Print "Fila 3 fallback detecta Fin: ", result

    sampleData(4, COL_TL_TIPO) = "OTROS - Observación"
    sampleData(4, COL_TL_DIVISION) = "Division C"
    sampleData(4, COL_TL_VEHICULO) = "VH-03"

    result = EnsureClienteSiteVisitValue(sampleData, 4, vbNullString)
    Debug.Print "Fila 4 fallback detecta Otros: ", result

    sampleData(5, COL_TL_TIPO) = "ENGANCHE"
    sampleData(5, COL_TL_DIVISION) = "Division D"
    sampleData(5, COL_TL_VEHICULO) = "VH-04"

    result = EnsureClienteSiteVisitValue(sampleData, 5, "Enganche")
    Debug.Print "Fila 5 reclasificada a Enganche: ", result
End Sub

Private Function NzDouble(ByVal value As Variant) As Double
    If IsNumeric(value) Then
        NzDouble = CDbl(value)
    Else
        NzDouble = SENTINEL_TIME
    End If
End Function

Private Function NzString(ByVal value As Variant) As String
    If IsError(value) Then
        NzString = vbNullString
    ElseIf IsNull(value) Then
        NzString = vbNullString
    Else
        NzString = CStr(value)
    End If
End Function

