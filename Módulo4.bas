Attribute VB_Name = "Módulo4"
Option Explicit

Public Const HOJA_COT As String = "Cotizacion"
Public Const HOJA_RES As String = "Resumen_Agregado"
Public Const HOJA_RES_TOT As String = "Resumen_Agregado_Total"
Public Const HOJA_RES_GLOBAL As String = "Resumen_Global"
Public Const HOJA_TL As String = "Linea_Tiempo"

Public Const USE_FIXED_FOLDER As Boolean = False
Public Const FIXED_FOLDER As String = "C:\\Ruta\\A\\Tu\\Carpeta"

Public Const COT_DATE_ORDER As String = "MDY"
Public Const MOV_DATE_ORDER As String = "DMY"

Public Function ColIndexExact(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long, c As Long, target As String
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    target = Normalize(headerName)
    For c = 1 To lastCol
        If Normalize(ws.Cells(1, c).value) = target Then
            ColIndexExact = c
            Exit Function
        End If
    Next c
    ColIndexExact = 0
End Function

Public Function Normalize(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(CStr(s)))
    t = Replace(t, "á", "a")
    t = Replace(t, "é", "e")
    t = Replace(t, "í", "i")
    t = Replace(t, "ó", "o")
    t = Replace(t, "ú", "u")
    t = Replace(t, "ü", "u")
    t = Replace(t, " ", "")
    t = Replace(t, "_", "")
    t = Replace(t, ".", "")
    t = Replace(t, "-", "")
    Normalize = t
End Function

Private Function ClienteSiteVisitAliases() As Variant
    ClienteSiteVisitAliases = Array( _
        "Cliente / SiteVisit", _
        "Cliente/ SiteVisit", _
        "Cliente /SiteVisit", _
        "Cliente/SiteVisit", _
        "Cliente  /  SiteVisit", _
        "Cliente SiteVisit", _
        "Cliente-SiteVisit", _
        "Cliente - SiteVisit", _
        "Cliente - Site Visit")
End Function

Public Function GetClienteSiteVisitColumn(ByVal ws As Worksheet, Optional ByVal headerRow As Long = 1) As Long
    GetClienteSiteVisitColumn = FindColAnyInRow(ws, headerRow, ClienteSiteVisitAliases())
    If GetClienteSiteVisitColumn = 0 Then
        GetClienteSiteVisitColumn = ColIndexExact(ws, "Cliente / SiteVisit")
    End If
End Function

Public Sub WarnMissingColumns(ByVal sheetName As String, ByVal missingList As String)
    If Len(missingList) = 0 Then Exit Sub
    MsgBox "Advertencia: No se encontraron las siguientes columnas en '" & sheetName & "':" & missingList, _
           vbExclamation, "Columnas faltantes"
End Sub

Public Function IsReporteUnidadesSheet(ByVal sheetName As String) As Boolean
    Dim n As String
    n = Normalize(sheetName)
    IsReporteUnidadesSheet = (InStr(n, "reporte") > 0 And InStr(n, "unidades") > 0)
End Function

Public Function FindHeaderRow(ws As Worksheet) As Long
    Dim r As Long, lastCol As Long, c As Long
    Dim need1 As String, need2 As String, need3 As String
    need1 = Normalize("Kilómetros")
    need2 = Normalize("Fecha Inicio")
    need3 = Normalize("Hora Inicio")
    For r = 1 To 25
        lastCol = ws.Cells(r, ws.Columns.count).End(xlToLeft).Column
        If Application.CountA(ws.rows(r)) = 0 Then lastCol = 1
        Dim f1 As Boolean, f2 As Boolean, f3 As Boolean
        For c = 1 To lastCol
            Dim h As String
            h = Normalize(ws.Cells(r, c).value)
            If h = need1 Or h = Normalize("kms") Or h = Normalize("kilometros") Then f1 = True
            If h = need2 Or h = Normalize("F Servicio") Or h = Normalize("Fecha") Or h = Normalize("F_Servicio") Then f2 = True
            If h = need3 Or h = Normalize("HoraInicial") Or h = Normalize("Inicio") Or h = Normalize("HI") Then f3 = True
        Next c
        If f1 And f2 And f3 Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r
    FindHeaderRow = 6
End Function

Public Function FindColAnyInRow(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal aliases As Variant) As Long
    Dim lastCol As Long, c As Long, i As Long
    lastCol = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        Dim h As String
        h = Normalize(ws.Cells(headerRow, c).value)
        For i = LBound(aliases) To UBound(aliases)
            If h = Normalize(CStr(aliases(i))) Then
                FindColAnyInRow = c
                Exit Function
            End If
        Next i
    Next c
    FindColAnyInRow = 0
End Function

Public Function DateOnlyEx2(ByVal v As Variant, ByVal order As String) As Double
    On Error GoTo fallback
    If IsDate(v) Then
        DateOnlyEx2 = Int(CDbl(CDate(v)))
        Exit Function
    End If
    Dim s As String
    s = Trim$(CStr(v))
    If s = "" Then
        DateOnlyEx2 = 0
        Exit Function
    End If
    Dim arr() As String, y As Integer, m As Integer, d As Integer
    If InStr(s, "-") > 0 Then
        arr = Split(s, "-")
    ElseIf InStr(s, "/") > 0 Then
        arr = Split(s, "/")
    Else
        GoTo fallback
    End If
    If UBound(arr) <> 2 Then GoTo fallback
    order = UCase$(Trim$(order))
    Select Case order
        Case "MDY"
            m = Val(arr(0))
            d = Val(arr(1))
            y = Val(arr(2))
        Case "DMY"
            d = Val(arr(0))
            m = Val(arr(1))
            y = Val(arr(2))
        Case Else
            GoTo fallback
    End Select
    If y < 100 Then y = 2000 + y
    DateOnlyEx2 = Int(CDbl(DateSerial(y, m, d)))
    Exit Function
fallback:
    DateOnlyEx2 = 0
End Function

Public Function TimeToSecEx(ByVal v As Variant) As Long
    If IsNumeric(v) Then
        If CDbl(v) < 1 Then
            TimeToSecEx = CLng(Round(CDbl(v) * 86400, 0))
            Exit Function
        End If
    End If
    Dim s As String
    s = Trim$(CStr(v))
    If s = "" Then
        TimeToSecEx = 0
        Exit Function
    End If
    If InStr(s, ":") > 0 Then
        Dim p() As String
        p = Split(s, ":")
        Dim h As Long, m As Long, sec As Long
        h = Val(p(0))
        If UBound(p) >= 1 Then
            m = Val(p(1))
        Else
            m = 0
        End If
        If UBound(p) >= 2 Then
            sec = Val(p(2))
        Else
            sec = 0
        End If
        If h < 0 Or h > 47 Or m < 0 Or m > 59 Or sec < 0 Or sec > 59 Then
            TimeToSecEx = 0
        Else
            TimeToSecEx = h * 3600 + m * 60 + sec
        End If
    Else
        Dim n As Long
        n = CLng(Val(s))
        Dim hh As Long
        hh = n \ 100
        Dim mm As Long
        mm = n Mod 100
        If hh < 0 Or hh > 47 Or mm < 0 Or mm > 59 Then
            TimeToSecEx = 0
        Else
            TimeToSecEx = hh * 3600 + mm * 60
        End If
    End If
End Function

Public Function NormalizeVehiculoKey(ByVal v As Variant) As String
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim outS As String

    If IsNull(v) Then
        s = ""
    ElseIf VarType(v) = vbError Then
        s = ""
    ElseIf IsObject(v) Then
        s = ""
    Else
        On Error GoTo ConversionError
        s = Trim$(CStr(v))
        On Error GoTo 0
    End If

    For i = 1 To Len(s)
        ch = mid$(s, i, 1)
        Select Case ch
            Case "0" To "9", "A" To "Z", "a" To "z"
                outS = outS & UCase$(ch)
        End Select
    Next i

    NormalizeVehiculoKey = outS
    Exit Function

ConversionError:
    On Error GoTo 0
    NormalizeVehiculoKey = ""
End Function

Private Function NormalizeVehiculoId(ByVal veh As String) As String
    Dim vehKey As String

    vehKey = NormalizeVehiculoKey(veh)
    If Len(vehKey) = 0 Then vehKey = "NA"

    NormalizeVehiculoId = vehKey
End Function

Public Function ServicioIdFromComponents(ByVal veh As String, ByVal fecha As Variant, ByVal secuencial As Long) As String
    Dim vehKey As String
    Dim fechaKey As String
    vehKey = NormalizeVehiculoId(Trim$(CStr(veh)))
    If IsDate(fecha) Then
        fechaKey = Format$(CDate(fecha), "yyyymmdd")
    ElseIf IsNumeric(fecha) Then
        fechaKey = Format$(CDate(CDbl(fecha)), "yyyymmdd")
    Else
        fechaKey = "00000000"
    End If
    If secuencial < 0 Then secuencial = 0
    ServicioIdFromComponents = vehKey & "-" & fechaKey & "-" & Format$(secuencial, "000")
End Function


Public Function ServicioIdFormComponents(ByVal veh As String, ByVal fecha As Variant, ByVal secuencial As Long) As String
    ServicioIdFormComponents = ServicioIdFromComponents(veh, fecha, secuencial)
End Function


Public Function KmsToDouble(ByVal v As Variant) As Double
    If IsNumeric(v) Then
        KmsToDouble = CDbl(v)
        Exit Function
    End If
    Dim s As String
    s = Trim$(CStr(v))
    If s = "" Then
        KmsToDouble = 0
        Exit Function
    End If
    s = Replace(s, "km", "", 1, -1, vbTextCompare)
    s = Replace(s, "kms", "", 1, -1, vbTextCompare)
    s = Trim$(s)
    If InStr(s, ".") > 0 And InStr(s, ",") > 0 Then
        If InStrRev(s, ",") > InStrRev(s, ".") Then
            s = Replace(s, ".", "")
        Else
            s = Replace(s, ",", "")
        End If
    End If
    s = Replace(s, ",", ".")
    KmsToDouble = Val(s)
End Function

Public Function ExtraerDivisionDesdeNombre(ByVal fileName As String) As String
    Dim base As String
    base = Trim$(CStr(fileName))

    Dim p As Long
    p = InStrRev(base, ".")
    If p > 0 Then base = left$(base, p - 1)
    base = Trim$(base)

    If Len(base) = 0 Then
        ExtraerDivisionDesdeNombre = ""
        Exit Function
    End If

    Dim lowerBase As String
    lowerBase = LCase$(base)

    Dim idx As Long
    idx = 1
    Do While idx <= Len(lowerBase) And Not IsAlphaNumericChar(mid$(lowerBase, idx, 1))
        idx = idx + 1
    Loop

    Dim afterPrefix As Long
    afterPrefix = 0
    If idx + 7 <= Len(lowerBase) And mid$(lowerBase, idx, 8) = "division" Then
        afterPrefix = idx + 8
    ElseIf idx + 7 <= Len(lowerBase) And mid$(lowerBase, idx, 8) = "división" Then
        afterPrefix = idx + 8
    ElseIf idx + 2 <= Len(lowerBase) And mid$(lowerBase, idx, 3) = "div" Then
        afterPrefix = idx + 3
    End If

    If afterPrefix > 0 Then
        Do While afterPrefix <= Len(lowerBase) And Not IsAlphaNumericChar(mid$(lowerBase, afterPrefix, 1))
            afterPrefix = afterPrefix + 1
        Loop

        Dim idStart As Long
        idStart = afterPrefix

        Do While afterPrefix <= Len(lowerBase) And IsAlphaNumericChar(mid$(lowerBase, afterPrefix, 1))
            afterPrefix = afterPrefix + 1
        Loop

        Dim identifier As String
        identifier = Trim$(mid$(base, idStart, afterPrefix - idStart))
        If Len(identifier) > 0 Then
            ExtraerDivisionDesdeNombre = "Division " & FormatearDivisionIdentificador(identifier)
            Exit Function
        End If
    End If

    ExtraerDivisionDesdeNombre = base
End Function

Private Function IsAlphaNumericChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then
        IsAlphaNumericChar = False
        Exit Function
    End If

    Dim code As Long
    code = AscW(ch)

    Select Case code
        Case 48 To 57, 65 To 90, 97 To 122
            IsAlphaNumericChar = True
        Case 192 To 214, 216 To 246, 248 To 255
            IsAlphaNumericChar = True
        Case Else
            IsAlphaNumericChar = False
    End Select
End Function

Private Function FormatearDivisionIdentificador(ByVal identifier As String) As String
    identifier = Trim$(identifier)
    If Len(identifier) = 0 Then
        FormatearDivisionIdentificador = ""
    ElseIf Len(identifier) = 1 Then
        FormatearDivisionIdentificador = UCase$(identifier)
    Else
        FormatearDivisionIdentificador = StrConv(LCase$(identifier), vbProperCase)
    End If
End Function

Public Sub ProbarExtraerDivisionDesdeNombre()
    Debug.Assert ExtraerDivisionDesdeNombre("Division Norte.xlsx") = "Division Norte"
    Debug.Assert ExtraerDivisionDesdeNombre("Div. C.xlsx") = "Division C"
    Debug.Assert ExtraerDivisionDesdeNombre("DIV d reporte.xlsm") = "Division D"
    Debug.Assert ExtraerDivisionDesdeNombre(" div   -   Sur.xls") = "Division Sur"
    Debug.Assert ExtraerDivisionDesdeNombre("Regional.xlsx") = "Regional"
End Sub

Public Function GetFolderPath() As String
    If USE_FIXED_FOLDER Then
        GetFolderPath = FIXED_FOLDER
        Exit Function
    End If
    On Error GoTo fallback
    Dim fd As Object
    Set fd = Application.FileDialog(4)
    With fd
        .Title = "Selecciona la carpeta con los archivos VE (Divisiones)"
        If .Show <> -1 Then
            GetFolderPath = ""
        Else
            GetFolderPath = .SelectedItems(1)
        End If
    End With
    Exit Function
fallback:
    GetFolderPath = InputBox("Escribe la ruta completa de la carpeta con los archivos de divisiones:", _
                             "Carpeta de divisiones")
End Function

Public Function GetVisitasPath() As String
    On Error GoTo fallback
    Dim fd As Object
    Set fd = Application.FileDialog(3)
    With fd
        .Title = "Selecciona el archivo de Reporte de Visitas"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel", "*.xlsx;*.xlsm;*.xlsb;*.xls"
        If .Show <> -1 Then
            GetVisitasPath = ""
        Else
            GetVisitasPath = .SelectedItems(1)
        End If
    End With
    Exit Function
fallback:
    GetVisitasPath = InputBox("Ruta completa del Reporte de Visitas:", "Archivo de Visitas")
End Function

Public Function MapClienteNull(ByVal cli As String, ByVal tipoCorto As String) As String
    Dim c As String
    c = Trim$(cli)
    Dim t As String
    t = UCase$(Trim$(tipoCorto))
    If c = "" Or UCase$(c) = "NULL" Then
        Select Case t
            Case "GD"
                MapClienteNull = "Guardias"
                Exit Function
            Case "VE"
                MapClienteNull = "Viajes especiales"
                Exit Function
        End Select
    End If
    MapClienteNull = cli
End Function

Public Function CatNormalize(ByVal s As String) As String
    Dim n As String
    n = LCase$(Trim$(s))
    If InStr(n, "patio") > 0 Then
        CatNormalize = "Patio"
        Exit Function
    End If
    If InStr(n, "gaso") > 0 Or InStr(n, "diesel") > 0 Or InStr(n, "gasolinera") > 0 Then
        CatNormalize = "Diesel"
        Exit Function
    End If
    If InStr(n, "taller") > 0 Then
        CatNormalize = "Taller"
        Exit Function
    End If
    CatNormalize = "Otros"
End Function

Public Function GuessCatFromSite(ByVal sitio As String) As String
    Dim n As String
    n = LCase$(Trim$(sitio))
    If n = "" Then
        GuessCatFromSite = "Otros"
        Exit Function
    End If
    If InStr(n, "taller") > 0 Or InStr(n, "mecan") > 0 Or InStr(n, "servicio") > 0 Then
        GuessCatFromSite = "Taller"
        Exit Function
    End If
    If InStr(n, "pemex") > 0 Or InStr(n, "diesel") > 0 Or InStr(n, "gasol") > 0 Or InStr(n, "estacion") > 0 Then
        GuessCatFromSite = "Diesel"
        Exit Function
    End If
    If InStr(n, "patio") > 0 Or InStr(n, "base") > 0 Or InStr(n, "cendi") > 0 Then
        GuessCatFromSite = "Patio"
        Exit Function
    End If
    GuessCatFromSite = "Otros"
End Function

Public Function FindHeaderRowVisitas(ByVal ws As Worksheet, Optional ByVal defaultRow As Long = 8) As Long
    Dim lastCol As Long, r As Long, c As Long
    Dim fUnidad As Boolean, fFL As Boolean, fHL As Boolean, fCat As Boolean
    Dim h As String

    For r = 1 To 100
        lastCol = ws.Cells(r, ws.Columns.count).End(xlToLeft).Column
        If Application.CountA(ws.rows(r)) = 0 Then lastCol = 1
        fUnidad = False
        fFL = False
        fHL = False
        fCat = False
        For c = 1 To lastCol
            h = Normalize(ws.Cells(r, c).value)
            If h = "unidad" Or h = "economico" Or h = "economico#" Or h = "noeconomico" Or h = "unidadvehiculo" Then fUnidad = True
            If h = "fechallegada" Or h = "fecha llegada" Or h = "fechaarribo" Or h = "fechaarrive" Or h = "fllegada" Or h = "fecha" Then fFL = True
            If h = "horallegada" Or h = "hora llegada" Or h = "hllegada" Or h = "horafecha" Or h = "hora" Then fHL = True
            If h = "categoria" Or h = "categoría" Or h = "categoriavisita" Or h = "tipo" Or h = "tipovisita" Then fCat = True
        Next c
        If fUnidad And fFL And fHL And fCat Then
            FindHeaderRowVisitas = r
            Exit Function
        End If
    Next r

    FindHeaderRowVisitas = defaultRow
End Function

Public Function FindHeaderRowCargas(ByVal ws As Worksheet, Optional ByVal defaultRow As Long = 1) As Long
    Dim lastCol As Long, r As Long, c As Long
    Dim fUnidad As Boolean, fFecha As Boolean, fDiv As Boolean

    For r = 1 To 100
        lastCol = ws.Cells(r, ws.Columns.count).End(xlToLeft).Column
        If Application.CountA(ws.rows(r)) = 0 Then lastCol = 1
        fUnidad = False
        fFecha = False
        fDiv = False
        For c = 1 To lastCol
            Dim h As String
            h = Normalize(ws.Cells(r, c).value)
            If h = "unidad" Or h = "vehiculo" Or h = "vehículo" Or h = "carro" Then fUnidad = True
            If h = "fecha" Or h = "fregistro" Or h = "f registro" Or h = "fservicio" Then fFecha = True
            If h = "division" Or h = "división" Or h = "div" Then fDiv = True
        Next c
        If fUnidad And fFecha And fDiv Then
            FindHeaderRowCargas = r
            Exit Function
        End If
    Next r

    FindHeaderRowCargas = defaultRow
End Function

Public Function FindHeaderRowBD(ByVal ws As Worksheet) As Long
    Dim r As Long, c As Long, lastCol As Long
    Dim fVeh As Boolean, fFI As Boolean, fHI As Boolean
    For r = 1 To 50
        lastCol = ws.Cells(r, ws.Columns.count).End(xlToLeft).Column
        If Application.CountA(ws.rows(r)) = 0 Then lastCol = 1
        fVeh = False
        fFI = False
        fHI = False
        For c = 1 To lastCol
            Dim h As String
            h = Normalize(ws.Cells(r, c).value)
            If h = "kcarro" Or h = "unidad" Or h = "carro" Or h = "vehiculo" Or h = "vehículo" Then fVeh = True
            If h = "fechainicial" Or h = "fechainicio" Or h = "fservicio" Or h = "f_servicio" Or h = "fecha" Then fFI = True
            If h = "horainicial" Or h = "horainicio" Or h = "hhmm1" Or h = "inicio" Or h = "hi" Then fHI = True
        Next c
        If fVeh And fFI And fHI Then
            FindHeaderRowBD = r
            Exit Function
        End If
    Next r
    FindHeaderRowBD = 1
End Function

Public Function VehLenOK(ByVal veh As String, Optional ByVal maxLen As Long = 6) As Boolean
    Dim s As String
    s = Trim$(veh)
    If s = "" Then
        VehLenOK = False
        Exit Function
    End If
    If Len(s) > maxLen Then
        VehLenOK = False
    Else
        VehLenOK = True
    End If
End Function


