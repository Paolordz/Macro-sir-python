Attribute VB_Name = "Módulo6"
Option Explicit

Public Const PROGRESS_EMPTY_CHAR As String = "."
Public Const DIAG As Boolean = True
Public Const DIAG_MAX_FILAS As Long = 0

Public diagRows As Long

Private Function BuildBar(ByVal pct As Double) As String
    Dim blocks As Long
    blocks = 24
    If pct < 0 Then pct = 0
    If pct > 1 Then pct = 1
    Dim filled As Long
    filled = CLng(pct * blocks + 0.5)
    BuildBar = "[" & String(filled, "#") & String(blocks - filled, PROGRESS_EMPTY_CHAR) & "] " & Format(pct, "0%")
End Function

Public Sub SetStatus(ByVal msg As String, Optional ByVal pct As Double = -1)
    On Error Resume Next
    If pct >= 0 Then
        Application.StatusBar = msg & "  " & BuildBar(pct)
    Else
        Application.StatusBar = msg
    End If
    DoEvents
End Sub

Public Function CountFilesRecursively(ByVal folderPath As String) As Long
    Dim cnt As Long
    cnt = 0
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    Dim f As String
    f = Dir(folderPath & "*.csv")
    Do While Len(f) > 0
        cnt = cnt + 1
        f = Dir
    Loop
    f = Dir(folderPath & "*.xlsx")
    Do While Len(f) > 0
        cnt = cnt + 1
        f = Dir
    Loop
    f = Dir(folderPath & "*.xlsm")
    Do While Len(f) > 0
        cnt = cnt + 1
        f = Dir
    Loop
    f = Dir(folderPath & "*.xlsb")
    Do While Len(f) > 0
        cnt = cnt + 1
        f = Dir
    Loop
    f = Dir(folderPath & "*.xls")
    Do While Len(f) > 0
        cnt = cnt + 1
        f = Dir
    Loop
    Dim subf As String
    subf = Dir(folderPath & "*", vbDirectory)
    Do While Len(subf) > 0
        If subf <> "." And subf <> ".." Then
            If (GetAttr(folderPath & subf) And vbDirectory) <> 0 Then
                cnt = cnt + CountFilesRecursively(folderPath & subf)
            End If
        End If
        subf = Dir
    Loop
    CountFilesRecursively = cnt
End Function

Public Sub LogStart(wb As Workbook)
    diagRows = 0
End Sub

Public Sub LogLine(ByVal wb As Workbook, ByVal src As String, ByVal hdr As Long, _
                    ByVal vVeh As Long, ByVal vFS As Long, ByVal vFFS As Long, _
                    ByVal vH1 As Long, ByVal vH2 As Long, ByVal vKms As Long, _
                    ByVal lrVeh As Long, ByVal lrKms As Long, ByVal rowsToRead As Long, _
                    ByVal notas As String)
    ' NO-OP placeholder for bitácora detallada
End Sub

Public Sub LogFileSheet(ByVal archivo As String, ByVal hoja As String, ByVal divisionName As String, ByVal procesada As Boolean)
    ' NO-OP placeholder para registrar archivos y hojas
End Sub

