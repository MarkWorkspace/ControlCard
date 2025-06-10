Option Explicit
'=============================================================
' Раскрой утеплителя
'=============================================================
Sub CuttingPlanInsulation()
    Const DATA_SHEET As String = "ИсходныеДанные"
    Const PARAM_SHEET As String = "Параметры"
    Const OUT_SHEET As String = "Раскрой Утеплителя"
    Const FIRST_ROW As Long = 12
    ' В ячейках AA2 и AB2 листа Параметры должны быть указаны
    ' максимальная ширина и высота блока утеплителя (мм)

    Dim wsData As Worksheet, wsParam As Worksheet, wsOut As Worksheet
    On Error Resume Next
    Set wsData = Worksheets(DATA_SHEET)
    Set wsParam = Worksheets(PARAM_SHEET)
    On Error GoTo 0
    If wsData Is Nothing Or wsParam Is Nothing Then
        MsgBox "Отсутствует лист данных или параметров", vbCritical: Exit Sub
    End If

    Dim maxW As Double, maxH As Double
    maxW = Val(wsParam.Range("AA2").Value)
    maxH = Val(wsParam.Range("AB2").Value)
    If maxW <= 0 Or maxH <= 0 Then
        MsgBox "Неверные размеры блока утеплителя в Параметры!AA2:AB2", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wsOut = Worksheets(OUT_SHEET)
    On Error GoTo 0
    If wsOut Is Nothing Then
        Set wsOut = Worksheets.Add(After:=wsData)
        wsOut.Name = OUT_SHEET
    Else
        wsOut.Cells.Clear
    End If
    wsOut.Activate
    Application.DisplayAlerts = True

    wsOut.Cells.Font.Name = "Calibri"
    wsOut.Cells.HorizontalAlignment = xlCenter
    wsOut.Cells.VerticalAlignment = xlCenter
    wsOut.Columns("A").ColumnWidth = 6
    wsOut.Columns("B").ColumnWidth = 15
    wsOut.Columns("C").ColumnWidth = 8
    wsOut.Columns("D").ColumnWidth = 20

    wsOut.Cells(1, 1).Resize(1, 4).Value = Array("№", "Размер", "Кол-во", "Источник")

    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, "C").End(xlUp).Row

    Dim r As Long, outRow As Long, partIdx As Long
    outRow = 2
    partIdx = 1
    For r = FIRST_ROW To lastRow
        Dim layer As String, dims As String, qty As Long
        layer = Trim(wsData.Cells(r, "C").Value)
        dims = Trim(wsData.Cells(r, "D").Value)
        qty = Val(wsData.Cells(r, "F").Value)

        If qty <= 0 Then GoTo nextRow
        If InStr(1, layer, "утеп", vbTextCompare) = 0 Then GoTo nextRow
        If InStr(1, dims, "x", vbTextCompare) = 0 Then GoTo nextRow

        Dim arrDims() As String
        arrDims = Split(dims, "x")
        If UBound(arrDims) < 1 Then GoTo nextRow
        Dim width As Double, height As Double
        width = Val(arrDims(0))
        height = Val(arrDims(1))
        If width <= 0 Or height <= 0 Then GoTo nextRow

        Dim y As Double, x As Double, pw As Double, ph As Double
        y = 0
        Do While y < height
            ph = maxH
            If height - y < maxH Then ph = height - y
            x = 0
            Do While x < width
                pw = maxW
                If width - x < maxW Then pw = width - x
                wsOut.Cells(outRow, 1).Value = partIdx
                wsOut.Cells(outRow, 2).Value = Format(pw, "0") & "x" & Format(ph, "0")
                wsOut.Cells(outRow, 3).Value = qty
                wsOut.Cells(outRow, 4).Value = "Деталь " & r
                partIdx = partIdx + 1
                outRow = outRow + 1
                x = x + maxW
            Loop
            y = y + maxH
        Loop
nextRow:
    Next r

    Application.ScreenUpdating = True
End Sub
