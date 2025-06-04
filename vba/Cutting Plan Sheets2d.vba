Option Explicit
'=====================================================================
' Раскрой плит
'=====================================================================
Sub CuttingPlanSheets2D()
    Const DATA_SHEET_NAME As String = "ИсходныеДанные"
    Const START_ROW = 12
    Const KERF_MM = 3
    Const MAX_PX_SIDE = 400
    Const N_MAP_ROWS = 12
    Const STD_ROW_HEIGHT = 16

    Dim wsData As Worksheet, wsOut As Worksheet
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    Dim blockStats As Object: Set blockStats = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error Resume Next
    Set wsOut = ThisWorkbook.Sheets("Раскрой Плит")
    On Error GoTo 0

    If wsOut Is Nothing Then
        ' Если листа нет — создаём
        Set wsOut = ThisWorkbook.Sheets.Add(After:=wsData)
        wsOut.Name = "Раскрой Плит"
    Else
        ' Если лист есть — очищаем всё
        wsOut.Cells.Clear
        wsOut.Cells.ClearFormats
    End If

    wsOut.Activate
    Application.DisplayAlerts = True

    With wsOut
        .Cells.Clear
        Dim shpTmp As Shape
        For Each shpTmp In .Shapes: shpTmp.Delete: Next
        .Cells.Font.Name = "Calibri"
        .Cells.HorizontalAlignment = xlCenter
        .Columns("A").ColumnWidth = 8
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 10
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 12
        .Columns("F").ColumnWidth = 12
        .Columns("G").ColumnWidth = 1
        .Columns("H").ColumnWidth = 100
    End With

    Dim palette(1 To 4) As Long
    palette(1) = RGB(192, 96, 96)
    palette(2) = RGB(224, 192, 96)
    palette(3) = RGB(96, 192, 96)
    palette(4) = RGB(96, 192, 192)

    Dim NotPlacedDetails As Collection
    Set NotPlacedDetails = New Collection
    Dim maxRegionW As Double: maxRegionW = 0
    Dim dictLayers As Object: Set dictLayers = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long: lastRow = wsData.Cells(wsData.Rows.Count, "C").End(xlUp).Row

    Dim r As Long, layerName As String, section As String, qty As Long, idx As Variant
    For r = START_ROW To lastRow
        layerName = Trim(wsData.Cells(r, "C").Value)
        section = Trim(wsData.Cells(r, "D").Value)
        qty = Val(wsData.Cells(r, "F").Value)
        If layerName <> "" And section <> "" And qty > 0 And Val(wsData.Cells(r, "E").Value) <= 10 Then
            If Not dictLayers.Exists(layerName) Then Set dictLayers(layerName) = New Collection
            dictLayers(layerName).Add r
        End If
    Next

    Dim curRow As Long: curRow = 1
    Dim layerKey As Variant, plateGroups As Object, plateKey As Variant
    Dim dimsPlate As Variant, dimsPart As Variant
    Dim plateW As Double, plateH As Double, plateArea As Double, usedArea As Double

    For Each layerKey In dictLayers.Keys
        wsOut.Range(wsOut.Cells(curRow, 1), wsOut.Cells(curRow, 8)).Merge
        wsOut.Cells(curRow, 1).Value = layerKey
        wsOut.Cells(curRow, 1).Interior.Color = RGB(192, 192, 192)
        wsOut.Cells(curRow, 1).Font.Bold = True
        With wsOut.Range(wsOut.Cells(curRow, 1), wsOut.Cells(curRow, 8))
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(100, 100, 100)
        End With

        curRow = curRow + 1

        Set plateGroups = CreateObject("Scripting.Dictionary")
        For Each idx In dictLayers(layerKey)
            plateKey = Trim(wsData.Cells(idx, "J").Value)
            If Not plateGroups.Exists(plateKey) Then Set plateGroups(plateKey) = New Collection
            plateGroups(plateKey).Add idx
        Next

        For Each plateKey In plateGroups.Keys
            If InStr(1, plateKey, "x", vbTextCompare) = 0 Then GoTo SkipPlateKey
            dimsPlate = Split(plateKey, "x")
            If UBound(dimsPlate) < 1 Then GoTo SkipPlateKey
            plateW = Val(dimsPlate(0))
            plateH = Val(dimsPlate(1))
            If plateW <= 0 Or plateH <= 0 Then GoTo SkipPlateKey

            Dim needRotate As Boolean
            If plateH > plateW Then
                needRotate = True
                Dim tmpVal As Double: tmpVal = plateW: plateW = plateH: plateH = tmpVal
            End If

            Dim Remaining As Object: Set Remaining = New Collection
            For Each idx In plateGroups(plateKey)
                qty = Val(wsData.Cells(idx, "F").Value)
                dimsPart = Split(wsData.Cells(idx, "D").Value, "x")
                If UBound(dimsPart) < 1 Then GoTo NextIdx
                Dim partWOrig As Double, partHOrig As Double
                partWOrig = Val(dimsPart(0)): partHOrig = Val(dimsPart(1))
                Dim copyIndex As Long
                For copyIndex = 1 To qty
                    Remaining.Add Array(partWOrig, partHOrig, wsData.Cells(idx, "D").Value, idx)
                Next
NextIdx:
            Next

            SortByLongestSide Remaining

            Dim Blocks As Object: Set Blocks = New Collection
            Do While Remaining.Count > 0
                Dim Block As Object: Set Block = New Collection
                Dim freeRects As Object: Set freeRects = New Collection
                freeRects.Add Array(0#, 0#, plateW, plateH)

                Dim placedAny As Boolean: placedAny = True
                Do While placedAny And Remaining.Count > 0
                    placedAny = False
                    Dim bestIdx As Long: bestIdx = 0
                    Dim bestRectIdx As Long: bestRectIdx = 0
                    Dim bestW As Double, bestH As Double, bestX As Double, bestY As Double
                    Dim i As Long, rct As Variant

                    For i = 1 To Remaining.Count
                        Dim w0 As Double: w0 = Remaining(i)(0)
                        Dim h0 As Double: h0 = Remaining(i)(1)
                        Dim j As Long
                        For j = 1 To freeRects.Count
                            rct = freeRects(j)
                            If w0 <= rct(2) And h0 <= rct(3) Then
                                bestIdx = i: bestRectIdx = j
                                bestW = w0: bestH = h0
                                bestX = rct(0): bestY = rct(1)
                                Exit For
                            End If
                            If h0 <= rct(2) And w0 <= rct(3) Then
                                bestIdx = i: bestRectIdx = j
                                bestW = h0: bestH = w0
                                bestX = rct(0): bestY = rct(1)
                                Exit For
                            End If
                        Next
                        If bestIdx > 0 Then Exit For
                    Next
                    If bestIdx = 0 Then Exit Do

                    Block.Add Array(bestW, bestH, Remaining(bestIdx)(2), bestX, bestY)
                    Remaining.Remove bestIdx
                    placedAny = True

                    Dim fx As Double, fy As Double, fw As Double, fh As Double
                    fx = rct(0): fy = rct(1): fw = rct(2): fh = rct(3)
                    ' Исправление: удалять только если индекс валиден
                    If bestRectIdx > 0 And bestRectIdx <= freeRects.Count Then
                        freeRects.Remove bestRectIdx
                    End If
                    If fx + bestW < fx + fw Then freeRects.Add Array(fx + bestW, fy, fw - bestW, bestH)
                    If fy + bestH < fy + fh Then freeRects.Add Array(fx, fy + bestH, fw, fh - bestH)
                Loop
                If Block.Count = 0 Then Exit Do
                Blocks.Add Block
            Loop

            If Remaining.Count > 0 Then
                For i = 1 To Remaining.Count
                    NotPlacedDetails.Add Array(wsData.Cells(Remaining(i)(3), "A").Value, plateKey, Remaining(i)(2), "Не удалось разместить на листе " & plateKey)
                Next
            End If

           ' === 1. Группировка одинаковых блоков по сигнатуре (типу заполнения) ===
            Dim b As Long, p As Long, rowStart As Long, shp As Shape
            Dim DictUniqueBlocks As Object: Set DictUniqueBlocks = CreateObject("Scripting.Dictionary")
            Dim BlockSignatures() As String
            ReDim BlockSignatures(1 To Blocks.Count)
            For b = 1 To Blocks.Count
                BlockSignatures(b) = GetBlockSignature(Blocks(b))
                If Not DictUniqueBlocks.Exists(BlockSignatures(b)) Then
                    Set DictUniqueBlocks(BlockSignatures(b)) = New Collection
                End If
                DictUniqueBlocks(BlockSignatures(b)).Add b
            Next

            ' === 2. Проход по уникальным блокам, оформление шапки и подготовка таблицы ===
            Dim uniqKey As Variant, blockIdx As Long
            Dim plateIndex As Long: plateIndex = 1
            For Each uniqKey In DictUniqueBlocks.Keys
                blockIdx = DictUniqueBlocks(uniqKey)(1)
                Dim copiesCount As Long: copiesCount = DictUniqueBlocks(uniqKey).Count

                


                ' --- 2.1. Формируем шапку для блока (плиты) ---
                wsOut.Range(wsOut.Cells(curRow, 1), wsOut.Cells(curRow, 8)).Merge
                Dim blockHeader As String
                blockHeader = "#" & plateIndex & " (" & plateW & "x" & plateH & ")"
                If copiesCount > 1 Then blockHeader = blockHeader & " x" & copiesCount
                wsOut.Cells(curRow, 1).Value = blockHeader
                wsOut.Cells(curRow, 1).Interior.Color = RGB(220, 220, 220)
                wsOut.Cells(curRow, 1).Font.Bold = True

                With wsOut.Range(wsOut.Cells(curRow, 1), wsOut.Cells(curRow, 8))
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                    .Borders.Color = RGB(140, 140, 140)
                End With

                plateIndex = plateIndex + 1
                plateArea = plateW * plateH / 1000000
                usedArea = 0
                wsOut.Cells(curRow, 5).Value = "Плита м²: " & Format(plateArea, "0.000")
                wsOut.Cells(curRow, 6).Value = "Обрезки м²: " & Format(plateArea - usedArea, "0.000")
                curRow = curRow + 1
                wsOut.Cells(curRow, 1).Resize(1, 8).Value = Array("#", "Размер", "Деталь м^2", "S исп. м^2", "S ост. м^2", "шт.", "", "Раскрой                   ")
                curRow = curRow + 1
                rowStart = curRow
                Dim detailRowDict As Object: Set detailRowDict = CreateObject("Scripting.Dictionary")

                ' --- 2.2. Подсчет количества одинаковых деталей на плите ---
                Dim itm As Variant, pw As Double, ph As Double, px As Double, py As Double
                Dim lnkAddr As String, tempArr As Variant, detailKey As Variant
                Dim detailsDict As Object: Set detailsDict = CreateObject("Scripting.Dictionary")
                For p = 1 To Blocks(blockIdx).Count
                    itm = Blocks(blockIdx)(p)
                    pw = itm(0): ph = itm(1)
                    detailKey = pw & "x" & ph
                    If Not detailsDict.Exists(detailKey) Then
                        detailsDict(detailKey) = Array(pw, ph, itm(2), 1)
                    Else
                        tempArr = detailsDict(detailKey)
                        tempArr(3) = tempArr(3) + 1
                        detailsDict(detailKey) = tempArr
                    End If
                Next

                ' --- 2.2.1. Подсчёт общей площади деталей на плите (usedArea) ---
                Dim detArr As Variant

                usedArea = 0
                For Each detailKey In detailsDict.Keys
                    detArr = detailsDict(detailKey)
                    usedArea = usedArea + detArr(0) * detArr(1) / 1000000 * detArr(3)
                Next
                Dim restArea As Double: restArea = plateArea - usedArea
                If restArea < 0 Then restArea = 0

                ' --- 2.3. Вывод данных по деталям в итоговую таблицу блока ---
                Dim outRow As Long: outRow = rowStart
                Dim k As Long: k = 1
                For Each detailKey In detailsDict.Keys
                    detArr = detailsDict(detailKey)
                    wsOut.Cells(outRow, 1).Value = k
                    wsOut.Cells(outRow, 2).Value = detArr(0) & "x" & detArr(1)
                    Dim partArea As Double: partArea = detArr(0) * detArr(1) / 1000000
                    wsOut.Cells(outRow, 3).Value = partArea
                    ' Только для первой строки блока
                    If k = 1 Then
                        wsOut.Cells(outRow, 4).Value = usedArea
                        wsOut.Cells(outRow, 5).Value = restArea
                    Else
                        wsOut.Cells(outRow, 4).Value = ""
                        wsOut.Cells(outRow, 5).Value = ""
                    End If
                    wsOut.Cells(outRow, 6).Value = detArr(3)
                    wsOut.Cells(outRow, 8).Value = ""
                    detailRowDict(detailKey) = outRow
                    k = k + 1
                    outRow = outRow + 1
                Next

                
                 ' --- 2.4. Рисуем раскладку деталей на плите (визуализация) ---
                Dim iRow As Long
                For iRow = 0 To N_MAP_ROWS - 1
                    wsOut.Rows(curRow + iRow).RowHeight = STD_ROW_HEIGHT
                Next
                ' Оформление
                With wsOut.Range(wsOut.Cells(curRow - 1, "G"), wsOut.Cells(curRow + N_MAP_ROWS, "H"))
                    .Interior.Color = RGB(255, 255, 255)
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeLeft).Weight = xlThin
                    .Borders(xlEdgeLeft).Color = RGB(180, 180, 180)
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).Weight = xlThin
                    .Borders(xlEdgeRight).Color = RGB(180, 180, 180)
                    ' Не трогаем верхнюю и нижнюю границы
                    .Borders(xlInsideVertical).LineStyle = xlNone
                    .Borders(xlInsideHorizontal).LineStyle = xlNone
                End With


                Dim regionW As Double, regionH As Double
                regionW = wsOut.Columns("H").Width
                regionH = N_MAP_ROWS * STD_ROW_HEIGHT
                Dim leftPos As Double, topPos As Double, scaleX As Double, scaleY As Double
                leftPos = wsOut.Cells(curRow, "H").Left
                topPos = wsOut.Cells(curRow, "H").Top
                scaleX = regionW / plateW
                scaleY = regionH / plateH
                If scaleY < scaleX Then scaleX = scaleY: regionW = plateW * scaleX
                If scaleX < scaleY Then scaleY = scaleX: regionH = plateH * scaleY

                If regionW > maxRegionW Then maxRegionW = regionW

                Set shp = wsOut.Shapes.AddShape(msoShapeRectangle, leftPos, topPos, regionW, regionH)
                shp.Line.Visible = msoFalse
                shp.Fill.ForeColor.RGB = RGB(200, 200, 200)
                For p = 1 To Blocks(blockIdx).Count
                    itm = Blocks(blockIdx)(p)
                    pw = itm(0): ph = itm(1)
                    px = itm(3): py = itm(4)
                    detailKey = pw & "x" & ph
                    Set shp = wsOut.Shapes.AddShape(msoShapeRectangle, leftPos + px * scaleX, topPos + py * scaleY, pw * scaleX, ph * scaleY)
                    shp.Name = "DET_" & (rowStart + p - 1)
                    ' --- ссылка ведет на строку итоговой таблицы (объединенную), соответствующую размеру детали ---
                    lnkAddr = "'" & wsOut.Name & "'!B" & detailRowDict(detailKey)
                    wsOut.Hyperlinks.Add Anchor:=shp, Address:="", SubAddress:=lnkAddr, TextToDisplay:=""
                    With shp.TextFrame2
                        .VerticalAnchor = msoAnchorMiddle
                        .HorizontalAnchor = msoAnchorCenter
                        .TextRange.Text = pw & "x" & ph
                        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                    End With
                    shp.AlternativeText = "Размер: " & pw & " x " & ph
                    shp.Fill.ForeColor.RGB = palette((p Mod 4) + 1)
                    shp.Line.ForeColor.RGB = vbBlack: shp.Line.Weight = 0.25
                    usedArea = usedArea + pw * ph / 1000000

                    Application.ScreenUpdating = True
                    wsOut.Cells(curRow, 1).Select
                    DoEvents
                    Application.ScreenUpdating = False
                Next

                curRow = curRow + N_MAP_ROWS + 1

                ' === Здесь добавляем сбор данных для итоговой таблицы ===
                If Blocks.Count > 0 Then
                    Dim nBlocks As Long: nBlocks = Blocks.Count
                    Dim sumRestArea As Double: sumRestArea = 0
                    Dim restAreas As Collection: Set restAreas = New Collection
                    For b = 1 To Blocks.Count
                        Dim blockUsedArea As Double: blockUsedArea = 0
                        For p = 1 To Blocks(b).Count
                            itm = Blocks(b)(p)
                            blockUsedArea = blockUsedArea + itm(0) * itm(1) / 1000000
                        Next
                        Dim plateRest As Double
                        plateRest = plateW * plateH / 1000000 - blockUsedArea
                        If plateRest < 0 Then plateRest = 0
                        restAreas.Add plateRest
                    Next

                    sumRestArea = 0
                    For i = 1 To restAreas.Count
                        sumRestArea = sumRestArea + restAreas(i)
                    Next

                    Dim cat As String, paint As String, colorText As String
                    Dim wsP As Worksheet
                    On Error Resume Next
                    Set wsP = Worksheets("Параметры")
                    On Error GoTo 0
                    If Not wsP Is Nothing Then
                        Dim mRow As Variant: mRow = Application.Match(layerKey, wsP.Range("S:S"), 0)
                        If Not IsError(mRow) Then
                            cat = wsP.Cells(mRow, "T").Value
                            paint = wsP.Cells(mRow, "U").Value
                        End If
                    End If

                    colorText = ""
                    For r = START_ROW To lastRow
                        If Trim(wsData.Cells(r, "C").Value) = layerKey And Trim(wsData.Cells(r, "J").Value) = plateKey Then
                            colorText = wsData.Cells(r, "K").Text
                            Exit For
                        End If
                    Next
                    Dim comboKey As Variant
                    comboKey = layerKey & "|" & plateKey
                    blockStats(comboKey) = Array(layerKey, plateKey, plateW, plateH, restAreas.Count, sumRestArea, cat, paint, colorText, sumRestArea)

                End If
            Next
SkipPlateKey:
        Next
    Next

    If NotPlacedDetails.Count > 0 Then
        curRow = curRow + 2
        wsOut.Range(wsOut.Cells(curRow, 1), wsOut.Cells(curRow, 5)).Merge
        wsOut.Cells(curRow, 1).Value = "Детали, не попавшие в раскрой"
        wsOut.Cells(curRow, 1).Font.Bold = True
        wsOut.Cells(curRow, 1).Interior.Color = RGB(255, 220, 220)
        curRow = curRow + 1
        wsOut.Cells(curRow, 1).Resize(1, 8).Value = Array("№", "Плита", "Деталь", "", "", "", "Причина")
        wsOut.Rows(curRow).Font.Bold = True
        Dim npd As Variant
        For i = 1 To NotPlacedDetails.Count
            npd = NotPlacedDetails(i)
            wsOut.Cells(curRow + i, 1).Value = npd(0)
            wsOut.Cells(curRow + i, 2).Value = npd(1)
            wsOut.Cells(curRow + i, 3).Value = npd(2)
            wsOut.Cells(curRow + i, 8).Value = npd(3)
        Next
        wsOut.Columns("A:F").AutoFit
    End If

    Dim pxPerColWidth As Double
    pxPerColWidth = wsOut.Columns("H").Width / 100
    If maxRegionW > 0 Then
        wsOut.Columns("H").ColumnWidth = (maxRegionW) / pxPerColWidth + 1
    End If

       ' === Формируем итоговую таблицу для плит ===
    Dim tblCol As Long: tblCol = wsOut.Columns("Q").Column
    Dim tblRow As Long: tblRow = 1
    Dim summaryDict As Object: Set summaryDict = CreateObject("Scripting.Dictionary")
    

    ' Собираем данные по всем плитам и слоям
    For Each layerKey In dictLayers.Keys
        For Each plateKey In plateGroups.Keys
            If InStr(1, plateKey, "x", vbTextCompare) = 0 Then GoTo SkipPlate
            dimsPlate = Split(plateKey, "x")
            If UBound(dimsPlate) < 1 Then GoTo SkipPlate
            plateW = Val(dimsPlate(0))
            plateH = Val(dimsPlate(1))
            If plateW <= 0 Or plateH <= 0 Then GoTo SkipPlate
SkipPlate:
        Next
    Next

    ' --- Оформление итоговой таблицы ---
    Dim hdrs As Variant: hdrs = Array("Слой", "Ширина", "Длина", "Кол-во", "S ост. м^2", "Категория", "Покраска", "Цвет")
    Dim hi As Long
    For hi = 0 To UBound(hdrs)
        With wsOut.Cells(tblRow, tblCol + hi)
            .Value = hdrs(hi)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Interior.Color = RGB(230, 230, 230)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = vbBlack
            .Borders.Weight = xlThin
        End With

        ' Установка ширины колонок
        Select Case hi
            Case 0: wsOut.Columns(tblCol + hi).ColumnWidth = 48    ' Слой
            Case 1: wsOut.Columns(tblCol + hi).ColumnWidth = 10    ' Ширина
            Case 2: wsOut.Columns(tblCol + hi).ColumnWidth = 10    ' Длина
            Case 3: wsOut.Columns(tblCol + hi).ColumnWidth = 10    ' Кол-во плит
            Case 4: wsOut.Columns(tblCol + hi).ColumnWidth = 17    ' Площадь ост., м^2
            Case 5: wsOut.Columns(tblCol + hi).ColumnWidth = 38    ' Категория
            Case 6: wsOut.Columns(tblCol + hi).ColumnWidth = 13    ' Покраска
            Case 7: wsOut.Columns(tblCol + hi).ColumnWidth = 44    ' Цвет
        End Select
    Next hi
    tblRow = tblRow + 1

    ' --- Вывод данных итоговой таблицы ---
    Dim iTbl As Long: iTbl = tblRow

    Dim restSumByLayer As Object: Set restSumByLayer = CreateObject("Scripting.Dictionary")
    For Each comboKey In blockStats.Keys
        Dim v As Variant: v = blockStats(comboKey)
        If Not restSumByLayer.Exists(v(0)) Then restSumByLayer(v(0)) = 0
        restSumByLayer(v(0)) = restSumByLayer(v(0)) + v(9)
    Next

   Dim outRows As Object: Set outRows = CreateObject("Scripting.Dictionary")
For Each comboKey In blockStats.Keys
    v = blockStats(comboKey)
    Dim restSum As Double: restSum = 0
    If restSumByLayer.Exists(v(0)) Then restSum = restSumByLayer(v(0))
    wsOut.Cells(iTbl, tblCol).Resize(1, 8).Value = Array( _
        v(0), v(2), v(3), v(4), restSum, v(6), v(7), v(8))  ' restSum в пятое место (U)

    ' --- Оформление итоговой таблицы (легкая обводка) ---
        Dim tblLastRow As Long
        tblLastRow = iTbl - 1 ' Последняя строка итоговой таблицы

        With wsOut.Range(wsOut.Cells(tblRow - (iTbl - tblRow), tblCol), wsOut.Cells(tblLastRow, tblCol + 7))
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(140, 140, 140)
        End With

    iTbl = iTbl + 1
    tblRow = iTbl
Next

    ' Фиксированная ширина колонок I:P
    wsOut.Columns("I:P").ColumnWidth = 2

    ' ======= Лог для плит, не попавших в итоговую таблицу =======
    ' Собираем все уникальные плиты по слоям из исходных данных
    Dim missingPlates As Collection
    Set missingPlates = New Collection

    Dim allPlatesDict As Object
    Set allPlatesDict = CreateObject("Scripting.Dictionary")
    For Each layerKey In dictLayers.Keys
        Set plateGroups = CreateObject("Scripting.Dictionary")
        For Each idx In dictLayers(layerKey)
            plateKey = Trim(wsData.Cells(idx, "J").Value)
            If Not plateGroups.Exists(plateKey) Then Set plateGroups(plateKey) = New Collection
            plateGroups(plateKey).Add idx
        Next
        For Each plateKey In plateGroups.Keys
            comboKey = layerKey & "|" & plateKey
            allPlatesDict(comboKey) = Array(layerKey, plateKey)
        Next
    Next

    ' Находим те плиты, которых нет в blockStats (итоговой таблице)
    For Each comboKey In allPlatesDict.Keys
        If Not blockStats.Exists(comboKey) Then
            missingPlates.Add allPlatesDict(comboKey)
        End If
    Next

    ' Формируем лог по отсутствующим плитам
    If missingPlates.Count > 0 Then
        Dim logRow As Long
        logRow = tblRow + 3
        wsOut.Range(wsOut.Cells(logRow, tblCol), wsOut.Cells(logRow, tblCol + 2)).Merge
        With wsOut.Cells(logRow, tblCol)
            .Value = "Плиты, не попавшие в итоговую таблицу"
            .Font.Bold = True
            .Interior.Color = RGB(255, 220, 220)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        logRow = logRow + 1
        wsOut.Cells(logRow, tblCol).Resize(1, 4).Value = Array("Слой", "Плита", "Причина", "")
        wsOut.Rows(logRow).Font.Bold = True

        Dim mp As Variant
        For i = 1 To missingPlates.Count
            mp = missingPlates(i)
            wsOut.Cells(logRow + i, tblCol).Value = mp(0)
            wsOut.Cells(logRow + i, tblCol + 1).Value = mp(1)
            wsOut.Cells(logRow + i, tblCol + 2).Value = "Не использована в раскрое"
        Next
        wsOut.Columns(tblCol).ColumnWidth = 48
        wsOut.Columns(tblCol + 1).ColumnWidth = 15
        wsOut.Columns(tblCol + 2).ColumnWidth = 26
    End If
    MsgBox "Раскрой плит (2D) завершён!", vbInformation
    Application.Goto wsOut.Range("Q1"), True
End Sub


Private Sub SortByLongestSide(aList As Object)
    Dim arr() As Variant
    Dim i As Long, j As Long, n As Long
    n = aList.Count
    If n <= 1 Then Exit Sub
    ReDim arr(1 To n)
    For i = 1 To n: arr(i) = aList(i): Next
    Dim maxI As Double, maxJ As Double, tmp As Variant
    For i = 1 To n - 1
        For j = i + 1 To n
            maxI = IIf(arr(i)(0) > arr(i)(1), arr(i)(0), arr(i)(1))
            maxJ = IIf(arr(j)(0) > arr(j)(1), arr(j)(0), arr(j)(1))
            If maxJ > maxI Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next
    Next
    Do While aList.Count > 0: aList.Remove 1: Loop
    For i = 1 To n: aList.Add arr(i): Next
End Sub

Private Function GetBlockSignature(Block As Collection) As String
    Dim arr() As String, i As Long
    ReDim arr(1 To Block.Count)
    For i = 1 To Block.Count
        arr(i) = Block(i)(0) & "x" & Block(i)(1) & "@" & Block(i)(3) & "@" & Block(i)(4)
    Next i
    ' Сортировка для уникальности подписи
    Dim j As Long, tmpStr As String
    For i = 1 To Block.Count - 1
        For j = i + 1 To Block.Count
            If arr(i) > arr(j) Then
                tmpStr = arr(i)
                arr(i) = arr(j)
                arr(j) = tmpStr
            End If
        Next j
    Next i
    GetBlockSignature = Join(arr, "|")
End Function
