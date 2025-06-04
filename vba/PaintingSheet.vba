Sub GeneratePaintingShippingReport_Full()
    Dim wsOut As Worksheet, wsPar As Worksheet
    Dim wsSrc As Worksheet, wsList As Worksheet
    Dim lastRow As Long, outRow As Long, i As Long, j As Long
    Dim layer As String, plate As String
    Dim panelGroups As Object, shippingGroups As Object
    Dim layerMatch As Boolean
    Dim srcSheets As Variant
    Dim layerPanel As Variant, layerShip As Variant
    Dim sheetName As Variant, v As Variant

    ' === Используем существующий шаблонный лист ===
    Set wsOut = Worksheets("Покраска")
    Set wsPar = Worksheets("Параметры")

    ' --- Очистка старых данных (только строки ниже шапки, не трогая оформление/шапку) ---
    wsOut.Rows("6:" & wsOut.Rows.Count).ClearContents
    wsOut.Rows("6:" & wsOut.Rows.Count).Interior.ColorIndex = xlNone
    wsOut.Rows("6:" & wsOut.Rows.Count).Borders.LineStyle = xlNone



    ' --- Определяем листы с итоговыми таблицами ---
    Dim logArr() As String
    Dim logReasonArr() As String
    Dim logCount As Long
    logCount = 0

    srcSheets = Array("Раскрой Древесины")
    Set panelGroups = CreateObject("Scripting.Dictionary")
    Set shippingGroups = CreateObject("Scripting.Dictionary")

    For Each sheetName In srcSheets
        On Error Resume Next
        Set wsList = Worksheets(sheetName)
        If wsList Is Nothing Then On Error GoTo 0: GoTo NextSheet
        lastRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
        ' Все значения сразу из строки
        Dim paintFlag As String, catVal As String, layerName As String
        paintFlag = Trim(wsList.Cells(i, 25).Value)   ' Y — красим/не красим
        CatVal = Trim(wsList.Cells(i, 27).Value)      ' AA — категория
        layerName = wsList.Cells(i, 17).Value         ' Q — слой (для логов)

        ' Пустой слой — пропуск
        If Len(Trim(layerName)) = 0 Then GoTo NextRow

        Dim layerReason As String

        If paintFlag = "НЕ красим" Then
            layerReason = "НЕ красим"
            GoTo AddToLog
        ElseIf paintFlag = "Красим" Then
        Else
            layerReason = "Не выбрана покраска"
            GoTo AddToLog
        End If

        If CatVal = "Панели" Then
            panelGroups(panelGroups.Count) = Array(sheetName, i)
        ElseIf CatVal = "Отправка" Then
            shippingGroups(shippingGroups.Count) = Array(sheetName, i)
        Else
            layerReason = "Категория не Панели/Отправка"
            GoTo AddToLog
        End If

        GoTo NextRow

    AddToLog:
        sizeStr = wsList.Cells(i, 18).Value & "x" & wsList.Cells(i, 19).Value & "x" & wsList.Cells(i, 20).Value
        logCount = logCount + 1
        ReDim Preserve logArr(1 To logCount)
        ReDim Preserve logReasonArr(1 To logCount)
        logArr(logCount) = layerName & " " & sizeStr
        logReasonArr(logCount) = layerReason

    NextRow:
    Next i

NextSheet:
        Set wsList = Nothing
    Next

    ' --- Формируем данные после шапки ---
    outRow = 6

    Dim groupTypes(1) As Object
    Dim groupTitles(1) As String
    Dim currGroup As Object
    Dim sh As Double, vh As Double, dl As Double

    Set groupTypes(0) = panelGroups
    Set groupTypes(1) = shippingGroups
    groupTitles(0) = "На панели"
    groupTitles(1) = "На отправку"

    For k = 0 To 1
    Set currGroup = groupTypes(k)
    If currGroup.Count > 0 Then
        ' Объединить ячейки по всей ширине таблицы для названия группы
        With wsOut.Range(wsOut.Cells(outRow, 2), wsOut.Cells(outRow, 21))
            .Merge
            .Value = groupTitles(k)
            .Font.Bold = True
            .Interior.Color = RGB(220, 220, 220)   ' Светло-серый
        End With
        outRow = outRow + 1


            For j = 0 To currGroup.Count - 1
            Set wsList = Worksheets(currGroup(j)(0))
            i = currGroup(j)(1)
            sizeStr = wsList.Cells(i, 18).Value & "x" & wsList.Cells(i, 19).Value & "x" & wsList.Cells(i, 20).Value
            Dim widthStr As String, heightStr As String, lengthStr As String, catStr As String
            widthStr = wsList.Cells(i, 18).Value
            heightStr = wsList.Cells(i, 19).Value
            lengthStr = wsList.Cells(i, 20).Value
            catStr = wsList.Cells(i, 24).Value ' Если категория в X, иначе скорректируйте по вашей структуре
            wsOut.Cells(outRow, 2).Value = sizeStr              ' B — Доска

            ' Динамический расчет объема элемента (пример — если в "Раскрой Древесины"!V находится объем)
            Dim lastSrcRow As Long
            lastSrcRow = Worksheets("Раскрой Древесины").Cells(Rows.Count, "A").End(xlUp).Row
            wsOut.Cells(outRow, 3).Formula = _
                "=SUMIFS('Раскрой Древесины'!V2:V" & lastSrcRow & "," & _
                        "'Раскрой Древесины'!R2:R" & lastSrcRow & ",""" & widthStr & """," & _
                        "'Раскрой Древесины'!S2:S" & lastSrcRow & ",""" & heightStr & """," & _
                        "'Раскрой Древесины'!T2:T" & lastSrcRow & ",""" & lengthStr & """," & _
                        "'Раскрой Древесины'!X2:X" & lastSrcRow & ",""" & catStr & """)"
            wsOut.Cells(outRow, 4).Value = "" ' D — выпадающий список, не трогаем

            ' Аналогично для J (10 колонка) — количество доски (пример, если U — количество)
            wsOut.Cells(outRow, 10).Formula = _
                "=SUMIFS('Раскрой Древесины'!U2:U" & lastSrcRow & "," & _
                        "'Раскрой Древесины'!R2:R" & lastSrcRow & ",""" & widthStr & """," & _
                        "'Раскрой Древесины'!S2:S" & lastSrcRow & ",""" & heightStr & """," & _
                        "'Раскрой Древесины'!T2:T" & lastSrcRow & ",""" & lengthStr & """," & _
                        "'Раскрой Древесины'!X2:X" & lastSrcRow & ",""" & catStr & """)"

           wsOut.Cells(outRow, 12).Formula = "=IF('Раскрой Древесины'!S" & i & "<'Раскрой Древесины'!R" & i & "," & _
                "('Раскрой Древесины'!T" & i & "*'Раскрой Древесины'!S" & i & "*2 + 'Раскрой Древесины'!T" & i & "*'Раскрой Древесины'!R" & i & ")/1000000," & _
                "('Раскрой Древесины'!T" & i & "*'Раскрой Древесины'!R" & i & "*2 + 'Раскрой Древесины'!T" & i & "*'Раскрой Древесины'!S" & i & ")/1000000)"

            ' Остальные поля оставьте как есть, либо аналогично заменяйте на формулы если нужен динамический расчет
            
            wsOut.Cells(outRow, 21).Value = wsList.Cells(i, 26).Value  ' U — Оттенок
            wsOut.Rows(outRow).ClearFormats

            outRow = outRow + 1
        Next j

        End If
    Next k

    ' === ДОБАВЛЯЕМ ВЫПАДАЮЩИЙ СПИСОК ===
    Dim parLastRow As Long
    parLastRow = wsPar.Cells(wsPar.Rows.Count, "AX").End(xlUp).Row
    Dim paintListRange As String
    paintListRange = "='Параметры'!$AX$2:$AX$" & parLastRow

    Dim firstDataRow As Long, lastDataRow As Long
    firstDataRow = 6
    lastDataRow = outRow - 1

    With wsOut.Range(wsOut.Cells(firstDataRow, 4), wsOut.Cells(lastDataRow, 4)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:=paintListRange
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

    Dim r As Long
    For r = firstDataRow To lastDataRow
        With wsOut.Cells(r, 4)
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlExpression, Formula1:="=" & .Address(False, False) & "="""""
            With .FormatConditions(.FormatConditions.Count).Interior
                .Pattern = xlPatternLightDown
                .PatternColor = RGB(200, 100, 140)
            End With
        End With
    Next r


    ' === ФОРМУЛА РАСХОДА В E ===
    Dim formulaStr As String
    formulaStr = "=IF(D6<>"""",ROUND(IFERROR(VLOOKUP(D6,'Параметры'!AX:AY,2,FALSE)*L6*E5,""""),3),"""")"
    For r = firstDataRow To lastDataRow
        wsOut.Cells(r, 5).Formula = Replace(formulaStr, "6", r)
    Next r

    ' === ОФОРМЛЕНИЕ ===
    With wsOut
        Dim ranges As Variant, rr As Variant
        ranges = Array( _
            "E6:H" & outRow - 1, _
            "I6:K" & outRow - 1, _
            "L6:N" & outRow - 1, _
            "T6:T" & outRow - 1, _
            "U6:U" & outRow - 1, _
            "O6:S" & outRow - 1, _
            "B6:D" & outRow - 1)

        .Range(ranges(0)).Interior.Color = RGB(202, 220, 231)
        .Range(ranges(1)).Interior.Color = RGB(187, 217, 187)
        .Range(ranges(2)).Interior.Color = RGB(217, 185, 185)
        .Range(ranges(3)).Interior.Color = RGB(255, 250, 214)
        .Range(ranges(4)).Interior.Color = RGB(238, 244, 255)

        For Each rr In ranges
            With .Range(rr)
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).Weight = xlThin
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).Weight = xlThin
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlMedium
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlMedium
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlMedium
            End With
        Next rr
    End With

    Dim logStartRow As Long
    For i = 1 To 3
        wsOut.Rows(outRow).ClearFormats
        outRow = outRow + 1
    Next i
    logStartRow = outRow
    If logCount > 0 Then
        wsOut.Cells(logStartRow, 2).Value = "ЛОГ: Не попали в покраску"
        wsOut.Cells(logStartRow, 2).Font.Bold = True
        wsOut.Range(wsOut.Cells(logStartRow, 2), wsOut.Cells(logStartRow, 10)).Interior.Color = RGB(255, 224, 224)
        For i = 1 To logCount
            wsOut.Rows(logStartRow + i).ClearFormats
            wsOut.Cells(logStartRow + i, 2).Value = logArr(i)
            wsOut.Range(wsOut.Cells(logStartRow + i, 8), wsOut.Cells(logStartRow + i, 10)).Merge
            wsOut.Cells(logStartRow + i, 8).Value = logReasonArr(i)
        Next i
    End If

    MsgBox "Лист 'Покраска' сформирован!", vbInformation
End Sub
