Option Explicit
  '------------------------------------------------------------------------------
  ' Macro: GenerateCuttingPlan v3.7 
  '------------------------------------------------------------------------------
  Public Sub GenerateCuttingPlan()
      Const SRC_SHEET As String = "ИсходныеДанные"
      Const PAR_SHEET As String = "Параметры"
      Const DST_SHEET As String = "Раскрой Древесины"
      Const FIRST_ROW As Long = 12
      Const STOCK_MM As Long = 6000
      Const EXTRA_PT As Double = 4
  
    Dim wsS As Worksheet, wsD As Worksheet, wsP As Worksheet, answer As VbMsgBoxResult
    Dim errorLog As Collection
    Set errorLog = New Collection
    On Error Resume Next
    Set wsS = Worksheets(SRC_SHEET)
    Set wsP = Worksheets(PAR_SHEET)
    On Error GoTo 0
    If wsS Is Nothing Or wsP Is Nothing Then
        MsgBox "Отсутствует лист данных или параметров", vbCritical: Exit Sub
    End If

    Dim KERF_MM As Double: KERF_MM = Val(wsP.Range("F2").Value)
    If KERF_MM <= 0 Then MsgBox "Неверная ширина реза в Параметры!F2", vbCritical: Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' --- Вместо удаления: только очистка или создание если нет ---
    On Error Resume Next
    Set wsD = Worksheets(DST_SHEET)
    On Error GoTo 0

    answer = MsgBox("Перезаписать уже существующий раскрой?", vbYesNo + vbQuestion, "Подтверждение")
    If answer = vbNo Then Exit Sub
    If wsD Is Nothing Then
        Set wsD = Worksheets.Add(After:=wsS)
        wsD.Name = DST_SHEET
    Else
        wsD.Cells.Clear
    End If

    wsD.Activate
    Application.DisplayAlerts = True
      
  
      '--- Общее оформление ---
      wsD.Cells.Font.Name = "Calibri"
      wsD.Cells.HorizontalAlignment = xlCenter
      wsD.Cells.VerticalAlignment = xlCenter
      wsD.Columns("A").ColumnWidth = 6
      wsD.Columns("B").ColumnWidth = 8
      wsD.Columns("C").ColumnWidth = 100
      wsD.Columns("D").ColumnWidth = 10
      wsD.Columns("E").ColumnWidth = 10
      wsD.Columns("F").ColumnWidth = 10
      wsD.Columns("G:I").ColumnWidth = 2
      wsD.Columns("M:P").ColumnWidth = 2
  
      '--- Масштаб (pt/мм) для отрисовки ---
      wsD.Cells(1, 3) = "x": wsD.Cells(1, 4) = "x"
      Dim scalePt As Double
      scalePt = (wsD.Cells(1, 4).Left - wsD.Cells(1, 3).Left - EXTRA_PT) / STOCK_MM
      wsD.rows(1).Clear
  
      '--- Собираем детали по ключу "слой|сечение" ---
      Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
      Dim blockStats As Object: Set blockStats = CreateObject("Scripting.Dictionary")

      ' summary per section
      Dim summaryDict As Object: Set summaryDict = CreateObject("Scripting.Dictionary")
      Dim lastR As Long: lastR = wsS.Cells(wsS.rows.Count, "C").End(xlUp).row
      Dim r As Long, layer As String, sect As String, l As Long, qty As Long, k As Long, comboKey As String, sectionLen As Long, paramRow As Variant
      For r = FIRST_ROW To lastR
        layer = Trim(wsS.Cells(r, "C").Value)
        sect = Trim(wsS.Cells(r, "D").Value)
        l = Val(wsS.Cells(r, "E").Value)
        qty = Val(wsS.Cells(r, "F").Value)
        Dim reason As String
        reason = ""
        ' --- Фильтр: не учитываем детали длиной 10 мм и меньше ---
        If l <= 10 Then GoTo SkipRow
        If layer = "" Or sect = "" Then
            reason = "Пустой слой или сечение"
        ElseIf l > STOCK_MM Then
            reason = "Длина больше стандартной доски"
        ElseIf qty <= 0 Then
            reason = "Количество ≤ 0"
        End If
        If reason <> "" Then
            Dim szStr As String
            szStr = wsS.Cells(r, "D").Value & "x" & wsS.Cells(r, "E").Value
            errorLog.Add Array(wsS.Cells(r, "A").Value, szStr, reason, qty, r)
        Else
            comboKey = layer & "|" & sect
            If Not dict.Exists(comboKey) Then Set dict(comboKey) = New Collection
            For k = 1 To qty: dict(comboKey).Add l: Next k
        End If
SkipRow:
    Next r

      If dict.Count = 0 Then MsgBox "Подходящих деталей не найдено", vbExclamation: Exit Sub
  
      '--- Палитра и цвета ---
      Dim palette(1 To 4) As Long
      palette(1) = RGB(192, 96, 96): palette(2) = RGB(224, 192, 96)
      palette(3) = RGB(96, 192, 96): palette(4) = RGB(96, 192, 192)
      Const COL_OFFCUT As Long = &H808080, COL_OUTLINE As Long = &H404040, COL_FRAME As Long = vbBlack, COL_HDR As Long = &HC0C0C0
  
      '--- Результат ---
      Dim curRow As Long: curRow = 1
      Dim ck As Variant
      For Each ck In dict.Keys
          layer = Split(ck, "|")(0): sect = Split(ck, "|")(1)
          '--- определяем стандартную длину доски для данного сечения (приоритет из Параметры!A:B) ---
          sectionLen = STOCK_MM
          On Error Resume Next
          sectionLen = CLng(Application.WorksheetFunction.VLookup(sect, wsP.Range("A:B"), 2, False))
          On Error GoTo 0
          If sectionLen <= 0 Then sectionLen = STOCK_MM
          ' Заголовок блока (слой + сечение)
          ' -- разбиваем на две объединенные области: A:C и D:F
          With wsD.Range(wsD.Cells(curRow, 1), wsD.Cells(curRow, 6))
              .Merge: .Value = layer & " " & sect
              .Interior.Color = COL_HDR: .Font.Bold = True
              .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
              .Borders.LineStyle = xlContinuous: .Borders.Color = COL_FRAME: .Borders.Weight = xlThin
          End With
              
          wsD.rows(curRow).RowHeight = 15
          curRow = curRow + 1
  
          ' --- шапка столбцов ---
          Dim hdrArr: hdrArr = Array("№", "Длина мм", "Карта раскроя (ширина реза " & KERF_MM & " мм )", _
                                   "Кол-во целых шт", "Сечение", "Остаток мм")
          Dim jj As Long
          For jj = 1 To 6
              wsD.Cells(curRow, jj) = hdrArr(jj - 1)
              With wsD.Range(wsD.Cells(curRow, jj), wsD.Cells(curRow + 1, jj))
                  .Merge: .WrapText = True: .Font.Bold = True
                  .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
                  .Borders.LineStyle = xlContinuous
                  .Borders.Color = COL_FRAME
                  .Borders.Weight = xlThin
              End With
          Next jj
          wsD.rows(curRow).RowHeight = 15: wsD.rows(curRow + 1).RowHeight = 15
          curRow = curRow + 2
  
          ' Уникальные постройки досок в блоке
          Dim cuts As Object: Set cuts = dict(ck)

            ' --- Сортировка элементов cuts (Collection) по убыванию ---
            Dim arrCuts() As Long
            Dim i As Long
            Dim L2 As Variant

            ReDim arrCuts(1 To cuts.Count)
            i = 1
            For Each L2 In cuts
                arrCuts(i) = CLng(L2)
                i = i + 1
            Next

            ' Сортировка по убыванию (от большего к меньшему)
            Dim j As Long, tmp As Long
            For i = 1 To UBound(arrCuts) - 1
                For j = i + 1 To UBound(arrCuts)
                    If arrCuts(i) < arrCuts(j) Then
                        tmp = arrCuts(i): arrCuts(i) = arrCuts(j): arrCuts(j) = tmp
                    End If
                Next j
            Next

            ' Обратно в Collection
            Set cuts = New Collection
            For i = 1 To UBound(arrCuts)
                cuts.Add arrCuts(i)
            Next

          Dim boardDict As Object: Set boardDict = CreateObject("Scripting.Dictionary")
          Do While cuts.Count > 0
              Dim usedMm As Long: usedMm = 0
              Dim partsUsed As String: partsUsed = ""
              Dim remain As Object: Set remain = New Collection
              For Each L2 In cuts
                  If usedMm + IIf(usedMm > 0, KERF_MM, 0) + CLng(L2) <= sectionLen Then
                      usedMm = usedMm + IIf(usedMm > 0, KERF_MM, 0) + CLng(L2)
                      partsUsed = partsUsed & CLng(L2) & "-"
                  Else
                      remain.Add L2
                  End If
              Next L2
              Dim rest As Long: rest = sectionLen - usedMm
              Dim patt As String: patt = partsUsed & "|" & rest
              If boardDict.Exists(patt) Then
                  boardDict(patt) = boardDict(patt) + 1
              Else
                  boardDict(patt) = 1
              End If
              Set cuts = remain
          Loop
  
          ' Вывод уникальных карт
          Dim pattKey As Variant, copies As Long, boardNo As Long: boardNo = 1
          Dim colorIdx As Integer, pVal As Variant, leftPos As Double, boardScale As Double, restTotal As Long
          restTotal = 0
            Dim boardTotal As Long: boardTotal = 0 ' Общее кол-во досок (сумма copies)
            For Each pattKey In boardDict.Keys
                copies = boardDict(pattKey)
                partsUsed = Split(pattKey, "|")(0)
                rest = CLng(Split(pattKey, "|")(1))
              ' Тонкая черная обводка ячеек
              With wsD.Range(wsD.Cells(curRow, 1), wsD.Cells(curRow, 6)).Borders
                  .LineStyle = xlContinuous: .Color = COL_FRAME: .Weight = xlThin
              End With
              ' Центрируем содержимое ячеек по вертикали и горизонтали
              With wsD.Range(wsD.Cells(curRow, 1), wsD.Cells(curRow, 6))
              .HorizontalAlignment = xlCenter
              .VerticalAlignment = xlCenter
          End With
          
  
              wsD.Cells(curRow, 1) = boardNo: wsD.Cells(curRow, 2) = sectionLen
              wsD.Cells(curRow, 4) = copies: wsD.Cells(curRow, 5) = sect: wsD.Cells(curRow, 6) = rest
  
              '--- масштаб для текущей доски: вся ширина колонки C на длину sectionLen ---
              boardScale = (wsD.Columns("C").Width - EXTRA_PT) / sectionLen
  
              ' Рисуем карту (только shapes)
              usedMm = 0: colorIdx = 1
              For Each pVal In Split(partsUsed, "-")
                  If Len(pVal) > 0 Then
                      leftPos = wsD.Cells(curRow, 3).Left + (usedMm + IIf(usedMm > 0, KERF_MM, 0)) * boardScale
                      DrawBar wsD, leftPos, wsD.rows(curRow).Top + 0.25, CLng(pVal) * boardScale, wsD.rows(curRow).Height - 0.5, _
                              palette(colorIdx), COL_OUTLINE, CLng(pVal)
                      usedMm = usedMm + IIf(usedMm > 0, KERF_MM, 0) + CLng(pVal)
                      colorIdx = colorIdx Mod 4 + 1
                  End If
              Next pVal
              If rest > 0 Then
                  DrawBar wsD, wsD.Cells(curRow, 3).Left + usedMm * boardScale, wsD.rows(curRow).Top + 0.25, rest * boardScale, _
                          wsD.rows(curRow).Height - 0.5, COL_OFFCUT, COL_OUTLINE, ""
              End If
  
                curRow = curRow + 1
                boardNo = boardNo + 1
                restTotal = restTotal + rest * copies
                boardTotal = boardTotal + copies

            Next pattKey
            '--- Итоговая строка блока ---
            With wsD.Range(wsD.Cells(curRow, 1), wsD.Cells(curRow, 6)).Borders
                .LineStyle = xlContinuous: .Color = COL_FRAME: .Weight = xlThin
            End With
            With wsD.Range(wsD.Cells(curRow, 1), wsD.Cells(curRow, 6))
                .Interior.Color = RGB(230, 230, 230)   ' светло-серый фон
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous: .Borders.Color = COL_FRAME: .Borders.Weight = xlThin
            End With
            wsD.Cells(curRow, 3) = "Итог:"
            wsD.Cells(curRow, 3).HorizontalAlignment = xlRight
            wsD.Cells(curRow, 4) = boardTotal      ' кол-во досок в блоке (сумма всех copies)
            wsD.Cells(curRow, 5) = sect             ' сечение
            wsD.Cells(curRow, 6) = restTotal        ' общий остаток, мм
            curRow = curRow + 0

            Application.ScreenUpdating = True
            wsD.Cells(curRow, 1).Select   ' Можно и без .Select, если не нужен переход курсора
            DoEvents                     ' Чтобы экран точно обновился
            Application.ScreenUpdating = False

            Dim blockCount As Long: blockCount = boardTotal
            If Not summaryDict.Exists(sect) Then
                summaryDict(sect) = blockCount
            Else
                summaryDict(sect) = summaryDict(sect) + blockCount
            End If
            blockStats.Add ck, Array(blockCount, restTotal)
            curRow = curRow + 1
            
        Next ck
        '--- Таблица итоговых данных ---
        Dim tblCol As Long: tblCol = wsD.Columns("Q").Column
        Dim tblRow As Long: tblRow = 1
        Dim vol As Double
        Dim volFull As Double
        ' Заголовки таблицы
        Dim hdrs As Variant: hdrs = Array("Слой", "Ширина", "Высота", "Длина", "Кол-во", "Объем м^3", "Объем ост.м^3", "Категория", "Покраска", "Цвет", "Назначение", "Материал")
        Dim hi As Long
        For hi = 0 To UBound(hdrs)
            With wsD.Cells(tblRow, tblCol + hi)
                .Value = hdrs(hi)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous: .Borders.Color = COL_FRAME: .Borders.Weight = xlThin
            End With
        Next hi
        tblRow = tblRow + 1

            ' Заполняем строки по блокам
        Dim key As Variant
        Dim purpose As String
        For Each key In blockStats.Keys
            Dim dimsBlk As Variant
            Dim parts() As String: parts = Split(key, "|")
            Dim lay As String: lay = parts(0)
            Dim sec As String: sec = parts(1)
            ' Стандартная длина доски
            Dim stdLen As Long: stdLen = STOCK_MM
            On Error Resume Next
            stdLen = CLng(Application.WorksheetFunction.VLookup(sec, wsP.Range("A:B"), 2, False))
            On Error GoTo 0
            If stdLen <= 0 Then stdLen = STOCK_MM
            ' Данные по блоку
            Dim stats As Variant: stats = blockStats(key)
            Dim cnt As Long: cnt = stats(0)
            Dim remMm As Long: remMm = stats(1)
            dimsBlk = Split(sec, "x")
            ' Объем остатка (м³)
            vol = (remMm / 1000) * (CDbl(dimsBlk(0)) / 1000) * (CDbl(dimsBlk(1)) / 1000)
            ' Категория и покраска
            Dim mRow As Variant: mRow = Application.Match(lay, wsP.Range("S:S"), 0)
            Dim cat As String, paint As String
            If Not IsError(mRow) Then
                cat = wsP.Cells(mRow, "T").Value
                paint = wsP.Cells(mRow, "U").Value
                purpose = wsP.Cells(mRow, "V").Value
            Else
                purpose = ""
            End If

          ' Запись строки
            wsD.Cells(tblRow, tblCol) = lay
            wsD.Cells(tblRow, tblCol + 1) = dimsBlk(0)
            wsD.Cells(tblRow, tblCol + 2) = dimsBlk(1)
            wsD.Cells(tblRow, tblCol + 3) = stdLen
            wsD.Cells(tblRow, tblCol + 4) = cnt
            volFull = cnt * (CDbl(dimsBlk(0)) / 1000) * (CDbl(dimsBlk(1)) / 1000) * (stdLen / 1000)
            wsD.Cells(tblRow, tblCol + 5) = Round(volFull, 3)
            wsD.Cells(tblRow, tblCol + 6) = Round(vol, 3)

            ' --- выпадающий список для Категории (X)
            Dim lastK As Long: lastK = wsP.Cells(wsP.Rows.Count, "K").End(xlUp).Row
            With wsD.Cells(tblRow, tblCol + 7)
                .Value = cat
                .Validation.Delete
                .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                    Formula1:="='" & PAR_SHEET & "'!$K$2:$K$" & lastK
                .Validation.InCellDropdown = True
                .Validation.InputMessage = "Выберите категорию применения"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Color = COL_FRAME
                .Borders.Weight = xlThin
        

        ' --- условное форматирование: штриховка если не выбрано
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlExpression, Formula1:="=" & .Address(True, True) & "="""""
            With .FormatConditions(.FormatConditions.Count).Interior
                .Pattern = xlPatternLightDown
                .PatternColor = RGB(200, 100, 140)
            End With
        End With

            ' --- выпадающий список для Покраски (Y)
            Dim lastL As Long: lastL = wsP.Cells(wsP.Rows.Count, "L").End(xlUp).Row
            With wsD.Cells(tblRow, tblCol + 8)
                .Value = paint
                .Validation.Delete
                .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                    Formula1:="='" & PAR_SHEET & "'!$L$2:$L$" & lastL
                .Validation.InCellDropdown = True
                .Validation.InputMessage = "Выберите вариант покраски"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Color = COL_FRAME
                .Borders.Weight = xlThin
           

            ' --- условное форматирование: штриховка если не выбрано
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlExpression, Formula1:="=" & .Address(True, True) & "="""""
            With .FormatConditions(.FormatConditions.Count).Interior
                .Pattern = xlPatternLightDown
                .PatternColor = RGB(200, 100, 140)
            End With
        End With

            ' Цвет: текстовое значение из ИсходныеДанные!K
            Dim colorText As String: colorText = ""
            Dim rr As Long
            For rr = FIRST_ROW To lastR
                If Trim(wsS.Cells(rr, "C").Value) = lay And Trim(wsS.Cells(rr, "D").Value) = sec Then
                    colorText = wsS.Cells(rr, "K").Text
                    Exit For
                End If
            Next rr
            wsD.Cells(tblRow, tblCol + 9) = colorText

            ' --- Поиск названия материала
            Dim materialName As String: materialName = ""
            For rr = FIRST_ROW To lastR
                If Trim(wsS.Cells(rr, "C").Value) = lay And Trim(wsS.Cells(rr, "D").Value) = sec Then
                    materialName = wsS.Cells(rr, "L").Text
                    Exit For
                End If
            Next rr
            wsD.Cells(tblRow, tblCol + 11) = materialName

            ' --- выпадающий список для Назначения (AA) ---
            Dim lastM As Long: lastM = wsP.Cells(wsP.Rows.Count, "M").End(xlUp).Row
            With wsD.Cells(tblRow, tblCol + 10)
                .Value = purpose
                .Validation.Delete
                .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                    Formula1:="='" & PAR_SHEET & "'!$M$2:$M$" & lastM
                .Validation.InCellDropdown = True
                .Validation.InputMessage = "Выберите назначение"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Color = COL_FRAME
                .Borders.Weight = xlThin

                ' Штриховка, если пусто
                .FormatConditions.Delete
                .FormatConditions.Add Type:=xlExpression, Formula1:="=" & .Address(True, True) & "="""""
                With .FormatConditions(.FormatConditions.Count).Interior
                    .Pattern = xlPatternLightDown
                    .PatternColor = RGB(200, 100, 140)
                End With
            End With


            With wsD.Range(wsD.Cells(tblRow, tblCol), wsD.Cells(tblRow, tblCol + 11)).Borders
                .LineStyle = xlContinuous: .Color = COL_FRAME: .Weight = xlThin
            End With
            tblRow = tblRow + 1
        Next key

        ' --- Оформление итоговой таблицы ---
        With wsD.Columns("Q")     ' Слой
        .ColumnWidth = 48
        .HorizontalAlignment = xlLeft
        End With
        With wsD.Columns("U")     ' Кол-во
            .ColumnWidth = 7
        End With
        With wsD.Columns("V")     ' Объем м³
            .ColumnWidth = 15
        End With
        With wsD.Columns("W")     ' Объем остатков
            .ColumnWidth = 15
        End With
        With wsD.Columns("X")     ' Категория
            .ColumnWidth = 38
        End With
        With wsD.Columns("Y")     ' Покраска
            .ColumnWidth = 12
        End With
        With wsD.Columns("Z")     ' Цвет
            .ColumnWidth = 44
            .HorizontalAlignment = xlLeft
        End With
        With wsD.Columns("AA")     ' Назначение
            .ColumnWidth = 12
        End With
        With wsD.Columns("AB") ' Материал
            .ColumnWidth = 22
            .HorizontalAlignment = xlLeft
        End With


        '--- вывод сводной таблицы через одну колонку справа
        Dim summRow As Long: summRow = 1
            With wsD.Range(wsD.Cells(summRow, "J"), wsD.Cells(summRow, "L"))
                .Merge
                .Value = "Сводная таблица"
                .Interior.Color = COL_HDR
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            wsD.rows(summRow).RowHeight = 15
            ' --- Заголовки столбцов сводки (2 строки) ---
            Dim hdrSumm As Variant: hdrSumm = Array("Сечение", "Кол-во досок", "Объем, м^2")
            Dim sCol As Integer
            For sCol = 0 To 2
                wsD.Cells(summRow + 1, 10 + sCol) = hdrSumm(sCol)
                With wsD.Range(wsD.Cells(summRow + 1, 10 + sCol), wsD.Cells(summRow + 2, 10 + sCol))
                    .Merge: .WrapText = True: .Font.Bold = True
                    .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
                End With
            Next sCol
            wsD.rows(summRow + 1).RowHeight = 15: wsD.rows(summRow + 2).RowHeight = 15
            summRow = summRow + 3
        Dim keySect As Variant
        For Each keySect In summaryDict.Keys
            wsD.Cells(summRow, "J") = keySect
            wsD.Cells(summRow, "K") = summaryDict(keySect)
            ' calculate volume m3: width x height x length
            Dim dimsSummary As Variant: dimsSummary = Split(keySect, "x")
            Dim boardLen As Long: boardLen = STOCK_MM
            On Error Resume Next
            boardLen = CLng(Application.WorksheetFunction.VLookup(keySect, wsP.Range("A:B"), 2, False))
            On Error GoTo 0
            If boardLen <= 0 Then boardLen = STOCK_MM
            vol = summaryDict(keySect) * (CDbl(dimsSummary(0)) / 1000) * (CDbl(dimsSummary(1)) / 1000) * (boardLen / 1000)
            wsD.Cells(summRow, "L") = Round(vol, 3)
            summRow = summRow + 1
        Next keySect
        '--- тонкая черная обводка для сводной таблицы ---
        Dim summaryStart As Long: summaryStart = 1
        Dim summaryEnd As Long: summaryEnd = summRow - 1
        With wsD.Range(wsD.Cells(summaryStart, "J"), wsD.Cells(summaryEnd, "L")).Borders
            .LineStyle = xlContinuous: .Color = COL_FRAME: .Weight = xlThin
        End With
    
        Application.ScreenUpdating = True
        ' --- наношу рамки на всю таблицу ---
        Dim lastOutRow As Long: lastOutRow = curRow - 1
        With wsD.Range(wsD.Cells(1, 1), wsD.Cells(lastOutRow, 6)).Borders
            .LineStyle = xlContinuous: .Color = COL_FRAME: .Weight = xlThin
        End With

        If errorLog.Count > 0 Then
            Dim logStartRow As Long
            Dim logIdx As Long
            logStartRow = lastOutRow + 3
            
            With wsD.Range(wsD.Cells(logStartRow, 1), wsD.Cells(logStartRow, 4))
                .Merge
                .Value = "Лог непринятых деталей"
                .Font.Bold = True
                .Interior.Color = RGB(255, 224, 224)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Color = COL_FRAME
                .Borders.Weight = xlThin
            End With

            wsD.Cells(logStartRow + 1, 1).Value = "№ ошибки"
            wsD.Cells(logStartRow + 1, 2).Value = "№ детали"
            wsD.Cells(logStartRow + 1, 3).Value = "Причина"
            wsD.Cells(logStartRow + 1, 4).Value = "Кол-во"
            wsD.Rows(logStartRow + 1).RowHeight = 30
            wsD.Rows(logStartRow + 1).WrapText = True


            For logIdx = 1 To errorLog.Count
                Dim errArr As Variant
                errArr = errorLog(logIdx)
                wsD.Cells(logStartRow + 1 + logIdx, 1).Value = logIdx
                wsD.Cells(logStartRow + 1 + logIdx, 2).Value = errArr(0)
                wsD.Cells(logStartRow + 1 + logIdx, 4).Value = errArr(3)

                ' --- Добавляем гиперссылку ---
                Dim srcRowNum As Long
                srcRowNum = errArr(4) ' номер строки в исходных данных
                Dim linkText As String
                linkText = errArr(1) & " : " & errArr(2)

                ' Если строка валидная — делаем гиперссылку
                If srcRowNum > 0 Then
                    wsD.Hyperlinks.Add _
                        Anchor:=wsD.Cells(logStartRow + 1 + logIdx, 3), _
                        Address:="", _
                        SubAddress:="'" & SRC_SHEET & "'!A" & srcRowNum, _
                        TextToDisplay:=linkText
                Else
                    wsD.Cells(logStartRow + 1 + logIdx, 3).Value = linkText
                End If
            Next logIdx

            With wsD.Range(wsD.Cells(logStartRow + 1, 1), wsD.Cells(logStartRow + 1 + errorLog.Count, 4)).Borders
                .LineStyle = xlContinuous: .Color = COL_FRAME: .Weight = xlThin
            End With
        End If
       
    MsgBox "Раскрой сформирован (v3.7)", vbInformation
    Application.Goto wsD.Range("Q1"), True
End Sub
  
Private Sub DrawBar(ws As Worksheet, l As Double, t As Double, w As Double, h As Double, _
                    fillRGB As Long, lineRGB As Long, txt As Variant)
    Dim s As Shape: Set s = ws.Shapes.AddShape(msoShapeRectangle, l, t, w, h)
    With s
        .Fill.Solid: .Fill.ForeColor.RGB = fillRGB
        .Line.Visible = msoTrue: .Line.ForeColor.RGB = lineRGB: .Line.Weight = 0.25
        If Len(txt) > 0 Then
            .TextFrame2.TextRange.Text = CStr(txt)
            .TextFrame2.TextRange.Font.Size = 14: .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.HorizontalAnchor = msoAnchorCenter: .TextFrame2.VerticalAnchor = msoAnchorMiddle
        End If
    End With
End Sub
