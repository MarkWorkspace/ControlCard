' Макрос для создания Расчётного листа с обработкой ошибок типа данных
Sub CreateCalculationSheet()
    Dim wb As Workbook
    Dim wsSrc As Worksheet, wsParam As Worksheet, wsDest As Worksheet
    Dim lastRowSrc As Long, lastRowParam As Long, lastColSrc As Long
    Dim catList As Collection
    Dim matDict As Object
    Dim i As Long, j As Long
    Dim width As Double, height As Double, length As Double
    Dim sizeStr As String
    Dim vol As Double, mass As Double, massPerM3 As Double
    Dim destRow As Long, destCol As Long
    Dim cat As Variant
    Dim key As Variant
    Dim qty As Double

    Set wb = ThisWorkbook
    Set wsSrc = wb.Sheets("Раскрой Древесины")
    Set wsParam = wb.Sheets("Параметры")
    
    ' 1. Получаем список категорий из Параметры!K2:K...
    Set catList = New Collection
    i = 2
    Do While wsParam.Cells(i, "K").Value <> ""
        catList.Add wsParam.Cells(i, "K").Value
        i = i + 1
    Loop
    
    ' 2. Масса 1 м3 из AF2
    massPerM3 = 0
    If IsNumeric(wsParam.Range("AF2").Value) Then
        massPerM3 = CDbl(wsParam.Range("AF2").Value)
    End If
    
    ' 3. Собираем пересечения [размер, категория] из итоговой таблицы + количество (шт.)
    Set matDict = CreateObject("Scripting.Dictionary")
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, "Q").End(xlUp).Row

    For i = 2 To lastRowSrc
        ' Безопасное чтение чисел с обработкой пустых и ошибочных значений
        If IsNumeric(wsSrc.Cells(i, "R").Value) Then
            width = CDbl(wsSrc.Cells(i, "R").Value)
        Else
            width = 0
        End If
        If IsNumeric(wsSrc.Cells(i, "S").Value) Then
            height = CDbl(wsSrc.Cells(i, "S").Value)
        Else
            height = 0
        End If
        If IsNumeric(wsSrc.Cells(i, "T").Value) Then
            length = CDbl(wsSrc.Cells(i, "T").Value)
        Else
            length = 0
        End If
        sizeStr = width & "x" & height & "x" & length
        cat = wsSrc.Cells(i, "X").Value ' Категория — это колонка X
        If IsNumeric(wsSrc.Cells(i, "U").Value) Then
            qty = wsSrc.Cells(i, "U").Value
        Else
            qty = 0
        End If
        If Not IsEmpty(cat) And cat <> "" Then
            If IsNumeric(wsSrc.Cells(i, "V").Value) Then
                vol = wsSrc.Cells(i, "V").Value
            Else
                vol = 0
            End If
            If vol > 0 Or qty > 0 Then
                If Not matDict.Exists(sizeStr) Then
                    Set matDict(sizeStr) = CreateObject("Scripting.Dictionary")
                End If
                If Not matDict(sizeStr).Exists(cat) Then
                    Set matDict(sizeStr)(cat) = CreateObject("Scripting.Dictionary")
                    matDict(sizeStr)(cat)("qty") = 0
                    matDict(sizeStr)(cat)("vol") = 0
                End If
                matDict(sizeStr)(cat)("qty") = matDict(sizeStr)(cat)("qty") + qty
                matDict(sizeStr)(cat)("vol") = matDict(sizeStr)(cat)("vol") + vol
            End If
        End If
    Next i

    ' --- ДОБАВЛЯЕМ ДАННЫЕ ИЗ "Раскрой Плит" ---
    Dim wsPlt As Worksheet
    Dim lastRowPlt As Long
    Dim widthPlt As Double, lengthPlt As Double
    Dim sizePlt As String
    Dim catPlt As String
    Dim qtyPlt As Double

    Set wsPlt = wb.Sheets("Раскрой Плит")
    lastRowPlt = wsPlt.Cells(wsPlt.Rows.Count, "Q").End(xlUp).Row

        For i = 2 To lastRowPlt
            Dim cellR As Variant, cellS As Variant
            cellR = wsPlt.Cells(i, "R").Value
            cellS = wsPlt.Cells(i, "S").Value
            
            If Not IsError(cellR) And IsNumeric(cellR) And Trim(cellR & "") <> "" Then
                widthPlt = CDbl(cellR)
            Else
                widthPlt = 0
            End If
            If Not IsError(cellS) And IsNumeric(cellS) And Trim(cellS & "") <> "" Then
                lengthPlt = CDbl(cellS)
            Else
                lengthPlt = 0
            End If
            
            sizePlt = widthPlt & "x" & lengthPlt
            catPlt = wsPlt.Cells(i, "V").Value
            If IsNumeric(wsPlt.Cells(i, "T").Value) Then
                qtyPlt = wsPlt.Cells(i, "T").Value
            Else
                qtyPlt = 0
            End If
            If Not IsEmpty(catPlt) And catPlt <> "" And qtyPlt > 0 Then
                If Not matDict.Exists(sizePlt) Then
                    Set matDict(sizePlt) = CreateObject("Scripting.Dictionary")
                End If
                If Not matDict(sizePlt).Exists(catPlt) Then
                    Set matDict(sizePlt)(catPlt) = CreateObject("Scripting.Dictionary")
                    matDict(sizePlt)(catPlt)("qty") = 0
                    matDict(sizePlt)(catPlt)("vol") = 0 ' объём не считаем, просто для совместимости структуры
                End If
                matDict(sizePlt)(catPlt)("qty") = matDict(sizePlt)(catPlt)("qty") + qtyPlt
            End If
        Next i


    ' 4. Создаём новый лист
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets("Расчётный лист").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set wsDest = wb.Sheets.Add(After:=wsSrc): wsDest.Name = "Расчётный лист"
    wsDest.Activate
    wsDest.Range("B3").Select
    ActiveWindow.FreezePanes = True
    
    ' 5. Заголовки таблицы (3 колонки на категорию)
    wsDest.Cells(1, 1).Value = "Материал"
        destCol = 2
        ' -- Блок "Итог" --
        wsDest.Cells(1, destCol).Value = "Итог"
        wsDest.Cells(2, destCol).Value = "шт."
        wsDest.Cells(2, destCol + 1).Value = "V, м3"
        wsDest.Cells(2, destCol + 2).Value = "M, кг"
        wsDest.Range(wsDest.Cells(1, destCol), wsDest.Cells(1, destCol + 2)).Merge
        destCol = destCol + 3

        ' -- Категории, как обычно --
        For Each cat In catList
            wsDest.Cells(1, destCol).Value = cat
            wsDest.Cells(2, destCol).Value = "шт."
            wsDest.Cells(2, destCol + 1).Value = "V, м3"
            wsDest.Cells(2, destCol + 2).Value = "M, кг"
            wsDest.Range(wsDest.Cells(1, destCol), wsDest.Cells(1, destCol + 2)).Merge
            destCol = destCol + 3
        Next cat


    ' 6. Заполняем строки
    destRow = 3
        ' Массив для хранения итогов по каждой строке
    Dim totalQty As Double, totalVol As Double, totalMass As Double

    For Each key In matDict.Keys
    dims = Split(key, "x")
    Dim widthStr As String, heightStr As String, lengthStr As String
    Dim isBoard As Boolean, isSheet As Boolean
    If UBound(dims) = 2 Then
        widthStr = dims(0)
        heightStr = dims(1)
        lengthStr = dims(2)
        isBoard = True
        isSheet = False
    ElseIf UBound(dims) = 1 Then
        widthStr = dims(0)
        lengthStr = dims(1)
        isBoard = False
        isSheet = True
    Else
        GoTo NextKey
    End If

    wsDest.Cells(destRow, 1).Value = key
    destCol = 5
    For Each cat In catList
        If isBoard Then
            ' Доски
            wsDest.Cells(destRow, destCol).Formula = _
                "=SUMIFS('Раскрой Древесины'!U2:U58," & _
                "'Раскрой Древесины'!R2:R58,""" & widthStr & """," & _
                "'Раскрой Древесины'!S2:S58,""" & heightStr & """," & _
                "'Раскрой Древесины'!T2:T58,""" & lengthStr & """," & _
                "'Раскрой Древесины'!X2:X58,""" & cat & """)"
            wsDest.Cells(destRow, destCol + 1).Formula = _
                "=SUMIFS('Раскрой Древесины'!V2:V58," & _
                "'Раскрой Древесины'!R2:R58,""" & widthStr & """," & _
                "'Раскрой Древесины'!S2:S58,""" & heightStr & """," & _
                "'Раскрой Древесины'!T2:T58,""" & lengthStr & """," & _
                "'Раскрой Древесины'!X2:X58,""" & cat & """)"
        ElseIf isSheet Then
            ' Плиты
            wsDest.Cells(destRow, destCol).Formula = _
                "=SUMIFS('Раскрой Плит'!T2:T58," & _
                "'Раскрой Плит'!R2:R58,""" & widthStr & """," & _
                "'Раскрой Плит'!S2:S58,""" & lengthStr & """," & _
                "'Раскрой Плит'!V2:V58,""" & cat & """)"
            wsDest.Cells(destRow, destCol + 1).Formula = "" ' Объем не считается для плит
        End If
        ' Масса для всех одинаково
        wsDest.Cells(destRow, destCol + 2).Formula = _
            "=" & wsDest.Cells(destRow, destCol + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & _
            "*" & Replace(Format(massPerM3, "0.00"), ",", ".")
        destCol = destCol + 3
    Next cat

    ' Итоговые формулы (оставляй как раньше)
    qtyCols = ""
    volCols = ""
    For k = 0 To catList.Count - 1
        If k > 0 Then
            qtyCols = qtyCols & ","
            volCols = volCols & ","
        End If
        qtyCols = qtyCols & wsDest.Cells(destRow, 5 + 3 * k).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        volCols = volCols & wsDest.Cells(destRow, 5 + 3 * k + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Next k

    wsDest.Cells(destRow, 2).Formula = "=SUM(" & qtyCols & ")"
    wsDest.Cells(destRow, 3).Formula = "=SUM(" & volCols & ")"
    wsDest.Cells(destRow, 4).FormulaR1C1 = "=RC[-1]*" & Replace(Format(massPerM3, "0.00"), ",", ".")
    destRow = destRow + 1
NextKey:
Next key


    ' 7. Форматирование
    wsDest.Columns("A:A").ColumnWidth = 15
    destCol = 2
    Do While wsDest.Cells(1, destCol).Value <> ""
        wsDest.Columns(destCol).ColumnWidth = 7
        wsDest.Columns(destCol + 1).ColumnWidth = 7
        wsDest.Columns(destCol + 2).ColumnWidth = 7
        destCol = destCol + 3
    Loop

    ' Объединение A1:A2
    wsDest.Range("A1:A2").Merge
    wsDest.Range("A1:A2").HorizontalAlignment = xlCenter
    wsDest.Range("A1:A2").VerticalAlignment = xlCenter

    ' Перенос текста в шапке
    wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(2, destCol - 1)).WrapText = True

    ' Высота первой строки
    wsDest.Rows(1).RowHeight = 45
    wsDest.Rows(2).RowHeight = 20

    wsDest.Rows("1:2").Font.Bold = True
    wsDest.Rows("1:2").HorizontalAlignment = xlCenter
    wsDest.Rows("1:2").VerticalAlignment = xlCenter

    ' Обводка всей таблицы (шапка + данные)
    Dim lastRow As Long, lastCol As Long
    lastRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row
    If lastRow < 3 Then lastRow = 3
    lastCol = wsDest.Cells(1, wsDest.Columns.Count).End(xlToLeft).Column
    
    ' --- Жирная обводка блока каждой категории ---
    destCol = 1
    For i = 1 To catList.Count + 2
        Set rngBlock = wsDest.Range(wsDest.Cells(1, destCol), wsDest.Cells(lastRow, destCol + 2))
        With rngBlock.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
        End With
        With rngBlock.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThick
        End With
        With rngBlock.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThick
        End With
        With rngBlock.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThick
        End With
        With rngBlock.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
        End With
        If destCol = 1 then 
        destCol = destCol + 1
        Else destCol = destCol + 3
        EndIf
    Next i


        ' --- Заливка данных через один блок категории ---
    Dim fillCol As Long
    Dim fillGray As Boolean
    destCol = 2
    fillGray = False ' Можно начать с заливки или без — меняй на True если надо первый блок цветом

    For i = 1 To catList.Count + 1
        If fillGray Then
            Set rngBlock = wsDest.Range(wsDest.Cells(3, destCol), wsDest.Cells(lastRow, destCol + 2))
            rngBlock.Interior.Color = RGB(237, 245, 240)
        End If
        fillGray = Not fillGray
        destCol = destCol + 3
    Next i


        ' --- Светло-серая заливка блока "Материал" и "Итог" (A, B, C, D) для данных ---
    With wsDest.Range(wsDest.Cells(3, 1), wsDest.Cells(lastRow, 4))
        .Interior.Color = RGB(240, 240, 240)
    End With

    ' Заливка шапки
     With wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(2, lastCol+2))
        .Interior.Color = RGB(40, 105, 67)
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub
