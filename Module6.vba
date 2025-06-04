Sub WriteReferenceFormulas()
    Dim wsFull As Worksheet, wsTest As Worksheet
    Set wsFull = ThisWorkbook.Sheets("Панели (все)")
    Set wsTest = ThisWorkbook.Sheets("Тест Панели")
    
    Dim lastFull As Long, lastTest As Long
    lastFull = wsFull.Cells(wsFull.Rows.Count, "D").End(xlUp).Row
    lastTest = wsTest.Cells(wsTest.Rows.Count, "C").End(xlUp).Row
    
    Dim countFull As Object
    Set countFull = CreateObject("Scripting.Dictionary")
    
    ' Индексация урезанной таблицы
    Dim posDict As Object
    Set posDict = CreateObject("Scripting.Dictionary")
    Dim keyTest As String, j As Long
    For j = 2 To lastTest
        keyTest = wsTest.Cells(j, "C").Value
        If Not posDict.Exists(keyTest) Then
            Set posDict(keyTest) = New Collection
        End If
        posDict(keyTest).Add j
    Next j
    
    ' Для каждой строки полной таблицы формируем ссылку или оставляем пусто
    Dim keyFull As String, i As Long
    For i = 2 To lastFull
        keyFull = wsFull.Cells(i, "D").Value
        If keyFull <> "" And posDict.Exists(keyFull) Then
            If countFull.Exists(keyFull) Then
                countFull(keyFull) = countFull(keyFull) + 1
            Else
                countFull(keyFull) = 1
            End If
            Dim idx As Long
            idx = countFull(keyFull)
            If posDict(keyFull).Count >= idx Then
                Dim testRow As Long
                testRow = posDict(keyFull)(idx)
                wsFull.Cells(i, "E").Formula = "='Тест Панели'!D" & testRow
            Else
                wsFull.Cells(i, "E").Formula = ""
            End If
        Else
            wsFull.Cells(i, "E").Formula = ""
        End If
    Next i
End Sub
