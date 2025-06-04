Sub CopyWoodQuantity()
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim lastSrc As Long, lastDest As Long
    Dim i As Long, j As Long
    Dim srcQ As Variant, srcR As Variant, srcS As Variant, srcU As Variant
    Dim destA As Variant, destB As Variant

    Set wsSrc = ThisWorkbook.Sheets("Раскрой Древесины")
    Set wsDest = ThisWorkbook.Sheets("Вспомогательная (Панели)")

    ' Найти последний заполненный ряд
    lastSrc = wsSrc.Cells(wsSrc.Rows.Count, "Q").End(xlUp).Row
    lastDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row

    ' Если только одна строка - делаем массив
    If lastSrc - 1 < 1 Then
        srcQ = Array(wsSrc.Range("Q2").Value)
        srcR = Array(wsSrc.Range("R2").Value)
        srcS = Array(wsSrc.Range("S2").Value)
        srcU = Array(wsSrc.Range("U2").Value)
    Else
        srcQ = wsSrc.Range("Q2:Q" & lastSrc).Value
        srcR = wsSrc.Range("R2:R" & lastSrc).Value
        srcS = wsSrc.Range("S2:S" & lastSrc).Value
        srcU = wsSrc.Range("U2:U" & lastSrc).Value
    End If

    If lastDest - 1 < 1 Then
        destA = Array(wsDest.Range("A2").Value)
        destB = Array(wsDest.Range("B2").Value)
        lastDest = 2
    Else
        destA = wsDest.Range("A2:A" & lastDest).Value
        destB = wsDest.Range("B2:B" & lastDest).Value
    End If

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' Формируем словарь: ключ = (Q & "|" & R & "x" & S), значение = U
    For i = 1 To UBound(srcQ, 1)
        If Not IsEmpty(srcQ(i, 1)) And Not IsEmpty(srcR(i, 1)) And Not IsEmpty(srcS(i, 1)) Then
            dict(srcQ(i, 1) & "|" & srcR(i, 1) & "x" & srcS(i, 1)) = srcU(i, 1)
        End If
    Next i

    Dim destCount As Long
    If IsArray(destA) Then
        destCount = UBound(destA, 1)
    Else
        destCount = 1
    End If

    ' Сопоставляем и переносим значения в G
    For j = 1 To destCount
        Dim key As String
        If IsArray(destA) Then
            key = destA(j, 1) & "|" & destB(j, 1)
        Else
            key = destA & "|" & destB
        End If
        If dict.Exists(key) Then
            wsDest.Cells(j + 1, "D").Value = dict(key)
        Else
            wsDest.Cells(j + 1, "D").Value = "" ' если нет совпадения, очищаем
        End If
    Next j

    MsgBox "Перенос значений завершен."
End Sub
