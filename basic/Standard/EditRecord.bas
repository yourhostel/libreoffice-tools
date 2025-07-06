REM  *****  BASIC  *****

' EditRecord.bas

Sub StartEdit()
    ResetPeopleTodayFilter(False)
    If Not IsInAllowedColumns() Then Exit Sub
    If Not IsInAllowedRows() Then Exit Sub
    MoveCursorToColumnA()
    ShowForm()
End Sub

' =====================================================
' === Функція IsInAllowedColumns ======================
' =====================================================
' → Перевіряє, чи курсор знаходиться в межах стовпців A–S.
' → Повертає True, якщо так, інакше False та показує повідомлення.
Function IsInAllowedColumns() As Boolean
    Dim oDoc, oSel, col
    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection
    col = oSel.RangeAddress.StartColumn

    If col < 0 Or col > 18 Then
        ShowDialog "Помилка", "Курсор має бути в межах стовпців A–S."
        IsInAllowedColumns = False
    Else
        IsInAllowedColumns = True
    End If
End Function

' =====================================================
' === Функція IsInAllowedRows =========================
' =====================================================
' → Перевіряє, чи курсор знаходиться в дозволеному діапазоні рядків.
' → Викликає FindEditRange для отримання меж діапазону.
' → Повертає True, якщо так, інакше False та показує повідомлення.
Function IsInAllowedRows() As Boolean
    Dim StartRow As Long, EndRow As Long
    Call FindEditRange(StartRow, EndRow)
    ' MsgBox "StartRow: " & StartRow & " EndRow: " & EndRow
    If StartRow = -1 Or EndRow = -1 Then
        ShowDialog "Помилка", "Не вдалося визначити діапазон редагування."
        IsInAllowedRows = False
        Exit Function
    End If

    Dim oDoc, oSel, row
    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection
    row = oSel.RangeAddress.StartRow + 1

    ' якщо курсор поза межами дозволеного діапазону
    If row < StartRow Or row > EndRow Then
        ShowDialog "Помилка", "Курсор поза дозволеним діапазоном рядків: " & (StartRow) & "–" & (EndRow) & "."
        IsInAllowedRows = False
    Else
        IsInAllowedRows = True
    End If
End Function

' =====================================================
' === Процедура FindEditRange =========================
' =====================================================
' → Визначає діапазон дозволених рядків для редагування.
' → StartRow — перший після останньої «інкасації» +2 (або 4‑й, якщо не знайдено).
' → EndRow — перший порожній у стовпці E після StartRow, з компенсацією +1.
Sub FindEditRange(ByRef StartRow As Long, ByRef EndRow As Long)
    Dim oDoc, oSheet, oRange, oDesc, oFoundAll
    Dim i As Long

    oDoc = ThisComponent
    oSheet = oDoc.Sheets(0)

    StartRow = -1
    EndRow = -1

    ' діапазон стовпця E від рядка 4 до кінця
    oRange = oSheet.getCellRangeByPosition(4, 3, 4, oSheet.Rows.Count - 1)

    ' шукаємо останню "інкасація" у діапазоні
    oDesc = oRange.createSearchDescriptor()
    oDesc.SearchString = "інкасація"
    oDesc.SearchRegularExpression = False
    oDesc.SearchCaseSensitive = False

    oFoundAll = oRange.findAll(oDesc)

    If Not IsNull(oFoundAll) And oFoundAll.Count > 0 Then
        ' якщо знайдено
        ' — беремо останню + 2 компенсуємо ініціалізацію та рядок інкасації
        StartRow = oFoundAll.getByIndex(oFoundAll.Count - 1).RangeAddress.StartRow + 2
    Else
        ' якщо не знайдено — нова таблиця, стартуємо з 4‑го рядка
        StartRow = 3
    End If

    ' MsgBox "Початок діапазону (StartRow): " & StartRow

    ' шукаємо перший порожній рядок у стовпці E від StartRow вниз
    For i = StartRow To oSheet.Rows.Count - 1
        If Trim(oSheet.getCellByPosition(4, i).String) = "" Then
            EndRow = i - 1
            Exit For
        End If
    Next i

    ' якщо не знайшли порожній — беремо останній рядок аркуша
    If EndRow = -1 Then EndRow = oSheet.Rows.Count - 1

    EndRow = EndRow + 1

    ' MsgBox "Кінець діапазону (EndRow): " & EndRow
End Sub

' =====================================================
' === Процедура MoveCursorToColumnA ===================
' =====================================================
' → Переносить курсор у стовпець A поточного рядка.
Sub MoveCursorToColumnA()
    Dim oDoc, oCtrl, oSel, row
    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection
    row = oSel.RangeAddress.StartRow

    oCtrl = oDoc.CurrentController
    oCtrl.select(oDoc.Sheets(0).getCellByPosition(0, row))
End Sub