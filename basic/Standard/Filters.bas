REM  *****  BASIC  *****

' Filters.bas

' =====================================================
' === Процедура PeopleTodayFilter =====================
' =====================================================
' → Фільтрує людей, які зараз проживають, по трьох умовах:
'   — CheckIn <= сьогодні +1
'   — CheckOut >= сьогодні
'   — Код ≠ 7
' → Виводить повідомлення з підрахованими результатами.
Sub PeopleTodayFilter()
    Dim oRange As Object, oFilterDesc As Object
    Dim dToday As Double
    dToday = Int(Now())

    ' ==== Отримання діапазону даних ====
    Set oRange = GetRecordsRange()

    ' MsgBox oRange.RangeAddress.StartRow & " → " & oRange.RangeAddress.EndRow

    ' ==== Створення дескриптора фільтру ====
    Set oFilterDesc = oRange.createFilterDescriptor(True)

    ' ==== Оголошення фільтруючих полів (3 умови) ====
    Dim oFilterFields(3) As New com.sun.star.sheet.TableFilterField

    ' ==== Перша умова: CheckIn <= сьогодні (включно) ====
    With oFilterFields(0)
        .Field = 0 ' Колонка A
        .Operator = com.sun.star.sheet.FilterOperator.LESS_EQUAL ' Менше або дорівнює. (сьогодні + 1)
        .IsNumeric = True
        .NumericValue = dToday ' Щоб включити всіх хто заселився сьогодні
    End With

    ' ==== Друга умова: CheckOut >= сьогодні ====
    With oFilterFields(1)
        .Field = 4 ' Колонка E
        .Operator = com.sun.star.sheet.FilterOperator.GREATER_EQUAL ' Більше або дорівнює. Не пізніше ніж сьогодні
        .IsNumeric = True
        .NumericValue = dToday ' Тобто хто ще не виїхав
    End With

    ' ==== Третя умова: D <> 7 (фільтруємо особливий статус) ====
	With oFilterFields(2)
    	.Field = 3 ' Колонка D
    	.Operator = com.sun.star.sheet.FilterOperator.NOT_EQUAL ' Не дорівнює
    	.IsNumeric = True
    	.NumericValue = 7
	End With

	' ==== Четверта умова: M пусто ====
    'With oFilterFields(3)
        '.Field = 3 ' Колонка M
        '.Operator = com.sun.star.sheet.FilterOperator.NOT_EQUAL
    	'.IsNumeric = True
    	'.NumericValue = 28
    'End With

    ' ==== Застосовуємо фільтр ====
    oFilterDesc.FilterFields = oFilterFields()
    oRange.filter(oFilterDesc)

    ' ==== Підрахунок видимих рядків і тих, у кого CheckOut сьогодні ====
    Dim result As Variant
    result = CountVisibleRows(oRange)

    ' ==== Виведення результатів ====
    ShowCountResults result
End Sub

' =====================================================
' === Sub ShowCountResults ============================
' =====================================================
' → Виводить підсумкове повідомлення про кількість видимих рядків,
'     тих, що виселяються сьогодні, та в чорному списку.
' → Виводить імена людей з result(3) та result(4) кожного з нового рядка.
' → Приймає масив Variant з результатами функції CountVisibleRows.
' → Використовує MsgDlg для показу повідомлення.
Sub ShowCountResults(result As Variant)
    Dim sMsg As String

    sMsg = "Порахуйте. Повинно бути " & result(0) & " " & PersonWord(result(0)) & "." & Chr(10) & Chr(10) & _
           result(1) & " " & PersonWord(result(1)) & " до оплати, або на виселення:" & Chr(10) & _
           "____________________________________________________"

    ' ==== result(3) — на виселення сьогодні ====
    If IsArray(result(3)) Then
        sMsg = sMsg & Chr(10)
        For i = LBound(result(3)) To UBound(result(3))
            sMsg = sMsg & "    " & result(3)(i) & Chr(10)
        Next i
    End If

    ' ==== result(4) — чорний список ====
    If result(2) <> 0 Then
        sMsg = sMsg & Chr(10) & result(2) & " " & PersonWord(result(2)) & " виключено (чорний список):" & Chr(10)

        If IsArray(result(4)) Then
            sMsg = sMsg & "____________________________________________________" & Chr(10)
            For i = LBound(result(4)) To UBound(result(4))
                sMsg = sMsg & "    " & result(4)(i) & Chr(10)
            Next i
        End If
    End If

    MsgDlg "Людей зараз:  " & result(0), sMsg, True, 95
End Sub

' =====================================================
' === Функція PersonWord ==============================
' =====================================================
' → Повертає слово «особа», «особи» або «осіб» у правильній формі.
' → В залежності від числа n.
Function PersonWord(n As Long) As String
    n = Abs(n) Mod 100
    Dim n1 As Long
    n1 = n Mod 10

    If (n > 10 And n < 20) Then
        PersonWord = "осіб"
    ElseIf (n1 > 1 And n1 < 5) Then
        PersonWord = "особи"
    ElseIf (n1 = 1) Then
        PersonWord = "особа"
    Else
        PersonWord = "осіб"
    End If
End Function

' =====================================================
' === Функція CountVisibleRows ========================
' =====================================================
' → Рахує кількість видимих (не відфільтрованих) рядків у діапазоні.
' → Додатково рахує скільки з них з CheckOut на сьогодні.
' → Повертає масив: [видимих_рядків, на_виселення_сьогодні].
Function CountVisibleRows(oRange As Object) As Variant
    Dim oSheet As Object
    Dim iRow As Long, iStartRow As Long, iLastRow As Long
    Dim iVisibleCount As Long, iTermExpired As Long
    Dim oCellOut As Object
    Dim oCellSurname As Object
    Dim oNameAndPatronymic As Object
    Dim dToday As Double
    Dim aPeopleForEviction() As String
    Dim aPeopleFromBlackList() As String

    ' ==== Отримання поточного листа ====
    oSheet = ThisComponent.CurrentController.ActiveSheet

    ' ==== Отримання поточної дати без часу ====
    dToday = Int(Now())

    ' ==== Ініціалізація лічильників ====
    iVisibleCount = 0
    iTermExpired = 0
    iBlacklisted = 0

    ' ==== Межі діапазону ====
    iStartRow = oRange.RangeAddress.StartRow
    iLastRow = oRange.RangeAddress.EndRow

    ' ==== Проходимо по кожному рядку діапазону ====
    For iRow = iStartRow To iLastRow
        ' ==== Перевірка: рядок видимий ====
        If Not oSheet.getRows().getByIndex(iRow).IsFiltered Then
            iVisibleCount = iVisibleCount + 1

            ' ==== Чи виселення сьогодні ====
            Set oCellOut = oSheet.getCellByPosition(4, iRow)                        ' CheckOut (E)
            If Int(oCellOut.getValue()) = dToday Then
                Set oCellSurname = oSheet.getCellByPosition(1, iRow)                ' (B)
                Set oNameAndPatronymic = oSheet.getCellByPosition(2, iRow)          ' (C)
                ReDim Preserve aPeopleForEviction(iTermExpired)
                aPeopleForEviction(iTermExpired) = oCellSurname.getString() & " " & oNameAndPatronymic.getString()
                iTermExpired = iTermExpired + 1
            End If

            ' ==== Чи у чорному списку ====
            If oSheet.getCellByPosition(3, iRow).getValue = 28 Then
                Set oCellSurname = oSheet.getCellByPosition(1, iRow)                 ' (B)
                Set oNameAndPatronymic = oSheet.getCellByPosition(2, iRow)           ' (C)
                ReDim Preserve aPeopleFromBlackList(iBlacklisted)
                aPeopleFromBlackList(iBlacklisted) = oCellSurname.getString() & " " & oNameAndPatronymic.getString()
                iBlacklisted = iBlacklisted + 1
            End If
        End If
    Next iRow

    ' ==== Повертаємо масив з двома значеннями: [видимих, на виселення] ====
    CountVisibleRows = Array(iVisibleCount, iTermExpired, iBlacklisted, aPeopleForEviction, aPeopleFromBlackList)
End Function

' =====================================================
' === Процедура ResetPeopleTodayFilter ================
' =====================================================
' → Скидає всі фільтри на діапазоні людей.
' → Видаляє умови фільтрації та ставить курсор на першу пусту клітинку в колонці A.
Sub ResetFilter(SetCursor As Boolean)
    Dim oRange As Object, oFilterDesc As Object

    ' ==== Отримання діапазону ====
    Set oRange = GetRecordsRange()

    ' ==== Створення дескриптора фільтру ====
    Set oFilterDesc = oRange.createFilterDescriptor(True)

    ' ==== Скидання фільтру ====
    oFilterDesc.FilterFields = Array()

    oRange.filter(oFilterDesc)

    if SetCursor Then
        SelectFirstEmptyInA()
    End If
End Sub

' =====================================================
' === Процедура ResetFilter ===========================
' =====================================================
' → Скидає всі фільтри на діапазоні людей.
' → Видаляє умови фільтрації та ставить курсор на першу пусту клітинку в колонці A.
' → Використовується для прив'язування до кнопки з параметром позиціонування курсору.
Sub ResetPeopleTodayFilter()
    ResetFilter(True)
End Sub

' =====================================================
' === Функція GetPeopleRange ==========================
' =====================================================
' → Повертає діапазон даних з A4 до останнього заповненого рядка по колонці A.
' → Діапазон завжди від A4:E[останній рядок].
Function GetRecordsRange As Object
    Dim oDoc As Object, oSheet As Object
    Dim oRange As Object
    Dim iLastRow As Long

    ' ==== Отримання документа і активного аркуша ====
    oDoc = ThisComponent
    oSheet = oDoc.CurrentController.ActiveSheet

    ' ==== Пошук останнього рядка за колонкою A ====
    iLastRow = 3 ' Починаємо з рядка 4 (індексація з 0)

    ' ==== Пошук до першої порожньої комірки у колонці A ====
    Do While oSheet.getCellByPosition(0, iLastRow).getValue() <> 0 ' Перевіряємо колонку A
        iLastRow = iLastRow + 1
    Loop

    ' ==== Якщо немає жодного заповненого рядка ====
    If iLastRow = 3 Then
        ' Повертаємо діапазон хоча б A4:R4
        Set oRange = oSheet.getCellRangeByPosition(0, 3, 20, 3)
    Else
        ' Повертаємо діапазон від A4 до останнього заповненого
        Set oRange = oSheet.getCellRangeByPosition(0, 3, 20, iLastRow - 1)
    End If

    ' ==== Повертаємо знайдений діапазон ====
    Set GetRecordsRange = oRange
End Function

' =====================================================
' === Процедура DebugRangeValues ======================
' =====================================================
' → Виводить значення CheckIn, CheckOut, D та IsFiltered для кожного рядка.
' → Використовується для налагодження діапазону записів.
Sub DebugRangeValues()
    Dim oRange As Object       ' діапазон записів
    Dim oSheet As Object       ' аркуш
    Dim iRow As Long           ' індекс рядка
    Dim iStartRow As Long      ' початковий рядок діапазону
    Dim iEndRow As Long        ' кінцевий рядок діапазону
    Dim sOut As String         ' рядок результату
    Dim dToday As Double       ' сьогоднішня дата (без часу)

    ' ==== Отримуємо сьогоднішню дату без часу ====
    dToday = Int(Now())

    ' ==== Отримуємо діапазон і аркуш ====
    Set oRange = GetRecordsRange()
    Set oSheet = ThisComponent.CurrentController.ActiveSheet

    ' ==== Межі діапазону ====
    iStartRow = oRange.RangeAddress.StartRow
    iEndRow = oRange.RangeAddress.EndRow

    ' ==== Проходимо по кожному рядку та формуємо рядок виводу ====
    For iRow = iStartRow To iEndRow
       ' формуємо строку виводу:
       ' рядок у людських індексах (з 1)
       ' колонка A
       ' колонка E
       ' колонка D
       ' чи відфільтровано
        sOut = sOut & "Row " & (iRow+1) & ": " & _
        "CheckIn=" & oSheet.getCellByPosition(0, iRow).getValue() & ", " & _
        "CheckOut=" & oSheet.getCellByPosition(4, iRow).getValue() & ", " & _
        "D=" & oSheet.getCellByPosition(3, iRow).getValue() & ", " & _
        "IsFiltered=" & oSheet.getRows().getByIndex(iRow).IsFiltered & Chr(10)
    Next iRow

    ' ==== Показуємо результат ====
    MsgDlg "Debug Values", sOut, True
End Sub
