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
    Set oRange = GetPeopleRange()

    ' ==== Створення дескриптора фільтру ====
    Set oFilterDesc = oRange.createFilterDescriptor(True)

    ' ==== Оголошення фільтруючих полів (3 умови) ====
    Dim oFilterFields(2) As New com.sun.star.sheet.TableFilterField

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

    ' ==== Застосовуємо фільтр ====
    oFilterDesc.FilterFields = oFilterFields()
    oRange.filter(oFilterDesc)

    ' ==== Підрахунок видимих рядків і тих, у кого CheckOut сьогодні ====
    Dim result As Variant
    result = CountVisibleRows(oRange)

    ' ==== Виведення результатів ====
	MsgBox "Порахуйте. Повинно бути " & result(0) & " " & PersonWord(result(0)) & "." & Chr(10) & _
       		Chr(10) & result(1) & " " & PersonWord(result(1)) & " до оплати, або на виселення." _
       		, 32, "Людей зараз:  " & result(0)
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
    Dim dToday As Double

    ' ==== Отримання поточного листа ====
    oSheet = ThisComponent.CurrentController.ActiveSheet

    ' ==== Отримання поточної дати без часу ====
    dToday = Int(Now())

    ' ==== Ініціалізація лічильників ====
    iVisibleCount = 0
    iTermExpired = 0

    ' ==== Межі діапазону ====
    iStartRow = oRange.RangeAddress.StartRow
    iLastRow = oRange.RangeAddress.EndRow

    ' ==== Проходимо по кожному рядку діапазону ====
    For iRow = iStartRow To iLastRow
        ' ==== Отримання комірки дати виселення (колонка E) ====

        Set oCellOut = oSheet.getCellByPosition(4, iRow) ' Колонка E

		' ==== Перевірка: рядок не відфільтрований і має дату виселення !== 0 ====
		If Not oSheet.getRows().getByIndex(iRow).IsFiltered And oCellOut.getValue() <> 0 Then
    		iVisibleCount = iVisibleCount + 1
		End If

        ' ==== Перевірка: чи дата виселення сьогодні ====
        If Int(oCellOut.getValue()) = dToday Then
    		iTermExpired = iTermExpired + 1
		End If
    Next iRow

    ' ==== Повертаємо масив з двома значеннями: [видимих, на виселення] ====
    CountVisibleRows = Array(iVisibleCount, iTermExpired)
End Function

' =====================================================
' === Процедура ResetPeopleTodayFilter ================
' =====================================================
' → Скидає всі фільтри на діапазоні людей.
' → Видаляє умови фільтрації та ставить курсор на першу пусту клітинку в колонці A.
Sub ResetPeopleTodayFilter(SetCursor As Boolean)
    Dim oRange As Object, oFilterDesc As Object

    ' ==== Отримання діапазону ====
    Set oRange = GetPeopleRange()

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
Sub ResetFilter()
    ResetPeopleTodayFilter(True)
End Sub

' =====================================================
' === Функція GetPeopleRange ==========================
' =====================================================
' → Повертає діапазон даних з A4 до останнього заповненого рядка по колонці A.
' → Діапазон завжди від A4:E[останній рядок].
Function GetPeopleRange() As Object
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

    ' ==== Визначення діапазону A4:E(останній рядок) ====
    Set oRange = oSheet.getCellRangeByPosition(0, 3, 4, iLastRow - 1)

    ' ==== Повертаємо знайдений діапазон ====
    Set GetPeopleRange = oRange
End Function
