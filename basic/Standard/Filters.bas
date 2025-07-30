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
    Dim oRange      As Object
    Dim oFilterDesc As Object
    Dim dToday      As Double
    dToday = Int(Now())
    
    ' ==== Отримання діапазону даних ====
    Set oRange = GetRecordsRange()
    
    ' MsgBox oRange.RangeAddress.StartRow & " → " & oRange.RangeAddress.EndRow

    ' ==== Створення дескриптора фільтру ====
    oFilterDesc = oRange.createFilterDescriptor(True)

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
    	.Field = 18 ' Колонка S код
    	.Operator = com.sun.star.sheet.FilterOperator.NOT_EQUAL ' Не дорівнює
    	.IsNumeric = True
    	.NumericValue = 7
	End With
	
	' ==== Четверта умова: S код пусто ====
    'With oFilterFields(3)
        '.Field = 18 ' Колонка S код
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
           String(70, "-")

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
            sMsg = sMsg & String(70, "-") & Chr(10)
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
    Dim n1 As Long
    
    n = Abs(n) Mod 100
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
    Dim iRow                   As Long
    Dim iStartRow              As Long
    Dim iLastRow               As Long
    Dim iVisibleCount          As Long
    Dim iTermExpired           As Long
    Dim oSheet                 As Object
    Dim oCellOut               As Object
    Dim oCellSurname           As Object
    Dim oNameAndPatronymic     As Object
    Dim dToday                 As Double
    Dim aPeopleForEviction()   As String
    Dim aPeopleFromBlackList() As String
    
    ' ==== Отримання поточного листа ====
    oSheet = ThisComponent.CurrentController.ActiveSheet
    
    ' ==== Отримання поточної дати без часу ====  
    dToday = Int(Now())

    ' ==== Ініціалізація лічильників ====
    iVisibleCount = 0
    iTermExpired  = 0
    iBlacklisted  = 0

    ' ==== Межі діапазону ==== 
    iStartRow = oRange.RangeAddress.StartRow
    iLastRow = oRange.RangeAddress.EndRow

    ' ==== Проходимо по кожному рядку діапазону ====
    For iRow = iStartRow To iLastRow
        ' ==== Перевірка: рядок видимий ====		
        If Not oSheet.getRows().getByIndex(iRow).IsFiltered Then
            iVisibleCount = iVisibleCount + 1
            
            ' ==== Чи виселення сьогодні ====		
            Set oCellOut = oSheet.getCellByPosition(4, iRow)                    ' CheckOut (E) 
            If Int(oCellOut.getValue()) = dToday Then
                oCellSurname = oSheet.getCellByPosition(1, iRow)                ' (B) 
                oNameAndPatronymic = oSheet.getCellByPosition(2, iRow)          ' (C)
                ReDim Preserve aPeopleForEviction(iTermExpired)    
                aPeopleForEviction(iTermExpired) = oCellSurname.getString() & " " & oNameAndPatronymic.getString()
                iTermExpired = iTermExpired + 1
            End If
        
            ' ==== Чи у чорному списку ====           
            If oSheet.getCellByPosition(18, iRow).getValue = 28 Then
                oCellSurname = oSheet.getCellByPosition(1, iRow)                ' (B)
                oNameAndPatronymic = oSheet.getCellByPosition(2, iRow)          ' (C)
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
' === Процедура ResetFilter ============================
' =====================================================
' → Скидає фільтр на основному діапазоні записів (GetRecordsRange).
' → Створює пустий дескриптор фільтра й застосовує його до діапазону.
'
' → Якщо параметр SetCursor = True — додатково викликає SelectFirstEmptyInA
'     для переходу до першої порожньої клітинки в колонці A.
Sub ResetFilter(SetCursor As Boolean)
    Dim oRange      As Object
    Dim oFilterDesc As Object

    ' ==== Отримання діапазону ====   
    oRange = GetRecordsRange()

    ' ==== Створення дескриптора фільтру ====  
    oFilterDesc = oRange.createFilterDescriptor(True)
    
    ' ==== Скидання фільтру ==== 
    oFilterDesc.FilterFields = Array()

    oRange.filter(oFilterDesc)
    
    if SetCursor Then
        SelectFirstEmptyInA()
    End If
End Sub

' =====================================================
' === Процедура ResetFilterAdmin ======================
' =====================================================
' → Скидає фільтр з підтвердженням через діалог авторизації ShowNegetDialog.
' → Якщо авторизацію скасовано — виводить повідомлення та припиняє виконання.
' → Інакше викликає ResetFilter із передачею курсора.
Sub ResetFilterAdmin(True)
    If Not ShowNegetDialog(NEGET_RULES) Then
        MsgDlg "Помилка", String(18, " ") & "Операцію скасовано.", False, 50, 130
        Exit Sub
    End If
    
    ResetFilter(SetCursor)  
End Sub

' =====================================================
' === Процедура ResetFilterlimited ====================
' =====================================================
' → Скидає фільтр і обмежує видимість таблиці до останніх VISIBLE_ROWS рядків.
' → Використовується для обмежених прав доступу (без діалогу підтвердження).
Sub ResetFilterlimited()
    ResetFilter(True)
    LimitScroll()
End Sub

' =====================================================
' === Процедура LimitScroll ===========================
' =====================================================
' → Приховує усі рядки таблиці "Data", крім останніх VISIBLE_ROWS.
' → Якщо загальна кількість рядків менша за VISIBLE_ROWS — нічого не робить.
Sub LimitScroll()
    Dim oDoc     As Object : odoc     = ThisComponent
    Dim oSheet   As Object : oSheet   = oDoc.Sheets.getByName("data")
    Dim nLastRow As Long   : nLastRow = FindLastRow
    Dim oRange   As Object

    If nLastRow < VISIBLE_ROWS Then Exit Sub
    
    oRange = oSheet.getCellRangeByPosition(0, 3, 0, nLastRow - VISIBLE_ROWS)
    oRange.Rows.IsVisible = False
End Sub

' =====================================================
' === Функція FindLastRow =============================
' =====================================================
' → Визначає останній використаний рядок у таблиці "Data".
' → Повертає номер останнього непорожнього рядка (Long).
Function FindLastRow() As Long
    Dim oSheet  As Object : oSheet = ThisComponent.Sheets.getByName("Data")
    Dim oCursor As Object : oCursor = oSheet.createCursor()
    oCursor.gotoEndOfUsedArea(True)
    Dim lastRow As Long : lastRow = oCursor.RangeAddress.EndRow

    FindLastRow = lastRow
End Function

' =====================================================
' === Функція GetPeopleRange ==========================
' =====================================================
' → Повертає діапазон даних з A4 до останнього заповненого рядка по колонці A.
' → Діапазон завжди від A4:E[останній рядок].
Function GetRecordsRange As Object
    Dim oDoc     As Object
    Dim oSheet   As Object
    Dim oRange   As Object
    Dim iLastRow As Long

    ' ==== Отримання документа і активного аркуша ====   
    oDoc   = ThisComponent
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
        oRange = oSheet.getCellRangeByPosition(0, 3, 20, 3)
    Else
        ' Повертаємо діапазон від A4 до останнього заповненого
        oRange = oSheet.getCellRangeByPosition(0, 3, 20, iLastRow - 1)
    End If
    
    ' ==== Повертаємо знайдений діапазон ====
    GetRecordsRange = oRange
End Function

' =====================================================
' === Процедура DebugRangeValues ======================
' =====================================================
' → Виводить значення CheckIn, CheckOut, D та IsFiltered для кожного рядка.
' → Використовується для налагодження діапазону записів.
Sub DebugRangeValues()
    Dim oRange    As Object    ' діапазон записів
    Dim oSheet    As Object    ' аркуш
    Dim iRow      As Long      ' індекс рядка
    Dim iStartRow As Long      ' початковий рядок діапазону
    Dim iEndRow   As Long      ' кінцевий рядок діапазону
    Dim sOut      As String    ' рядок результату
    Dim dToday    As Double    ' сьогоднішня дата (без часу)

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

' =====================================================
' === Процедура FilterEncashmentAudit =================
' =====================================================
' → Виконує фільтрацію записів для аудиту інкасації (код 7 у колонці S).
'
' → Алгоритм:
'     1. Знаходить останній запис інкасації (код "7").
'     2. Якщо інкасацію не знайдено — показує всі записи.
'     3. Якщо знайдена — фільтрує записи за умовами:
'         • Created (O) < дати інкасації
'         • CheckIn (A) > дати інкасації
'     4. Додатково відображає інкасацію та всі записи після неї.
'
' → Для отримання діапазону використовується зовнішня функція GetRecordsRange.
Sub FilterEncashmentAudit()
    Dim oDoc      As Object : oDoc = ThisComponent
    Dim oSheet    As Object : oSheet = oDoc.Sheets.getByName("Data")
    Dim oRange    As Object : oRange = GetRecordsRange() ' Твоя функція отримання діапазону
    Dim nLastRow  As Long   : nLastRow = oSheet.getRows().getCount()
    
    ' === Пошук останньої інкасації (код 7 у колонці S — тобто індекс 18) ===
    Dim EncashRow As Long : EncashRow = -1
    Dim sCode     As String
    Dim i As Long
    
    For i = 3 To nLastRow - 1
        sCode = oSheet.getCellByPosition(18, i).String
        If sCode = "" Then Exit For
        If sCode = "7" Then EncashRow = i
    Next i

    If EncashRow = -1 Then
        ' Якщо інкасації немає — показуємо всі записи
        For i = 3 To nLastRow - 1
            If oSheet.getCellByPosition(18, i).String = "" Then Exit For
            oSheet.Rows.getByIndex(i).IsVisible = True
        Next i
        MsgDlg "Увага", "Інкасацію не знайдено. Виведено всі записи.", False, 50 
        Exit Sub
    End If

    ' === Отримуємо дату створення інкасації (колонка O — індекс 14) ===
    Dim EncashDate As Double
    EncashDate = oSheet.getCellByPosition(14, EncashRow).Value

    ' === Створюємо дескриптор фільтра ===
    Dim oFilterDesc As Object
    oFilterDesc = oRange.createFilterDescriptor(True)

    ' === Накладаємо 2 умови ===
    ' — дата створення < дати інкасації (Created < EncashDate)
    ' — дата вселення > дати інкасації (CheckIn > EncashDate)
    Dim oFilterFields(1) As New com.sun.star.sheet.TableFilterField

    With oFilterFields(0)
        .Field = 14 ' колонка O — Created
        .Operator = com.sun.star.sheet.FilterOperator.LESS
        .IsNumeric = True
        .NumericValue = EncashDate
    End With

    With oFilterFields(1)
        .Field = 0 ' колонка A — CheckIn
        .Operator = com.sun.star.sheet.FilterOperator.GREATER
        .IsNumeric = True
        .NumericValue = EncashDate
    End With

    oFilterDesc.FilterFields = oFilterFields()
    oRange.filter(oFilterDesc)

    ' === Відображаємо інкасацію та все що після ===
    
    For i = EncashRow To nLastRow - 1
        If oSheet.getCellByPosition(18, i).String = "" Then Exit For
        oSheet.Rows.getByIndex(i).IsVisible = True
    Next i
End Sub

