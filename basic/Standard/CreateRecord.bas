REM  *****  BASIC  *****

' CreateRecord.bas

Dim FormResult As Boolean

' ==== Запускає діалог введення нового запису ====
' Відповідає за створення, відображення та обробку діалогової форми.
'
' Алгоритм роботи:
' 1. Перевіряє правильність вибраної комірки за допомогою ValidateSelection().
'    — Якщо перевірка не пройдена, функція завершується і повертає "Cancel".
' 2. Створює діалог за допомогою функції CreateDialog().
' 3. Підключає обробник події натискання кнопки "Вставити" (InsertButton) через CreateUnoListener().
' 4. Підключає слухача для поля OffsetField — динамічне відображення/приховування поля Reason.
' 5. Відображає діалог методом oDialog.execute().
'    — Якщо користувач натискає кнопку "Вставити" і всі валідації успішні (FormResult = True),
'      то виводиться повідомлення "Дані збережено".
'    — Якщо користувач натискає "Закрити" або натискає "Вставити", але валідація не пройдена (FormResult = False),
'      то виводиться повідомлення "Вихід без змін".
' 6. Після завершення діалогу:
'    — Знімає слухач з кнопки InsertButton.
'    — Закриває (dispose) діалог.
' 7. Повертає результат виконання:
'    — "OK" — якщо діалог був підтверджений.
'    — "Cancel" — якщо діалог був скасований або не пройшов перевірку.
'
' Повертає:
' — String: "OK" або "Cancel"
Function ShowForm() As String


    FormResult = False

    ' === Перевірка чи є активна комірка ===
    If Not ValidateSelection() Then
        ShowForm = "Cancel"
        Exit Function
    End If

    Dim oDialog As Object
    Set oDialog = CreateDialog()

    ' === Кнопка Вставити ===
    Dim oButtonInsert As Object
    oButtonInsert = oDialog.getControl("InsertButton")

    ' === Обробник кнопки ===
    Dim oListenerInsert As Object
    oListenerInsert = CreateUnoListener("InsertButton_", "com.sun.star.awt.XActionListener")
    oButtonInsert.addActionListener(oListenerInsert)

	' === Підключення слухача до OffsetField ===
	Call AddTextFieldsOffsetListener(oDialog)

    ' === Змінна результату ===
    Dim sResult As String
    sResult = "Cancel"

    ' === Запуск діалогу ===
    If oDialog.execute() = 1 Then
        ' Натиснута кнопка (будь-яка) — тут можна перевіряти логіку
        sResult = "OK"
    End If

    If FormResult Then
    	MsgBox "" & Chr(10) & "Дані збережено", 64, "Готово"
	Else
    	MsgBox "" & Chr(10) & "Скасовано", 48, "Вихід без змін"
	End If

    ' === Очищення ===
    oButtonInsert.removeActionListener(oListenerInsert)
    oDialog.dispose()

    ShowForm = sResult
End Function

' ==== Обробник події натискання кнопки "Вставити" ====
' Основна процедура, яка виконується при натисканні кнопки InsertButton у формі.
'
' Алгоритм дій:
' 1. Отримує активний документ і поточну вибрану комірку (визначає рядок вставки).
' 2. Запускає послідовно перевірки правильності введених даних:
'    — OffsetReasonValidation — перевірка правильності зсуву і причини.
'    — DateRangeValidation — перевірка правильності діапазону дат (кількість днів).
'    — PersonDataValidation — перевірка заповнення персональних даних (ПІБ).
'    — CodeValidation — перевірка правильності вибору коду.
' 3. Якщо всі перевірки пройдені (allOk = True):
'    — Виконується вставка даних у таблицю:
'       • Offset і Reason → стовпці Q (16) та P (15)
'       • Дати заїзду, виїзду, створення → стовпці A (0), E (4), O (14)
'       • Персональні дані → стовпці B (1) та C (2)
'       • Сума оплати → стовпець F (5)
'    — Змінна FormResult встановлюється у True (флаг успішного виконання).
'    — Діалог закривається командою oDialog.endExecute().
'
' Якщо хоча б одна перевірка не пройдена — дані не записуються, діалог залишається відкритим для виправлення помилок.
'
' Параметри:
' — oEvent — об’єкт події, який містить контекст виклику (діалог, у якому натиснута кнопка).
Sub InsertButton_actionPerformed(oEvent As Object)
    Dim oDoc As Object, oSel As Object
    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection

    Dim oDialog As Object
    oDialog = oEvent.Source.getContext()


        Dim allOk As Boolean
        allOk = True

        If Not OffsetReasonValidation(oSel, oDialog) Then
            allOk = False
        End If

        If Not PersonDataValidation(oDialog) Then
    		allOk = False
		End If

        If Not FinanceAreNumbersValidation(oDialog, "ExpenseField;IncomeField") Then
    		allOk = False
		End If

		If Not FinanceCommentValidation(oDialog) Then
			allOk = False
		End If

		If Not PhoneValidation(oDialog) Then
			allOk = False
		End If

    	If allOk Then

    	    OffsetReasonInsertion(oSel, oDialog)	' Q, P      причина зсуву	зсув
    		DateRangeInsertion(oSel, oDialog)		' A, E, O   заселення, виселення, створено
    		PersonDataInsertion(oSel, oDialog)      ' B, C      прізвище, ім'я по батькові
    		PaidInsertion(oSel, oDialog)            ' F         сплачено
    		FinanceInsertion(oSel, oDialog)         ' G, H, I   видаток, прихід, коментар
    		PhoneInsertion(oSel, oDialog)			' J         телефон

        	FormResult = True   ' Ставимо True тільки якщо валідація пройшла та вставка відпрацювала коректно
        	oDialog.endExecute()
    	End If
End Sub

' ==== Звільнення ресурсів кнопки InsertButton ====
' Порожній метод, обов’язковий для відповідності інтерфейсу com.sun.star.awt.XActionListener.
' Викликається автоматично при знищенні кнопки або діалогу.
' У поточній реалізації не виконує жодних дій, але необхідний для коректної роботи UNO API.
Sub InsertButton_disposing(oEvent As Object)
End Sub

' ==== Перевіряє правильність вибраної комірки ====
' Перевіряє, що поточна вибрана комірка задовольняє всім вимогам для вставки нового запису.
' Умова для коректного вибору:
' — Вибраний об'єкт є саме коміркою (не діапазон, не рядок, не стовпець).
' — Комірка знаходиться у стовпці A (індекс 0).
' — Комірка не розташована у перших трьох рядках (рядки з індексом 0, 1, 2 заборонені).
' — Комірка є порожньою.
' — Комірка у колонці E (індекс 4) на тому ж рядку теж має бути порожньою.
'
' У разі порушення будь-якої умови показує повідомлення про помилку і повертає False.
' Якщо всі умови виконані — повертає True.
'
' Повертає:
' — True — вибір коректний, можна продовжувати.
' — False — вибір некоректний, дії припиняються.
Function ValidateSelection() As Boolean
    Dim oDoc As Object, oSel As Object
    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection

    ' ==== Перевірка чи це саме комірка ====
    If Not oSel.supportsService("com.sun.star.sheet.SheetCell") Then
        ShowDialog("Помилка", "Виділи комірку у стовпці A.")
        ValidateSelection = False
        Exit Function
    End If

    ' ==== Отримуємо адресу ====
    Dim oCursorAddress As Object
    oCursorAddress = oSel.getCellAddress()

    ' ==== Заборона на перші три рядки ====
    If oCursorAddress.Row < 3 Then
        ShowDialog("Помилка", "Заборонено використовувати перші три рядки.")
        ValidateSelection = False
        Exit Function
    End If

    ' ==== Перевірка: вибрана клітинка у стовпці A ====
    If oCursorAddress.Column <> 0 Then
        ShowDialog("Помилка", "Виберіть клітинку у стовпці A.")
        ValidateSelection = False
        Exit Function
    End If

    ' ==== Перевірка: клітинка має бути порожня ====
    'If oSel.getType() <> com.sun.star.table.CellContentType.EMPTY Then
        'ShowDialog("Помилка", "Комірка вже містить дані. Виберіть порожню комірку.")
        'ValidateSelection = False
        'Exit Function
    'End If

    ' ==== Перевірка комірки у колонці E на цьому ж рядку ====
    Dim oSheet As Object
    oSheet = oSel.Spreadsheet

    Dim oCellE As Object
    oCellE = oSheet.getCellByPosition(4, oCursorAddress.Row) ' E = 4

    'If oCellE.getType() <> com.sun.star.table.CellContentType.EMPTY Then
        'ShowDialog("Помилка", "Комірка у стовпці E на цьому ж рядку містить дані.")
        'ValidateSelection = False
        'Exit Function
    'End If

    ' ==== Все ок ====
    ValidateSelection = True
End Function

' ==== Перевірка полів Offset і Reason ====
' Перевіряє, що якщо значення Offset ≠ 0, то поле Reason не є порожнім.
' Повертає:
' - True — якщо умова виконана (або Offset = 0, або Reason заповнене).
' - False — якщо Offset ≠ 0 і Reason порожнє.
' У разі помилки показує діалогове попередження.
Function OffsetReasonValidation(oSel As Object, oDialog As Object) As Boolean
    ' ==== Читаємо значення полів ====
    Dim sOffset As String, sReason As String
    sOffset = oDialog.getControl("OffsetField").getText()
    sReason = oDialog.getControl("ReasonField").getText()

    ' ==== Перевірка ====
    If Val(sOffset) <> 0 And Trim(sReason) = "" Then
        ShowDialog("Увага!", "Поле 'Причина зсуву' не може бути порожнім при ненульовому зсуві!")
        OffsetReasonValidation = False
        Exit Function
    End If

    ' ==== Успішно ====
    OffsetReasonValidation = True
End Function

' ==== Вставка значень Offset і Reason у таблицю ====
' Вставляє значення з полів форми:
' - OffsetField → у колонку Q (індекс 16)
' - ReasonField → у колонку P (індекс 15)
' Працює по активному рядку, вибраному користувачем.
Sub OffsetReasonInsertion(oSel As Object, oDialog As Object)
    Dim oSheet As Object
    oSheet = oSel.Spreadsheet

    ' ==== Вставка Offset і Reason ====
    Dim sOffset As String, sReason As String
    sOffset = oDialog.getControl("OffsetField").getText()
    sReason = oDialog.getControl("ReasonField").getText()

    oSheet.getCellByPosition(16, oSel.CellAddress.Row).String = sOffset ' Q
    oSheet.getCellByPosition(15, oSel.CellAddress.Row).String = sReason ' P

End Sub

' ==== Запис діапазону дат у таблицю ====
' Обчислює дату заселення (з урахуванням зсуву Offset)
' та дату виселення (дата заселення + Duration)
' Записує:
' - дату заселення в колонку A
' - дату виселення в колонку E
' - дату створення запису (з часом) в колонку O
' Якщо Offset ≠ 0 — застосовує стиль комірки "створено" до дати створення
' Дати форматуються відповідно до українського локалю:
' - "DD.MM.YYYY" для дат
' - "DD.MM.YYYY HH:MM" для дати й часу
Sub DateRangeInsertion(oSel As Object, oDialog As Object)
    ' ==== Читаємо значення Offset і Duration ====
    Dim nOffset As Integer, nDuration As Integer
    nOffset = Val(oDialog.getControl("OffsetField").getText())
    nDuration = Val(oDialog.getControl("DurationCombo").getText())

    ' ==== Отримуємо таблицю і адресу ====
    Dim oSheet As Object, oCursorAddress As Object
    oSheet = oSel.Spreadsheet
    oCursorAddress = oSel.CellAddress

    ' ==== Обчислення базової і кінцевої дати ====
    Dim dBaseDate As Double, dEndDate As Double
    dBaseDate = Int(Now) + nOffset
    dEndDate = dBaseDate + nDuration

    ' ==== Формати дат ====
    Dim oFormats As Object, oLocale As New com.sun.star.lang.Locale
    Dim nFormatDate As Long, nFormatDateTime As Long

    oFormats = ThisComponent.getNumberFormats()
    oLocale.Language = "uk"
    oLocale.Country = "UA"

    ' Формат дати без часу
    nFormatDate = oFormats.queryKey("DD.MM.YYYY", oLocale, True)
    If nFormatDate = -1 Then
        nFormatDate = oFormats.addNew("DD.MM.YYYY", oLocale)
    End If

    ' Формат дати з часом
    nFormatDateTime = oFormats.queryKey("DD.MM.YYYY HH:MM", oLocale, True)
    If nFormatDateTime = -1 Then
        nFormatDateTime = oFormats.addNew("DD.MM.YYYY HH:MM", oLocale)
    End If

    ' ==== Вставка дати заселення в колонку A ====
    Dim oCheckInCell As Object
    Set oCheckInCell = oSheet.getCellByPosition(0, oCursorAddress.Row) ' A
    oCheckInCell.setValue(dBaseDate)
    oCheckInCell.NumberFormat = nFormatDate

    ' ==== Вставка дати виселення в колонку E ====
    Dim oCheckOutCell As Object
    Set oCheckOutCell = oSheet.getCellByPosition(4, oCursorAddress.Row) ' E
    oCheckOutCell.setValue(dEndDate)
    oCheckOutCell.NumberFormat = nFormatDate

    ' ==== Вставка дати створення з часом в колонку O ====
    Dim oCreatedCell As Object
    Set oCreatedCell = oSheet.getCellByPosition(14, oCursorAddress.Row) ' O
    oCreatedCell.setValue(Now())
    oCreatedCell.NumberFormat = nFormatDateTime

    ' ==== Якщо є зсув — застосовуємо стиль "створено" ====
    Dim bWasProtected As Boolean
    bWasProtected = oSheet.IsProtected

    ' ==== Якщо захищений — знімаємо захист ====
    If bWasProtected Then
        oSheet.unprotect(NEGET_RULES)
    End If

    If nOffset <> 0 Then
        oCreatedCell.CellStyle = "створено"
    End If

    ' ==== Повертаємо захист назад ====
    If bWasProtected Then
        oSheet.protect(NEGET_RULES)
    End If
End Sub

' ==== Перевіряє заповнення персональних даних ====
' Перевіряє, що всі три поля: Прізвище (LastNameField), Ім'я (FirstNameField)
' та По батькові (PatronymicField) не є порожніми
' Повертає True — якщо всі поля заповнені
' Повертає False — якщо хоча б одне поле порожнє
' У разі помилки виводить повідомлення через ShowDialog
Function PersonDataValidation(oDialog As Object) As Boolean
    Dim sLastName As String
    Dim sFirstName As String
    Dim sPatronymic As String

    ' ==== Отримання значень полів ====
    sLastName = Trim(oDialog.getControl("LastNameField").getText())
    sFirstName = Trim(oDialog.getControl("FirstNameField").getText())
    sPatronymic = Trim(oDialog.getControl("PatronymicField").getText())

    ' ==== Перевірка на порожні значення ====
    If sLastName = "" Or sFirstName = "" Or sPatronymic = "" Then
        ShowDialog("Увага!", "Необхідно заповнити всі поля: Прізвище, Ім'я, По батькові.")
        PersonDataValidation = False
        Exit Function
    End If

    PersonDataValidation = True
End Function

' ==== Вставляє персональні дані у таблицю ====
' Отримує значення прізвища, ім'я та по батькові з діалогової форми
' Форматує перші літери великими
' Прізвище записується у колонку B (індекс 1)
' Ім'я та по батькові об'єднуються у форматі "Ім'я По батькові" та записуються у колонку C (індекс 2)
' Дані вставляються у рядок, відповідний до поточної вибраної клітинки (oSel)
Sub PersonDataInsertion(oSel As Object, oDialog As Object)
    Dim sLastName As String
    Dim sFirstName As String
    Dim sPatronymic As String
    Dim sFullName As String

    ' ==== Отримання значень ====
    sLastName = Capitalize(Trim(oDialog.getControl("LastNameField").getText()))
    sFirstName = Capitalize(Trim(oDialog.getControl("FirstNameField").getText()))
    sPatronymic = Capitalize(Trim(oDialog.getControl("PatronymicField").getText()))

    ' ==== Формування повного імені ====
    sFullName = sFirstName & " " & sPatronymic

    ' ==== Вставка у таблицю ====
    Dim oSheet As Object
    oSheet = oSel.Spreadsheet

    ' Колонка B — прізвище
    oSheet.getCellByPosition(1, oSel.CellAddress.Row).setString(sLastName)

    ' Колонка C — ім'я та по батькові
    oSheet.getCellByPosition(2, oSel.CellAddress.Row).setString(sFullName)
End Sub

' ==== Розраховує суму оплати за кількістю днів ====
' Зчитує значення полів CodeCombo та DurationField
' Отримує відповідний лист з цінами за кодом (аркуш price1, price2 ... price11)
' Шукає відповідну ціну за значенням Duration у діапазоні A2:B11
' Записує знайдену суму у поле PaidField на формі
' Якщо Code порожній — встановлює PaidField = 0
Sub CalculatePaidFieldByDuration(oEvent)
    ' ==== Отримання діалогу ====
    Dim oDialog As Object
    oDialog = oEvent.Source.getContext()

    ' ==== Отримання Duration ====
    Dim nDuration As Long
    nDuration = Val(oDialog.getControl("DurationCombo").getText())

    ' ==== Отримання значення Code ====
    Dim sCode As String
    sCode = oDialog.getControl("CodeCombo").getText()

    ' ==== Перевірка чи обраний код ====
    If Trim(sCode) = "" Then
        oDialog.getControl("PaidField").setText("0")
        Exit Sub
    End If

    ' ==== Отримання документа і листа з цінами ====
    Dim oDoc As Object, oSheet As Object
    oDoc = ThisComponent
    oSheet = oDoc.Sheets.getByName("price" & sCode)

    ' ==== Пошук ціни у діапазоні A2:B11 ====
    Dim iRow As Long
    Dim dPrice As Double
    dPrice = 0

    For iRow = 1 To 10 ' рядки з 2 по 11 (індексація з 0)
        If oSheet.getCellByPosition(0, iRow).getValue() = nDuration Then
            dPrice = oSheet.getCellByPosition(1, iRow).getValue()
            Exit For
        End If
    Next iRow

    ' ==== Запис у поле Paid ====
    oDialog.getControl("PaidField").setText(CStr(dPrice))
End Sub

' ==== Вставляє значення з поля форми PaidField у таблицю ====
' Отримує значення з поля PaidField
' Вставляє це значення у колонку F (індекс 5) поточного рядка
' oSel — поточна вибрана комірка (визначає рядок)
' oDialog — екземпляр форми, з якої зчитується значення
Sub PaidInsertion(oSel As Object, oDialog As Object)
    ' Отримання листа
    Dim oSheet As Object
    oSheet = oSel.Spreadsheet

    ' Отримання значення з поля форми
    Dim dPrice As Double
    dPrice = Val(oDialog.getControl("PaidField").getText())

    ' Вставка у колонку 'сплачено'  F (індекс 5)
    oSheet.getCellByPosition(5, oSel.CellAddress.Row).String = dPrice
End Sub

' ==== Перевіряє, що поля містять числа ====
' Повертає True — якщо всі поля правильні
' Повертає False — якщо хоча б одне поле некоректне
Function FinanceAreNumbersValidation(oDialog As Object, sFields As String) As Boolean
    Dim aFields() As String
    aFields = Split(sFields, ";")

    Dim i As Integer
    Dim sValue As String
    Dim bIsValid As Boolean
    bIsValid = True

    For i = LBound(aFields) To UBound(aFields)

        sValue = Trim(oDialog.getControl(aFields(i)).getText())

        If Trim(sValue) = "" Or Not IsNumeric(sValue) Then
    		Dim Map As Variant
    		Map = GetFieldToColumnMap()

    		ShowDialog("Увага!", "Поле """ & MapGet(Map, aFields(i)) & """ повинно містити число.")
    		bIsValid = False
		End If

    Next i

    FinanceAreNumbersValidation = bIsValid
End Function

' ==== Перевірка: якщо заповнено Expense або Income, але порожній Comment ====
' Повертає:
' — True — якщо або всі порожні, або коментар є.
' — False — якщо Expense або Income ≠ 0, але Comment пустий.
Function FinanceCommentValidation(oDialog As Object) As Boolean

    Dim dExpense As Double
    Dim dIncome As Double
    Dim sComment As String

    dExpense = Val(oDialog.getControl("ExpenseField").getText())
    dIncome = Val(oDialog.getControl("IncomeField").getText())
    sComment = Trim(oDialog.getControl("CommentField").getText())

    If (dExpense <> 0 Or dIncome <> 0) And sComment = "" Then
        Dim Map As Variant
        Map = GetFieldToColumnMap()

        Dim fieldName As String
        If dExpense <> 0 Then
            fieldName = "ExpenseField"
        Else
            fieldName = "IncomeField"
        End If

        ShowDialog("Увага!", "Поле """ & MapGet(Map, fieldName) & """ заповнено, напишіть коментар.")
        FinanceCommentValidation = False
        Exit Function
    End If

    FinanceCommentValidation = True

End Function

' ==== Вставляє фінансові дані та коментар у таблицю ====
' Отримує значення з полів форми:
' — ExpenseField → витрати
' — IncomeField → надходження
' — CommentField → коментар
'
' Вставляє ці значення у відповідні стовпці таблиці поточного рядка:
' — колонка G (індекс 6) → "видаток"
' — колонка H (індекс 7) → "прихід"
' — колонка I (індекс 8) → "коментар"
'
' Параметри:
' — oSel — поточна вибрана комірка (визначає рядок вставки)
' — oDialog — екземпляр форми, з якої зчитуються значення
'
' Значення витрат і надходжень автоматично перетворюються у числа за допомогою Val().
Sub FinanceInsertion(oSel As Object, oDialog As Object)
    ' Отримання листа
    Dim oSheet As Object
    oSheet = oSel.Spreadsheet

    ' Отримання значення з поля форми Expense, Income
    Dim dExpense As Double
	Dim dIncome As Double
    Dim dComment As String
    dExpense = Val(oDialog.getControl("ExpenseField").getText())
	dIncome = Val(oDialog.getControl("IncomeField").getText())
	dComment = oDialog.getControl("CommentField").getText()

    ' Вставка у колонку 'видаток'  G (індекс 6)
    oSheet.getCellByPosition(6, oSel.CellAddress.Row).Value = dExpense
    ' Вставка у колонку 'прихід'  H (індекс 7)
    oSheet.getCellByPosition(7, oSel.CellAddress.Row).Value = dIncome
    ' Вставка у колонку 'коментар'  I (індекс 8)
    oSheet.getCellByPosition(8, oSel.CellAddress.Row).String = dComment
End Sub

' ==== Перевіряє коректність номера телефону ====
' Функція базової валідації телефонного номера.
' Перевіряє, що значення:
' - складається лише з цифр (після видалення пробілів, тире, дужок);
' - допускається наявність символу "+" на початку;
' - довжина номера від 8 до 15 цифр (включно) після очистки;
' Повертає:
' - True — якщо номер виглядає як коректний;
' - False — якщо номер є занадто коротким, занадто довгим, або містить недопустимі символи.
' Ця валідація не перевіряє код країни та не гарантує, що це мобільний номер —
' лише базова перевірка на формат без помилок типу "123", "asd", "0".
Function IsPhoneMinimalValid(sPhone As String) As Boolean
    ' ==== Очистка від пробілів, тире, дужок ====
    Dim sClean As String
    sClean = Replace(sPhone, " ", "")
    sClean = Replace(sClean, "-", "")
    sClean = Replace(sClean, "(", "")
    sClean = Replace(sClean, ")", "")

    ' ==== Перевірка на плюс на початку ====
    If Left(sClean, 1) = "+" Then
        sClean = Mid(sClean, 2)
    End If

    ' ==== Перевірка: залишилися лише цифри ====
    If Not IsNumeric(sClean) Then
        IsPhoneMinimalValid = False
        Exit Function
    End If

    ' ==== Перевірка довжини номера ====
    ' → Мінімум 8 цифр (розумний мінімум для більшості країн)
    ' → Максимум 15 цифр (згідно стандарту ITU E.164)
    Dim nLen As Integer
    nLen = Len(sClean)

    If nLen >= 8 And nLen <= 15 Then
        IsPhoneMinimalValid = True
    Else
        IsPhoneMinimalValid = False
    End If
End Function

' ==== Перевіряє правильність заповнення поля Телефон ====
' Перевіряє, що значення поля PhoneField відповідає базовим вимогам до номеру телефону:
' - складається з цифр (допускається символ "+" на початку);
' - довжина номера після очистки від пробілів, тире та дужок — від 8 до 15 символів;
' У разі помилки показує діалог із прикладами правильних форматів.
' Повертає:
' - True — якщо значення коректне;
' - False — якщо значення некоректне.
Function PhoneValidation(oDialog As Object) As Boolean
	Dim sPhone As String
    sPhone = oDialog.getControl("PhoneField").getText()

	If Not IsPhoneMinimalValid(sPhone) Then
		ShowDialog("Поле 'Телефон' заповнене некоректно." _
    	,"Приклад: +380671234567, 0671234567, +12025550198, 441234567890")
		PhoneValidation = False
		Exit Function
    End If
    PhoneValidation = True
End Function

' ==== Вставляє значення з поля форми PhoneField у таблицю ====
' Отримує значення з поля PhoneField
' Вставляє це значення у колонку J (індекс 9) поточного рядка
' oSel — поточна вибрана комірка (визначає рядок)
' oDialog — екземпляр форми, з якої зчитується значення
Sub PhoneInsertion(oSel As Object, oDialog As Object)
    ' Отримання листа
    Dim oSheet As Object
    oSheet = oSel.Spreadsheet

    ' Отримання значення з поля форми
    Dim sPhone As String
    sPhone = oDialog.getControl("PhoneField").getText()

    ' Вставка у колонку 'телефон'  J (індекс 9)
    oSheet.getCellByPosition(9, oSel.CellAddress.Row).String = sPhone
End Sub












' ==== Створення діалогу ====

Function CreateDialog() As Object

    Dim oDialog As Object
    Dim oDialogModel As Object
    Dim PriceWithSheet As String
    Dim sDateTime As String
    sDateTime = Format(Now, "DD.MM.YYYY HH:MM:SS")
    PriceWithSheet = ThisComponent.Sheets.getByName("price1").getCellByPosition(1, 1).getValue()
    oDialog = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    oDialogModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    oDialog.setModel(oDialogModel)

    ' ==== Параметри діалогу ====
    With oDialogModel
        .PositionX = 100
        .PositionY = 100
        .Width = 350
        .Height = 250
        .Title = "Новий запис"
    End With

' ==============================================
'   Функція: AddLogo
'   Призначення: Додає логотип на форму (діалог)
'   Параметри:
'     oDialogModel — модель форми
'     iName        — ім'я об'єкта
'     PositionX    — координат X (лівий верхній кут)
'     PositionY    — координат Y (лівий верхній кут)
'     Width        — ширина
'     Height       — висота
' ==============================================

    Call AddLogo (oDialogModel ,"logo", 290, 5, 50, 45)

' ==============================================
'   Функція: FieldTemplate
'   Призначення: Додає на форму мітку та поле вводу
'   Параметри:
'     oDialogModel  — модель форми
'     NamePrefix    — префікс імені для мітки та поля
'     LabelText     — текст мітки
'     PositionX     — координата X (верхній лівий кут мітки та поля)
'     PositionY     — координата Y (для поля, мітка вище)
'     vText         — значення за замовчуванням для поля
'     WidthLabel    — ширина мітки
'     WidthField    — ширина поля
'     ReadOnly      — (необов’язковий) якщо True — поле тільки для читання
' ==============================================

    ' ==== Поточна дата і час ====
    Call FieldTemplate(oDialogModel, "CurrentDate", "Поточні дата і час:", 10, 15, sDateTime, 65, 65, True)

    ' ==== Зсув у днях ====
    Call FieldTemplate(oDialogModel,      "Offset",        "Зсув у днях:", 10, 45, "0", 50, 50)

    ' ==== Причини зсуву ====
    Call FieldTemplate(oDialogModel,      "Reason",      "Причина зсуву:", 80, 45, "", 50, 60)

    ' ==== Кількість днів ====
    'Call FieldTemplate(oDialogModel,    "Duration",     "Кількість днів:", 160, 45, "1", 50, 50)
    Call ComboBoxTemplate(oDialogModel, "Duration", "Кількість днів:", 160, 45, "1", 50, 50, VALID_DURATIONS)
    Call AddDurationComboListeners(oDialog)
     ' ==== Персональні дані ====
    Call FieldTemplate(oDialogModel,    "LastName",           "Прізвище:", 10,  75, "", 70, 100)
    Call FieldTemplate(oDialogModel,   "FirstName",               "Ім'я:", 125,  75, "", 70, 100)
    Call FieldTemplate(oDialogModel,  "Patronymic",        "По батькові:", 240,  75, "_", 70, 100)

    ' ==== Фінансові поля ====
    Call ComboBoxTemplate(oDialogModel,     "Code",                "Код:", 10, 105, "1", 60, 60, LIST_OF_CODES)
    Call AddComboListeners(oDialog)
    Call FieldTemplate(oDialogModel, 	    "Paid",           "Сплачено:", 140, 105, PriceWithSheet, 70, 60)
    Call FieldTemplate(oDialogModel, 	 "Expense",            "Видаток:", 210, 105, "0", 70, 60)
    Call FieldTemplate(oDialogModel, 	  "Income",             "Прихід:", 280, 105, "0", 70, 60)

    ' ==== Інші дані ====
    Call FieldTemplate(oDialogModel,     "Comment",           "Коментар:", 10, 135, "", 70, 330)

    ' ==== Контактна інформація ====
    Call FieldTemplate(oDialogModel,       "Phone",            "Телефон:", 10, 165, "", 70, 100)
    Call FieldTemplate(oDialogModel,   "BirthDate",    "Дата народження:", 120, 165, "", 100, 100)

    ' ==== Документ ====
    Call FieldTemplate(oDialogModel,    "Passport",     "Паспортні дані:", 10, 195, "", 100, 330)

' ==============================================
'   Функція: AddButton
'   Призначення: Додає кнопку на форму (діалог)
'   Параметри:
'     oDialogModel  — модель форми
'     Name          — ім'я об'єкта кнопки
'     Label         — текст на кнопці
'     PositionX     — координата X (лівий верхній кут)
'     PositionY     — координата Y (лівий верхній кут)
'     Width         — ширина кнопки
'     Height        — висота кнопки
' ==============================================

    ' ==== Кнопка вставки ====
    Call AddButton(oDialogModel, "InsertButton", "Вставити", 150, 220, 60, 14)

    oDialog.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)

    ' ==== Повертаємо ====
    CreateDialog = oDialog
End Function