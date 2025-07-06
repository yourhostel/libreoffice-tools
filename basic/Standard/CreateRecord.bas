REM  *****  BASIC  *****

' CreateRecord.bas

Dim FormResult As Boolean

Sub StartCreate()
    SelectFirstEmptyInA()
    ShowForm()
End Sub

' =====================================================
' === Функція ShowForm ================================
' =====================================================
' → Запускає діалог введення нового запису.
' → Відображає форму, підключає слухачі, перевіряє введені дані та вставляє їх у таблицю.
' → Повертає рядок: "OK" — якщо дані збережені, "Cancel" — якщо відмінено.
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

' =====================================================
' === Процедура InsertButton_actionPerformed ==========
' =====================================================
' → Обробник події натискання кнопки "Вставити".
' → Перевіряє всі поля форми та вставляє дані у таблицю.
' → Якщо всі перевірки пройдені, закриває діалог.
Sub InsertButton_actionPerformed(oEvent As Object)
    Dim oDoc As Object, oSel As Object, oDialog As Object
    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection
    oDialog = oEvent.Source.getContext()

        If OffsetReasonValidation(oSel, oDialog) _
       And CheckOccupiedPlace(oDialog) _
       And FinanceAreNumbersValidation(oDialog, "ExpenseField;IncomeField") _
       And FinanceCommentValidation(oDialog) _
       And PersonDataValidation(oDialog) _
       And PhoneValidation(oDialog) _
       And BirthDateValidation(oDialog) _
       And PassportDataValidation(oDialog) Then

    	   OffsetReasonInsertion(oSel, oDialog)	   ' Q, P      причина зсуву	зсув
    	   DateRangeInsertion(oSel, oDialog)	   ' A, E, O   заселення, виселення, створено
    	   CodeInsertion(oSel, oDialog)            ' D         код
    	   PersonDataInsertion(oSel, oDialog)      ' B, C      прізвище, ім'я по батькові
    	   PaidInsertion(oSel, oDialog)            ' F         сплачено
    	   FinanceInsertion(oSel, oDialog)         ' G, H, I   видаток, прихід, коментар
    	   PhoneInsertion(oSel, oDialog)		   ' J         телефон
    	   PassportBirthInsertion(oSel, oDialog)   ' K, L      паспортні дані, дата народження
    	   HostelInsertion(oSel , oDialog)         ' N         хостел
    	   PlaceInsertion(oSel, oDialog)           ' R         місце

           FormResult = True   ' Ставимо True тільки якщо валідація пройшла та вставка відпрацювала коректно
           oDialog.endExecute()
    End If
End Sub

' =====================================================
' === Процедура InsertButton_disposing ===============
' =====================================================
' → Викликається при видаленні слухача з кнопки InsertButton.
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub InsertButton_disposing(oEvent As Object)
End Sub

' =====================================================
' === Функція ValidateSelection =======================
' =====================================================
' → Перевіряє правильність вибраної комірки для вставки нового запису.
' → Повертає True — якщо вибір коректний, False — якщо ні.
Function ValidateSelection() As Boolean
    Dim oDoc As Object, oSel As Object
    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection

    ' ==== Перевірка чи це саме комірка ====
    'If Not oSel.supportsService("com.sun.star.sheet.SheetCell") Then
        'ShowDialog "Помилка", "Виділи комірку у стовпці A."
        'ValidateSelection = False
        'Exit Function
    'End If

    ' ==== Отримуємо адресу ====
    Dim oCursorAddress As Object
    oCursorAddress = oSel.getCellAddress()

    ' ==== Заборона на перші три рядки ====
    'If oCursorAddress.Row < 3 Then
        'ShowDialog "Помилка", "Заборонено використовувати перші три рядки."
        'ValidateSelection = False
        'Exit Function
    'End If

    ' ==== Перевірка: вибрана клітинка у стовпці A ====
    'If oCursorAddress.Column <> 0 Then
        'ShowDialog "Помилка", "Виберіть клітинку у стовпці A."
        'ValidateSelection = False
        'Exit Function
    'End If

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

' =====================================================
' === Функція OffsetReasonValidation ==================
' =====================================================
' → Перевіряє, що якщо Offset ≠ 0, то заповнено поле Reason.
' → Повертає True — якщо умова виконана, False — якщо ні.
Function OffsetReasonValidation(oSel As Object, oDialog As Object) As Boolean
    ' ==== Читаємо значення полів ====
    Dim sOffset As String, sReason As String
    sOffset = oDialog.getControl("OffsetField").getText()
    sReason = oDialog.getControl("ReasonField").getText()

    ' ==== Перевірка ====
    If Val(sOffset) <> 0 And Trim(sReason) = "" Then
        ShowDialog "Увага!", "Поле 'Причина зсуву' не може бути порожнім при ненульовому зсуві!"
        OffsetReasonValidation = False
        Exit Function
    End If

    ' ==== Успішно ====
    OffsetReasonValidation = True
End Function

' =====================================================
' === Процедура OffsetReasonInsertion ================
' =====================================================
' → Вставляє значення Offset і Reason у таблицю (стовпці Q та P).
Sub OffsetReasonInsertion(oSel As Object, oDialog As Object)
    Dim oSheet As Object
    oSheet = oSel.Spreadsheet

    Dim sOffset As String, sReason As String
    sOffset = oDialog.getControl("OffsetField").getText()
    sReason = oDialog.getControl("ReasonField").getText()

    oSheet.getCellByPosition(16, oSel.CellAddress.Row).String = sOffset ' Q
    oSheet.getCellByPosition(15, oSel.CellAddress.Row).String = sReason ' P

End Sub

' =====================================================
' === Процедура DateRangeInsertion ====================
' =====================================================
' → Вставляє дати заселення, виселення та створення у таблицю.
' → Форматує дати у вигляді "DD.MM.YYYY" та "DD.MM.YYYY HH:MM".
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

    ' ==== Формат дати без часу ====
    nFormatDate = oFormats.queryKey("DD.MM.YYYY", oLocale, True)
    If nFormatDate = -1 Then
        nFormatDate = oFormats.addNew("DD.MM.YYYY", oLocale)
    End If

    ' ==== Формат дати з часом ====
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

' =====================================================
' === Функція PersonDataValidation ====================
' =====================================================
' → Перевіряє заповнення полів Прізвище, Ім'я та По батькові.
' → Повертає True — якщо всі заповнені, False — якщо ні.
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
        ShowDialog "Увага!", "Необхідно заповнити всі поля: Прізвище, Ім'я, По батькові."
        PersonDataValidation = False
        Exit Function
    End If

    PersonDataValidation = True
End Function

' =====================================================
' === Процедура PersonDataInsertion ===================
' =====================================================
' → Вставляє прізвище та ім'я по батькові у таблицю (стовпці B та C).
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

    ' ==== Колонка B — прізвище ====
    oSheet.getCellByPosition(1, oSel.CellAddress.Row).setString(sLastName)

    ' ==== Колонка C — ім'я та по батькові ====
    oSheet.getCellByPosition(2, oSel.CellAddress.Row).setString(sFullName)
End Sub

' =====================================================
' === Процедура PaidInsertion =========================
' =====================================================
' → Вставляє значення з поля PaidField у таблицю (стовпець F).
Sub PaidInsertion(oSel As Object, oDialog As Object)
    ' ==== Отримання листа ====
    Dim oSheet As Object
    oSheet = oSel.Spreadsheet

    ' ==== Отримання значення з поля форми ====
    Dim dPrice As Double
    dPrice = Val(oDialog.getControl("PaidField").getText())

    ' ==== Вставка у колонку 'сплачено'  F (індекс 5) ====
    oSheet.getCellByPosition(5, oSel.CellAddress.Row).String = dPrice
End Sub

' =====================================================
' === Функція FinanceAreNumbersValidation =============
' =====================================================
' → Перевіряє, що зазначені фінансові поля містять числові значення.
' → Повертає True — якщо все коректно, False — якщо є помилки.
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

    		ShowDialog "Увага!", "Поле """ & MapGet(Map, aFields(i)) & """ повинно містити число."
    		bIsValid = False
		End If

    Next i

    FinanceAreNumbersValidation = bIsValid
End Function

' =====================================================
' === Функція FinanceCommentValidation ================
' =====================================================
' → Перевіряє, що якщо Expense або Income ≠ 0, то заповнений Comment.
' → Повертає True — якщо умова виконана, False — якщо ні.
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

        ShowDialog "Увага!", "Поле """ & MapGet(Map, fieldName) & """ заповнено, напишіть коментар."
        FinanceCommentValidation = False
        Exit Function
    End If

    FinanceCommentValidation = True
End Function

' =====================================================
' === Процедура FinanceInsertion ======================
' =====================================================
' → Вставляє фінансові дані та коментар у таблицю (стовпці G, H, I).
Sub FinanceInsertion(oSel As Object, oDialog As Object)
    ' ==== Отримання листа ====
    Dim oSheet As Object
    oSheet = oSel.Spreadsheet

    ' ==== Отримання значення з поля форми Expense, Income ====
    Dim dExpense As Double
	Dim dIncome As Double
    Dim dComment As String
    dExpense = Val(oDialog.getControl("ExpenseField").getText())
	dIncome = Val(oDialog.getControl("IncomeField").getText())
	dComment = oDialog.getControl("CommentField").getText()

    ' ==== Вставка у колонку 'видаток'  G (індекс 6) ====
    oSheet.getCellByPosition(6, oSel.CellAddress.Row).Value = dExpense
    ' ==== Вставка у колонку 'прихід'  H (індекс 7) ====
    oSheet.getCellByPosition(7, oSel.CellAddress.Row).Value = dIncome
    ' ==== Вставка у колонку 'коментар'  I (індекс 8) ====
    oSheet.getCellByPosition(8, oSel.CellAddress.Row).String = dComment
End Sub

' =====================================================
' === Функція IsPhoneMinimalValid =====================
' =====================================================
' → Перевіряє мінімальну коректність номера телефону.
' → Повертає True — якщо виглядає коректним, False — якщо ні.
' → Базова перевірка на формат без помилок типу "123", "asd", "0".
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

' =====================================================
' === Функція PhoneValidation =========================
' =====================================================
' → Перевіряє правильність заповнення поля Телефон.
' → Повертає True — якщо коректно, False — якщо ні.
Function PhoneValidation(oDialog As Object) As Boolean
	Dim sPhone As String
    sPhone = oDialog.getControl("PhoneField").getText()

	If Not IsPhoneMinimalValid(sPhone) Then
		ShowDialog "Поле 'Телефон' заповнене некоректно." _
    	,"Приклад: +380671234567, 0671234567, +12025550198, 441234567890"
		PhoneValidation = False
		Exit Function
    End If
    PhoneValidation = True
End Function

' =====================================================
' === Процедура PhoneInsertion ========================
' =====================================================
' → Вставляє номер телефону у таблицю (стовпець J).
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

' =====================================================
' === Функція PassportDataValidation ==================
' =====================================================
' → Перевіряє коректність паспорта: кількість полів і мінімальні довжини.
' → Використовує DateFormatValidation для дати.
' → Якщо не валідно — викликає ShowPassportInvalid().
Function PassportDataValidation(oDialog As Object) As Boolean
    Dim parts() As String
    parts = Split(oDialog.getControl("PassportField").getText(), ",")

    If UBound(parts) <> 4 Then
        ShowPassportInvalid()
        PassportDataValidation = False
        Exit Function
    End If

    If UBound(parts) <> 4 _
       Or Len(Trim(parts(0))) < 7 _
       Or Len(Trim(parts(1))) < 10 _
       Or Len(Trim(parts(2))) < 7 _
       Or Not DateFormatValidation(Trim(parts(3))) _
       Or Len(Trim(parts(4))) < 15 Then

        ShowPassportInvalid()
        PassportDataValidation = False
        Exit Function
    End If

    PassportDataValidation = True
End Function

' =====================================================
' === Процедура ShowPassportInvalid ===================
' =====================================================
' → Виводить повідомлення про неправильні паспортні дані.
Sub ShowPassportInvalid()
    ShowDialog "Паспортні дані потребують уточнення.", _
               "Коректна форма подання:", _
               "Номер(≥7), Ким видан(≥10), Де(≥7), Коли(dd.mm.yyyy), Прописка(≥15)"
End Sub

' =====================================================
' === Функція BirthDateValidation =====================
' =====================================================
' → Перевіряє правильність формату дати народження.
' → Якщо не валідно — показує діалог із підказкою.
Function BirthDateValidation(oDialog As Object) As Boolean
    If Not DateFormatValidation(oDialog.getControl("BirthDateField").getText()) Then
        ShowDialog "Формат дати народження не валідний", _
                   "Подання має бути в такому вигляді: dd.mm.yyyy"
        BirthDateValidation = False
        Exit Function
    End If
    BirthDateValidation = True
End Function

' =====================================================
' === Функція DateFormatValidation ====================
' =====================================================
' → Перевіряє, що рядок перетворюється в дату.
' → Використовує CDate та обробку помилки.
' → Повертає True — якщо дата валідна, False — якщо ні.
Function DateFormatValidation(sDate As String) As Boolean
    If Trim(sDate) = "" Then
        DateFormatValidation = False
        Exit Function
    End If
    On Error GoTo ErrHandler
    Dim d As Date
    d = CDate(sDate)
    DateFormatValidation = True
    Exit Function
ErrHandler:
    DateFormatValidation = False
End Function

' =====================================================
' === Процедура PassportBirthInsertion ===============
' =====================================================
' → Вставляє паспортні дані та дату народження в таблицю (стовпці K, L).
Sub PassportBirthInsertion(oSel As Object, oDialog As Object)
    Dim oSheet As Object
    oSheet = oSel.Spreadsheet

    Dim sPassport As String
    Dim sBirthDate As String

    sPassport = oDialog.getControl("PassportField").getText()
    sBirthDate = oDialog.getControl("BirthDateField").getText()

    oSheet.getCellByPosition(10, oSel.CellAddress.Row).String = sPassport ' K
    oSheet.getCellByPosition(11, oSel.CellAddress.Row).String = sBirthDate ' L
End Sub

' =====================================================
' === Процедура HostelInsertion =======================
' =====================================================
' → Вставляє значення хостелу у таблицю (стовпець N).
Sub HostelInsertion(oSel As Object, oDialog As Object)
    Dim oSheet As Object
    oSheet = oSel.Spreadsheet

    oSheet.getCellByPosition(13, oSel.CellAddress.Row).String = HOSTEL ' N
End Sub

' =====================================================
' === Процедура CodeInsertion =========================
' =====================================================
' → Вставляє код у таблицю (стовпець D).
Sub CodeInsertion(oSel As Object, oDialog As Object)
    Dim oSheet As Object
    oSheet = oSel.Spreadsheet
    Dim sCode As String

    sCode = oDialog.getControl("CodeCombo").getText()

    oSheet.getCellByPosition(3, oSel.CellAddress.Row).String = sCode ' D
End Sub

' =====================================================
' === Процедура PlaceInsertion ========================
' =====================================================
' → Вставляє номер місця в таблицю (стовпець R) як число.
Sub PlaceInsertion(oSel As Object, oDialog As Object)
    Dim oSheet As Object
    oSheet = oSel.Spreadsheet
    Dim sPlace As String

    sPlace = oDialog.getControl("PlaceCombo").getText()

    oSheet.getCellByPosition(17, oSel.CellAddress.Row).Value = Val(sPlace) ' R
End Sub

' =====================================================
' === Функція CreateDialog ============================
' =====================================================
' → Створює та налаштовує діалогову форму введення нового запису.
' → Додає всі поля, мітки, кнопки, слухачі та повертає готовий діалог.
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

    ' ==== Кількість днів ====
    Call ComboBoxTemplate(oDialogModel, "Duration",     "Кількість днів:", 65, 45, "1", 50, 50, VALID_DURATIONS)
    Call AddDurationComboListeners(oDialog)

    ' ==== Причини зсуву ====
    Call FieldTemplate(oDialogModel,      "Reason",      "Причина зсуву:", 120, 45, "", 50, 165)

    ' ==== Поля визначення ====
    Call ComboBoxTemplate(oDialogModel,     "Code",                "Код:", 10, 75, "1", 50, 50, LIST_OF_CODES)
    Call AddCodeComboListeners(oDialog)
    Call ComboBoxTemplate(oDialogModel,     "Place",             "Місце:", 65, 75, "1", 50, 50, LIST_OF_PLACES)
    Call AddPlaceComboListeners(oDialog)

    ' ==== Фінансові поля ====
    Call FieldTemplate(oDialogModel, 	    "Paid",           "Сплачено:", 120, 75, PriceWithSheet, 50, 50, True)
    Call FieldTemplate(oDialogModel, 	 "Expense",            "Видаток:", 235, 75, "0", 50, 50)
    Call FieldTemplate(oDialogModel, 	  "Income",             "Прихід:", 290, 75, "0", 50, 50)

    ' ==== Інші дані ====
    Call FieldTemplate(oDialogModel,     "Comment",           "Коментар:", 10, 105, "", 70, 330)

    ' ==== Персональні дані ====
    Call FieldTemplate(oDialogModel,    "LastName",           "Прізвище:", 10,  135, "", 70, 100)
    Call FieldTemplate(oDialogModel,   "FirstName",               "Ім'я:", 125,  135, "", 70, 100)
    Call FieldTemplate(oDialogModel,  "Patronymic",        "По батькові:", 240,  135, "_", 70, 100)

    ' ==== Контактна інформація ====
    Call FieldTemplate(oDialogModel,       "Phone",            "Телефон:", 10, 165, "", 70, 100)
    Call FieldTemplate(oDialogModel,   "BirthDate",    "Дата народження:", 125, 165, "", 100, 100)

    ' ==== Документ ====
    Call FieldTemplate(oDialogModel,    "Passport",     "Паспортні дані:", 10, 195, "", 100, 330)

    Call CalculatePaidFieldWithPlace(oDialog)
    Call UpdatePlaceCombo(oDialog)

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
    Call AddButton(oDialogModel, "InsertButton", "Вставити", 150, 225, 60, 14)

    oDialog.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)

    Call CalculatePaidFieldWithPlace(oDialog)

    ' ==== Повертаємо ====
    CreateDialog = oDialog
End Function

