REM  *****  BASIC  *****

' ==== Глобальні змінні для слухача ====

Dim oOffsetControl As Object
Dim oReasonControl As Object
Dim oLabelReasonControl As Object

Sub PickDateAndInsert
    Dim oDoc As Object, oSheet As Object, oCell As Object
    Dim oDialog As Object, oDialogModel As Object
    Dim oDateModel As Object, oButtonModel As Object
    Dim oOffsetModel As Object, oDurationModel As Object
    Dim sOffset As String, sDuration As String
    Dim nOffset As Long, nDuration As Long
    Dim dBaseDate As Date, dEndDate As Date
    Dim oCursorAddress As Object, oTargetCell As Object
    Dim bValid As Boolean, bExit As Boolean

    ' ==== Отримуємо об'єкт форматів чисел для документа ====

	Dim oFormats As Object, oLocale As New com.sun.star.lang.Locale

	oFormats = ThisComponent.getNumberFormats()

	' ==== Встановлюємо локаль для формату дати — українська ====

	oLocale.Language = "uk"
	oLocale.Country = "UA"

	' ==== Отримання поточного документа і активного аркуша ====

    oDoc = ThisComponent
    oSheet = oDoc.CurrentController.ActiveSheet
    oCell = oDoc.CurrentSelection

    ' ==== Адреса вибраної клітинки ====

    oCursorAddress = oCell.getCellAddress()

    ' ==== Заборона на перші три рядки ====

    If oCursorAddress.Row < 3 Then
        MsgBox "Заборонено використовувати перші три рядки.", 16, "Помилка"
        Exit Sub
    End If

    ' ==== Перевірка: вибрана клітинка має бути у стовпці A (індекс 0) ====

    If oCursorAddress.Column <> 0 Then
        MsgBox "Виберіть клітинку у стовпці A.", 48, "Помилка"
        Exit Sub
    End If

    ' ==== Перевірка: клітинка не повинна містити дані ====

    If oCell.getType() <> com.sun.star.table.CellContentType.EMPTY Then
        MsgBox "Комірка вже містить дані. Виберіть порожню комірку.", 16, "Помилка"
        Exit Sub
    End If

    ' ==== Перевірка: клітинка в стовпці E (індекс 4) того ж рядка має бути порожньою ====

    Dim oCellE As Object
    Set oCellE = oSheet.getCellByPosition(4, oCursorAddress.Row)  ' колонка E = 4

    If oCellE.getType() <> com.sun.star.table.CellContentType.EMPTY Then
        MsgBox "Комірка у стовпці E містить дані.", 16, "Помилка"
        Exit Sub
    End If

    ' ==== Створення діалогу ====

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

    Call FieldTemplate(oDialogModel, "CurrentDate", "Поточні дата і час:", 10, 15, Format(Now, "DD.MM.YYYY HH:MM:SS"), 65, 65, True)

    ' ==== Зсув у днях ====

    Call FieldTemplate(oDialogModel, "Offset", "Зсув у днях:", 10, 45, "0", 50, 50)

    ' ==== Кількість днів ====

    Call FieldTemplate(oDialogModel, "Duration", "Кількість днів:", 160, 45, "1", 50, 50)

    ' ==== Причини зсуву ====

    Call FieldTemplate(oDialogModel, "Reason", "Причина зсуву:", 80, 45, "", 50, 60)

     ' ==== Персональні дані ====

    Call FieldTemplate(oDialogModel, "LastName",      "Прізвище:",          10,  75, "", 70, 100)
    Call FieldTemplate(oDialogModel, "FirstName",     "Ім'я:",              130,  75, "", 70, 100)
    Call FieldTemplate(oDialogModel, "Patronymic",    "По батькові:",      240,  75, "", 70, 100)

    ' ==== Фінансові поля ====

    Call FieldTemplate(oDialogModel, "Code",          "Код:",               10, 105, "", 70, 60)
    Call FieldTemplate(oDialogModel, "Paid",          "Сплачено:",          80, 105, "", 70, 60)
    Call FieldTemplate(oDialogModel, "Expense",       "Видаток:",          150, 105, "", 70, 60)
    Call FieldTemplate(oDialogModel, "Income",        "Прихід:",           220, 105, "", 70, 60)

    ' ==== Інші дані ====

    Call FieldTemplate(oDialogModel, "Comment",       "Коментар:",          10, 135, "", 70, 330)

    ' ==== Контактна інформація ====

    Call FieldTemplate(oDialogModel, "Phone",         "Телефон:",           10, 165, "", 70, 100)
    Call FieldTemplate(oDialogModel, "BirthDate",     "Дата народження:",  120, 165, "", 100, 100)

    ' ==== Документ ====

    Call FieldTemplate(oDialogModel, "Passport",      "Паспортні дані:",    10, 195, "", 100, 330)

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

    ' ==== Відображення діалогу ====

    oDialog.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)

    ' ==== Логіка слухача поля OffsetField ====

    Set oOffsetControl = oDialog.getControl("OffsetField")
    Set oReasonControl = oDialog.getControl("ReasonField")
    Set oLabelReasonControl = oDialog.getControl("ReasonLabel")
    oDialog.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)

    ' Підключення слухача до OffsetField
    Dim oOffsetControl As Object
    Set oOffsetControl = oDialog.getControl("OffsetField")

    Dim oReasonControl As Object
    Set oReasonControl = oDialog.getControl("ReasonField")

    Dim oLabelReasonControl As Object
	Set oLabelReasonControl = oDialog.getControl("ReasonLabel")

	' Перше встановлення дефолтного стану при запуску (значення = 0 — ховаємо)
	If Val(oOffsetControl.getText()) = 0 Then
    	oReasonControl.Visible = False
    	oLabelReasonControl.Visible = False
	Else
    	oReasonControl.Visible = True
    	oLabelReasonControl.Visible = True
	End If

	' Створення слухача
	Dim oListener As Object
	Set oListener = CreateUnoListener("OffsetField_", "com.sun.star.awt.XTextListener")

	' Підключаємо слухача
	oOffsetControl.addTextListener(oListener)

    ' ==== Валідація сроків ====

    bValid = False
    Do While Not bValid
        If oDialog.execute() = 0 Then
            oDialog.dispose()
            Exit Sub
        End If

        sOffset = oDialog.getControl("OffsetField").getText()
        sDuration = oDialog.getControl("DurationField").getText()

        If IsNumeric(sOffset) Then nOffset = CLng(sOffset) Else nOffset = 0
        If IsNumeric(sDuration) Then nDuration = CLng(sDuration) Else nDuration = 0

        ' Перевірка допустимих значень тривалості
        If nDuration < 1 Or Not (nDuration <= 7 Or nDuration = 14 Or nDuration = 21 Or nDuration = 28) Then
            MsgBox "Допустимі значення кількості днів: 1–7, 14, 21 або 28", 48, "Неправильне значення"
        Else
            bValid = True
        End If
    Loop

    ' ==== Обчислення базової дати ====

    dBaseDate = Int(Now)
	If nOffset <> 0 Then
    	dBaseDate = dBaseDate + nOffset
	End If

    ' ==== Обчислення кінцевої дати ====

     dEndDate = dBaseDate + nDuration

     ' ==== створюємо форматування ====

	Dim nFormatId As Long
	Dim nFormatDateTime As Long

	nFormatId = oFormats.queryKey("DD.MM.YYYY", oLocale, True)
	If nFormatId = -1 Then
    	nFormatId = oFormats.addNew("DD.MM.YYYY", oLocale)
	End If

	nFormatDateTime = oFormats.queryKey("DD.MM.YYYY HH:MM", oLocale, True)
	If nFormatDateTime = -1 Then
    	nFormatDateTime = oFormats.addNew("DD.MM.YYYY HH:MM", oLocale)
	End If

    ' ==== Вставка заселення ====

    oCell.setValue(dBaseDate)
    oCell.NumberFormat = nFormatId

    ' ==== Вставка виселення ====

    oTargetCell = oSheet.getCellByPosition(4, oCursorAddress.Row)
    oTargetCell.setValue(dEndDate)
    oTargetCell.NumberFormat = nFormatId

    ' ==== Вставка дати створення з часом у колонку O ====

    Dim oCellCreated As Object
    Set oCellCreated = oSheet.getCellByPosition(14, oCursorAddress.Row)
    oCellCreated.setValue(Now())
    oCellCreated.NumberFormat = nFormatDateTime
    If nOffset <> 0 Then oCellCreated.CellStyle = "створено"

    ' ==== Додавання опису причин зсуву до поля "причина"(P) ====

    oSheet.getCellByPosition(15, oCursorAddress.Row).setString(oDialog.getControl("ReasonField").getText())

    ' ==== Звільнення ресурсу ====

    oDialog.dispose()
End Sub

Sub OffsetField_textChanged(oEvent)
    Dim sValue As String
    sValue = oOffsetControl.getText()

    If Val(sValue) <> 0 Then
        oReasonControl.Visible = True
        oLabelReasonControl.Visible = True
    Else
        oReasonControl.Visible = False
        oLabelReasonControl.Visible = False
    End If
End Sub

Sub OffsetField_disposing(oEvent)
    ' Порожньо, обов’язково
End Sub

Sub PeopleTodayFilter()

    ' ==== Отримання документа і активного аркуша ====

    Dim oDoc As Object, oSheet As Object
    oDoc = ThisComponent
    oSheet = oDoc.CurrentController.ActiveSheet

    ' ==== Отримання поточної дати без часу ====

    Dim dToday As Double
    dToday = Int(Now())

    ' ==== Пошук останнього рядка за колонкою E ====

    Dim iLastRow As Long
    iLastRow = 3 ' Починаємо з рядка 4 (індексація з 0 — тобто 3 це рядок 4)

    Do While oSheet.getCellByPosition(0, iLastRow).getValue() <> 0 ' Перевіряємо колонку A
        iLastRow = iLastRow + 1
    Loop

    MsgBox "iLastRow " & iLastRow

    ' ==== Визначення діапазону для фільтру A4:E(останній рядок) ====

    Set oRange = oSheet.getCellRangeByPosition(0, 3, 4, iLastRow - 1)

    ' ==== Створення дескриптора фільтру ====

    Dim oFilterDesc As Object
    Set oFilterDesc = oRange.createFilterDescriptor(True)

    ' ==== Оголошення фільтруючих полів (2 умови) ====

    Dim oFilterFields(2) As New com.sun.star.sheet.TableFilterField

    ' ==== Перша умова: CheckIn < сьогодні + 1 ====

    With oFilterFields(0)
        .Field = 0 ' Колонка A
        .Operator = com.sun.star.sheet.FilterOperator.LESS_EQUAL ' Менше ніж (сьогодні + 1)
        .IsNumeric = True
        .NumericValue = dToday + 1 ' Щоб включити всіх хто заселився сьогодні
    End With

    ' ==== Друга умова: CheckOut >= сьогодні ====

    With oFilterFields(1)
        .Field = 4 ' Колонка E
        .Operator = com.sun.star.sheet.FilterOperator.GREATER_EQUAL ' Не пізніше ніж сьогодні
        .IsNumeric = True
        .NumericValue = dToday ' Тобто хто ще не виїхав
    End With

    ' Третя умова: D <> 7
	With oFilterFields(2)
    	.Field = 3 ' Колонка D
    	.Operator = com.sun.star.sheet.FilterOperator.NOT_EQUAL
    	.IsNumeric = True
    	.NumericValue = 7
	End With

    ' ==== Застосовуємо фільтр ====

    oFilterDesc.FilterFields = oFilterFields()
    oRange.filter(oFilterDesc)

    ' ==== Підрахунок видимих рядків і тих, у кого CheckOut сьогодні ====

    Dim result As Variant
    result = CountVisibleRows(oSheet, oRange)

    ' ==== Виведення результатів ====

	MsgBox "Порахуйте. Повинно бути " & result(0) & " " & PersonWord(result(0)) & "." & Chr(10) & _
       		Chr(10) & result(1) & " " & PersonWord(result(1)) & " до оплати, або на виселення." _
       		, 32, "Людей зараз:  " & result(0)
End Sub

' ==== Функція для відмінювання слова "особа" ====

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

Function CountVisibleRows(oSheet As Object, oRange As Object) As Variant
    Dim iRow As Long, iLastRow As Long
    Dim iVisibleCount As Long, iTermExpired As Long
    Dim oCellOut As Object

    dToday = Int(Now())

    iVisibleCount = 0
    iTermExpired = 0

    iStartRow = oRange.RangeAddress.StartRow
    iLastRow = oRange.RangeAddress.EndRow
    'MsgBox "iStartRow " & iStartRow & " iLastRow " & iLastRow

    For iRow = iStartRow To iLastRow
        Set oCellOut = oSheet.getCellByPosition(4, iRow) ' Колонка E

		If Not oSheet.getRows().getByIndex(iRow).IsFiltered And oCellOut.getValue() <> 0 Then
    		iVisibleCount = iVisibleCount + 1
		End If

		Set oCellOut = oSheet.getCellByPosition(4, iRow)

        If Int(oCellOut.getValue()) = dToday Then
    		iTermExpired = iTermExpired + 1
		End If

    Next iRow

    CountVisibleRows = Array(iVisibleCount, iTermExpired)
End Function

Sub ResetPeopleTodayFilter()
    Dim oDoc As Object, oSheet As Object
    Dim oRange As Object, oFilterDesc As Object

    ' ==== Отримання документа і активного аркуша ====

    oDoc = ThisComponent
    oSheet = oDoc.CurrentController.ActiveSheet

    ' ==== Діапазон, де був застосований фільтр ====

    Set oRange = oSheet.getCellRangeByName("A4:E1000")

    ' ==== Створення дескриптора фільтру ====

    Set oFilterDesc = oRange.createFilterDescriptor(True)

    ' ==== Скидання фільтру — просто видаляємо всі FilterFields ====

    oFilterDesc.FilterFields = Array()

    ' ==== Застосовуємо "порожній" фільтр для скидання ====

    oRange.filter(oFilterDesc)
End Sub