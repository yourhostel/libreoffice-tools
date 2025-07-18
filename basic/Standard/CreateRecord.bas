REM  *****  BASIC  *****

' CreateRecord.bas

Dim FormResult As Boolean

Sub StartCreate()
    ' False - не шукає першу порожню комірку по стовпцю А. (SelectFirstEmptyInA)
    ResetPeopleTodayFilter(False) 
    ShowForm(ACTION_CREATE)
End Sub

' =====================================================
' === Функція ShowForm ================================
' =====================================================
' → Запускає діалог введення нового запису.
' → Відображає форму, підключає слухачі, перевіряє введені дані та вставляє їх у таблицю.
' → Повертає рядок: "OK" — якщо дані збережені, "Cancel" — якщо відмінено.
Function ShowForm(sAction As String) As String  
    Dim oButtonInsert   As Object
    Dim oDialog         As Object
    Dim oListenerInsert As Object
    Dim sResult         As String
    
    FormResult = False 
    Set oDialog = CreateDialog(sAction)

    ' === Кнопка Вставити ===
    oButtonInsert = oDialog.getControl("InsertButton")

    ' === Обробник кнопки ===
    oListenerInsert = CreateUnoListener("InsertButton_", "com.sun.star.awt.XActionListener")
    oButtonInsert.addActionListener(oListenerInsert)
    oButtonInsert.Model.Tag = sAction
        
	' === Підключення слухача до OffsetField ===
	AddTextFieldsOffsetListener(oDialog)

    ' === Змінна результату ===
    sResult = "Cancel"   

    ' === Запуск діалогу ===
    If oDialog.execute() = 1 Then
        ' Натиснута кнопка (будь-яка) — тут можна перевіряти логіку
        sResult = "OK"
    End If
    
    If FormResult Then
        MsgDlg "Готово", String(20, " ") & "Дані збережено", False, 50, 120
	Else
    	MsgDlg "Скасовано", String(22, " ") & "Вихід без змін", False, 50, 120
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
    Dim oDoc    As Object
    Dim oSel    As Object
    Dim oDialog As Object
    Dim sAction As String
    sAction = oEvent.Source.Model.Tag
    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection
    oDialog = oEvent.Source.getContext()

    If Not CheckOccupiedPlace(oDialog, sAction) Then Exit Sub
    If Not OffsetReasonValidation(oSel, oDialog) Then Exit Sub
    If Not FinanceAreNumbersValidation(oDialog, "ExpenseField;IncomeField") Then Exit Sub
    If Not FinanceCommentValidation(oDialog) Then Exit Sub
    If Not PersonDataValidation(oDialog) Then Exit Sub
    If Not PhoneValidation(oDialog) Then Exit Sub
    If Not BirthDateValidation(oDialog) Then Exit Sub
    If Not PassportDataValidation(oDialog) Then Exit Sub

    ' якщо редагування — одразу пишемо історію
    If sAction = ACTION_EDIT Then
        If Not AppendHistory(oSel.RangeAddress.StartRow) Then 
            Exit Sub
        End If    
    End If
  	   
    OffsetReasonInsertion(oSel, oDialog)	' Q, P      причина зсуву зсув
    DateRangeInsertion(oSel, oDialog)	    ' A, E, O   заселення, виселення, створено
    CodeInsertion(oSel, oDialog)            ' D         код
    PersonDataInsertion(oSel, oDialog)      ' B, C      прізвище, ім'я по батькові
    PaidInsertion(oSel, oDialog)            ' F         сплачено
    FinanceInsertion(oSel, oDialog)         ' G, H, I   видаток, прихід, коментар
    PhoneInsertion(oSel, oDialog)           ' J         телефон
    PassportBirthInsertion(oSel, oDialog)   ' K, L      паспортні дані, дата народження
    HostelInsertion(oSel , oDialog)         ' N         хостел
    PlaceInsertion(oSel, oDialog)           ' R         місце
    		
    FormResult = True   ' Ставимо True тільки якщо валідація пройшла та вставка відпрацювала коректно 
           
    oDialog.endExecute()
End Sub

' =====================================================
' === Процедура InsertButton_disposing ===============
' =====================================================
' → Викликається при видаленні слухача з кнопки InsertButton.
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub InsertButton_disposing(oEvent As Object)
End Sub

' =====================================================
' === Функція OffsetReasonValidation ==================
' =====================================================
' → Перевіряє, що якщо Offset ≠ 0, то заповнено поле Reason.
' → Повертає True — якщо умова виконана, False — якщо ні.
Function OffsetReasonValidation(oSel As Object, oDialog As Object) As Boolean
    ' ==== Читаємо значення полів ====
    Dim sOffset As String
    Dim sReason As String
    
    sOffset = oDialog.getControl("OffsetField").getText()
    sReason = oDialog.getControl("ReasonField").getText()

    ' ==== Перевірка ====
    If Val(sOffset) <> 0 And Trim(sReason) = "" Then
        MsgDlg "Увага!", "Поле 'Причина зсуву' не може бути порожнім при ненульовому зсуві!", False, 50, 225
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
    Dim sOffset As String, sReason As String
    oSheet = oSel.Spreadsheet
    sOffset = oDialog.getControl("OffsetField").getText()
    sReason = oDialog.getControl("ReasonField").getText()
    oSheet.getCellByPosition(16, oSel.CellAddress.Row).setValue(Val(sOffset)) ' Q
    
    If Val(sOffset) = 0 Then
        oSheet.getCellByPosition(15, oSel.CellAddress.Row).setString("")      ' P
        Exit Sub
    End If
    oSheet.getCellByPosition(15, oSel.CellAddress.Row).setString(sReason)     ' P
End Sub

' =====================================================
' === Процедура DateRangeInsertion ====================
' =====================================================
' → Вставляє дати заселення, виселення та створення у таблицю.
' → Форматує дати у вигляді "DD.MM.YYYY" та "DD.MM.YYYY HH:MM".
Sub DateRangeInsertion(oSel As Object, oDialog As Object)
    Dim nOffset        As Integer
    Dim dBaseDate      As Double
    Dim dEndDate       As Double
    Dim nDuration      As Integer
    Dim oSheet         As Object
    Dim oCursorAddress As Object
    Dim oCheckInCell   As Object
    Dim oCheckOutCell  As Object
    Dim oCreatedCell   As Object
    Dim bWasProtected  As Boolean
      
    ' ==== Читаємо значення Offset і Duration ====
    nOffset = Val(oDialog.getControl("OffsetField").getText())
    nDuration = Val(oDialog.getControl("DurationCombo").getText())

    ' ==== Отримуємо таблицю і адресу ====
    oSheet = oSel.Spreadsheet
    oCursorAddress = oSel.CellAddress

    ' ==== Обчислення базової і кінцевої дати ====
    dBaseDate = CDate(oDialog.getControl("CurrentDateField").getText())
    dBaseDate = dBaseDate + nOffset
    dEndDate = dBaseDate + nDuration

    ' ==== Вставка дати заселення в колонку A ====
    oCheckInCell = oSheet.getCellByPosition(0, oCursorAddress.Row)              ' A
    oCheckInCell.setValue(Cdate(Format(dBaseDate, "DD.MM.YYYY")))

    ' ==== Вставка дати виселення в колонку E ====
    oCheckOutCell = oSheet.getCellByPosition(4, oCursorAddress.Row)             ' E
    oCheckOutCell.setValue(Cdate(Format(dEndDate, "DD.MM.YYYY")))

    ' ==== Вставка дати створення з часом в колонку O ====
    oCreatedCell = oSheet.getCellByPosition(14, oCursorAddress.Row)             ' O
    oCreatedCell.setValue(Cdate(Format(Now, "DD.MM.YYYY HH:MM:SS")))
 
    ' ==== Вставка терміна в колонку Т ====
    oSheet.getCellByPosition(19, oSel.CellAddress.Row).setValue(Val(nDuration)) ' Т
    
    ' ==== Вставка id у колонку U ====
    SetNextId(oSheet, oSel)                                                     ' U
    
    ' ==== Якщо є зсув — застосовуємо стиль "створено" ====
    bWasProtected = oSheet.IsProtected
    
    ' ==== Якщо захищений — знімаємо захист ====
    If bWasProtected Then oSheet.unprotect(NEGET_RULES)
    
    If nOffset <> 0 Then oCreatedCell.CellStyle = "створено" Else: oCreatedCell.CellStyle = "Типовий"
    
    ' ==== Повертаємо захист назад ====
    If bWasProtected Then oSheet.protect(NEGET_RULES)
End Sub

' =====================================================
' === Функція PersonDataValidation ====================
' =====================================================
' → Перевіряє заповнення полів Прізвище, Ім'я та По батькові.
' → Повертає True — якщо всі заповнені, False — якщо ні.
Function PersonDataValidation(oDialog As Object) As Boolean
    Dim sLastName   As String
    Dim sFirstName  As String
    Dim sPatronymic As String
    
    ' ==== Отримання значень полів ====
    sLastName   = Trim(oDialog.getControl("LastNameField").getText())
    sFirstName  = Trim(oDialog.getControl("FirstNameField").getText())
    sPatronymic = Trim(oDialog.getControl("PatronymicField").getText())
    
    ' ==== Перевірка на порожні значення ====
    If sLastName = "" Or sFirstName = "" Or sPatronymic = "" Then
        MsgDlg "Увага!", "Необхідно заповнити всі поля: Прізвище, Ім'я, По батькові.", False, 50, 200
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
    Dim sLastName   As String
    Dim sFirstName  As String
    Dim sPatronymic As String
    Dim sFullName   As String
    Dim oSheet      As Object    
    ' ==== Отримання значень ====
    sLastName   = Capitalize(Trim(oDialog.getControl("LastNameField").getText()))
    sFirstName  = Capitalize(Trim(oDialog.getControl("FirstNameField").getText()))
    sPatronymic = Capitalize(Trim(oDialog.getControl("PatronymicField").getText()))
    
    ' ==== Формування повного імені ====
    sFullName = sFirstName & " " & sPatronymic
    
    ' ==== Вставка у таблицю ====
    oSheet = oSel.Spreadsheet
    
    ' ==== 'прізвище' ====
    oSheet.getCellByPosition(1, oSel.CellAddress.Row).setString(sLastName)   ' B
    
    ' ==== 'ім'я та по батькові' ====
    oSheet.getCellByPosition(2, oSel.CellAddress.Row).setString(sFullName)   ' C
End Sub

' =====================================================
' === Процедура PaidInsertion =========================
' =====================================================
' → Вставляє значення з поля PaidField у таблицю (стовпець F).
Sub PaidInsertion(oSel As Object, oDialog As Object)
    Dim oSheet As Object
    Dim dPrice As Double
    
    ' ==== Отримання листа ====
    oSheet = oSel.Spreadsheet
    
    ' ==== Отримання значення з поля форми ====
    dPrice = Val(oDialog.getControl("PaidField").getText())
    
    ' ==== 'сплачено' (індекс 5) ====
    oSheet.getCellByPosition(5, oSel.CellAddress.Row).setValue(Val(dPrice))  ' F
End Sub

' =====================================================
' === Функція FinanceAreNumbersValidation =============
' =====================================================
' → Перевіряє, що зазначені фінансові поля містять числові значення.
' → Повертає True — якщо все коректно, False — якщо є помилки.
Function FinanceAreNumbersValidation(oDialog As Object, sFields As String) As Boolean
    Dim aFields() As String
    Dim i         As Integer
    Dim sValue    As String
    Dim bIsValid  As Boolean
    
    aFields = Split(sFields, ";")
    bIsValid = True

    For i = LBound(aFields) To UBound(aFields)
        sValue = Trim(oDialog.getControl(aFields(i)).getText())
        If Trim(sValue) = "" Or Not IsNumeric(sValue) Then
    		Dim Map As Variant
    		Map = GetFieldToColumnMap()
    		MsgDlg "Увага!", "Поле """ & MapGet(Map, aFields(i)) & """ повинно містити число.", False, 50, 145			
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
    dIncome  = Val(oDialog.getControl("IncomeField").getText())
    sComment = Trim(oDialog.getControl("CommentField").getText())

    If (dExpense <> 0 Or dIncome <> 0) And sComment = "" Then
        Dim Map As Variant
        Map = GetFieldToColumnMap()

        Dim fieldName As String
        If dExpense <> 0 Then fieldName = "ExpenseField" Else: fieldName = "IncomeField"
        
    	MsgDlg "Увага!", "Поле """ & MapGet(Map, fieldName) & """ заповнено, напишіть коментар.", False, 50, 165
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
    Dim oSheet   As Object
    Dim dExpense As Double
	Dim dIncome  As Double
    Dim dComment As String
    
    ' ==== Отримання листа ====
    oSheet = oSel.Spreadsheet

    ' ==== Отримання значення з поля форми Expense, Income ====
    dExpense = Val(oDialog.getControl("ExpenseField").getText())
	dIncome  = Val(oDialog.getControl("IncomeField").getText())
	dComment = oDialog.getControl("CommentField").getText()
		
    ' ==== Вставка 'видаток' (індекс 6) ====
    oSheet.getCellByPosition(6, oSel.CellAddress.Row).setValue(Val(dExpense)) ' G 
    
    ' ==== Вставка 'прихід' (індекс 7) ==== 
    oSheet.getCellByPosition(7, oSel.CellAddress.Row).setValue(Val(dIncome))  ' H
    
    ' ==== Вставка 'коментар' (індекс 8) ====
    oSheet.getCellByPosition(8, oSel.CellAddress.Row).setString(dComment)     ' I
End Sub

' =====================================================
' === Функція IsPhoneMinimalValid =====================
' =====================================================
' → Перевіряє мінімальну коректність номера телефону.
' → Повертає True — якщо виглядає коректним, False — якщо ні.
' → Базова перевірка на формат без помилок типу "123", "asd", "0".
Function IsPhoneMinimalValid(sPhone As String) As Boolean
    Dim sClean As String
    Dim nLen   As Integer
    
    ' ==== Очистка від пробілів, тире, дужок ====
    sClean = Replace(sPhone, " ", "")
    sClean = Replace(sClean, "-", "")
    sClean = Replace(sClean, "(", "")
    sClean = Replace(sClean, ")", "")
    
    ' ==== Перевірка на плюс на початку ====
    If Left(sClean, 1) = "+" Then sClean = Mid(sClean, 2)
    
    ' ==== Перевірка: залишилися лише цифри ====
    If Not IsNumeric(sClean) Then
        IsPhoneMinimalValid = False
        Exit Function
    End If
    
    ' ==== Перевірка довжини номера ====
    ' → Мінімум 8 цифр (розумний мінімум для більшості країн)
    ' → Максимум 15 цифр (згідно стандарту ITU E.164)
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
	    MsgDlg "Поле 'Телефон' заповнене некоректно.", "Приклади валідних записів:" & Chr(10) & _
	    Chr(10) & _
	    " +380671234567, 0671234567, +12025550198, 441234567890", False, 65, 190
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
    Dim oSheet As Object
    Dim sPhone As String
    oSheet = oSel.Spreadsheet
    sPhone = oDialog.getControl("PhoneField").getText()
    ' Вставка у колонку 'телефон'  J (індекс 9)
    oSheet.getCellByPosition(9, oSel.CellAddress.Row).setString(sPhone) ' J
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
    MsgDlg "Паспортні дані потребують уточнення.", _
    String(115, " ") & "ІНСТРУКЦІЯ." & Chr(10) & _
    "Через кому с пробілом 5 полів:" & Chr(10) & _
    "_______, __________, _______, dd.mm.yyyy, _______________" & Chr(10) & _
    Chr(10) & _
    "ПРАВИЛО: Номер(≥7), Ким видан(≥10), Де(≥7), Коли(dd.mm.yyyy), Прописка(≥15)" & Chr(10) & _
    Chr(10) & _
    "Де (≥n) не менше або дорівнює кількості символів ""n""" & Chr(10) & _
    Chr(10) & _
    "ПРИКЛАДИ:" & Chr(10) & _ 
    Chr(10) & _
    "ВР577927, Виноградовским РО ГУМС Украины в Закарпатской обл, с.Перикиво ул.Тихая 88, 27.05.2001, с.Перикиново ул.Тихая 5" & Chr(10) & _
    Chr(10) & _
    "МР459000, Ивановским РС УГМС Украины в Херсонской обл, пгт.Ивановка ул.Светская 3 кв.10, 29.05.2010, пгт.Ивановка ул.Светская 3 кв.100" & Chr(10) & _
    Chr(10) & _
    "ІНОЗЕМНІ ПРИКЛАДИ:" & Chr(10) & _
    Chr(10) & _
    "P USA 566900551, SEE PAGE51, _______, 26.07.2030, _______________" & Chr(10) & _
    Chr(10) & _
    "PA2727784, SEE PAGE51, _______, 26.07.2030, Bendigo Australia" & Chr(10) & _
    Chr(10) & _
    "T4422644 P IND, __________, 13.01.1992, 20.02.2029, Patiala Punjab Chandigarh" & Chr(10) & _
    Chr(10) & _
    "Використовуйте нижнє підкреслення (Shift + «-») для заповнення відсутніх полів для нестандартних паспортних даних. Четверте поле обов'язково повинно мати формат дати dd.mm.yyyy («d» - символ дня, “m” - місяця, «y» - року). Мінімальна кількість символів 7, 10, 7, dd.mm.yyyy, 15" _
    , True, 200, 350          
End Sub

' =====================================================
' === Функція BirthDateValidation =====================
' =====================================================
' → Перевіряє правильність формату дати народження.
' → Якщо не валідно — показує діалог із підказкою.
Function BirthDateValidation(oDialog As Object) As Boolean
    If Not DateFormatValidation(oDialog.getControl("BirthDateField").getText()) Then
        MsgDlg "Формат дати народження не валідний", "Подання має бути в такому вигляді: dd.mm.yyyy", False, 50, 165           
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
    Dim oSheet       As Object  
    Dim sPassport    As String
    Dim sBirthDate   As String
    Dim dBirthDate   As Date
    Dim oFormats     As Object, oLocale As New com.sun.star.lang.Locale
    Dim nFormatDate  As Long

    ' === Отримуємо дані з форми ===
    oSheet = oSel.Spreadsheet 
    sPassport = oDialog.getControl("PassportField").getText()
    sBirthDate = oDialog.getControl("BirthDateField").getText()   

    ' === Записуємо паспорт як текст ===
    oSheet.getCellByPosition(10, oSel.CellAddress.Row).setString(sPassport)  ' K

    ' === Перетворюємо дату народження у формат Date ===
    dBirthDate = CDate(sBirthDate)

    ' === Встановлюємо формат дати ===
    oFormats = ThisComponent.getNumberFormats()
    oLocale.Language = "uk"
    oLocale.Country = "UA"

    nFormatDate = oFormats.queryKey("DD.MM.YYYY", oLocale, True)
    If nFormatDate = -1 Then
        nFormatDate = oFormats.addNew("DD.MM.YYYY", oLocale)
    End If

    ' === Записуємо дату народження як значення + формат ===
    With oSheet.getCellByPosition(11, oSel.CellAddress.Row)                  ' L
        .setValue(dBirthDate)
        .NumberFormat = nFormatDate
    End With
End Sub

' =====================================================
' === Процедура HostelInsertion =======================
' =====================================================
' → Вставляє значення хостелу у таблицю (стовпець N).
Sub HostelInsertion(oSel As Object, oDialog As Object)
    Dim oSheet As Object
    oSheet = oSel.Spreadsheet
    oSheet.getCellByPosition(13, oSel.CellAddress.Row).setString(HOSTEL)     ' N
End Sub

' =====================================================
' === Процедура CodeInsertion =========================
' =====================================================
' → Вставляє код у таблицю (стовпець D).
Sub CodeInsertion(oSel As Object, oDialog As Object)
    Dim oSheet  As Object
    Dim sCode   As String
    oSheet = oSel.Spreadsheet 
    sCode = oDialog.getControl("CodeCombo").getText()
    oSheet.getCellByPosition(3, oSel.CellAddress.Row).setValue(Val(sCode))   ' D
End Sub

' =====================================================
' === Процедура PlaceInsertion ========================
' =====================================================
' → Вставляє номер місця в таблицю (стовпець R) як число.
Sub PlaceInsertion(oSel As Object, oDialog As Object)
    Dim oSheet  As Object
    Dim sPlace  As String     
    oSheet = oSel.Spreadsheet  
    sPlace = oDialog.getControl("PlaceCombo").getText()
    oSheet.getCellByPosition(17, oSel.CellAddress.Row).setValue(Val(sPlace)) ' R
End Sub

' =====================================================
' === Функція CreateDialog ============================
' =====================================================
' → Створює та налаштовує діалогову форму введення нового запису.
' → Додає всі поля, мітки, кнопки, слухачі та повертає готовий діалог.
Function CreateDialog(sAction As String) As Object
    Dim oDialog        As Object
    Dim oDialogModel   As Object
    Dim mInitialValues As Variant
    Dim sTitle         As String
    
    If sAction = ACTION_CREATE Then 
        sTitle = "Новий запис" 
    Else  
        sTitle = "Редагування запису"  
    End If      
    
    oDialog = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    oDialogModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")  
    oDialog.setModel(oDialogModel)
    
    mInitialValues = FormInitialization(sAction)
    
    ' ==== Параметри діалогу ====    
    With oDialogModel
        .PositionX = 100
        .PositionY = 100
        .Width = 350
        .Height = 250
        .Title = sTitle
    End With
    
    AddLogo (oDialogModel ,"logo", 290, 5, 50, 45)
     
    ' ==== Поточна дата і час ====
    FieldTemplate(oDialogModel, "CurrentDate", "Поточні дата і час:", 10, 15, MapGet(mInitialValues, "створено"), 65, 65, True)
    ' ==== Зсув у днях ====
    FieldTemplate(oDialogModel, "Offset", "Зсув у днях:", 10, 45, MapGet(mInitialValues, "зсув"), 50, 50)
    ' ==== Кількість днів ====
    ComboBoxTemplate(oDialogModel, "Duration", "Кількість днів:", 65, 45, MapGet(mInitialValues, "тривалість"), 50, 50, VALID_DURATIONS)
    AddDurationComboListeners(oDialog)
    ' ==== Причини зсуву ====
    FieldTemplate(oDialogModel, "Reason", "Причина зсуву:", 120, 45, MapGet(mInitialValues, "причина зсуву"), 50, 165)
    ' ==== Поля визначення ====
    ComboBoxTemplate(oDialogModel, "Code", "Код:", 10, 75, MapGet(mInitialValues, "код"), 50, 50, LIST_OF_CODES)
    AddCodeComboListeners(oDialog)
    ComboBoxTemplate(oDialogModel, "Place", "Місце:", 65, 75, MapGet(mInitialValues, "місце"), 50, 50, LIST_OF_PLACES)
    AddPlaceComboListeners(oDialog)
    ' ==== Фінансові поля ====
    FieldTemplate(oDialogModel, "Paid", "Сплачено:", 120, 75, MapGet(mInitialValues, "сплачено"), 50, 50, True)
    FieldTemplate(oDialogModel, "Expense", "Видаток:", 235, 75, MapGet(mInitialValues, "видаток"), 50, 50)
    FieldTemplate(oDialogModel, "Income", "Прихід:", 290, 75, MapGet(mInitialValues, "прихід"), 50, 50)
    ' ==== Інші дані ====
    FieldTemplate(oDialogModel, "Comment", "Коментар:", 10, 105, MapGet(mInitialValues, "коментар"), 70, 330)
    ' ==== Персональні дані ====
    FieldTemplate(oDialogModel, "LastName", "Прізвище:", 10, 135, MapGet(mInitialValues, "прізвище"), 70, 100)
    FieldTemplate(oDialogModel, "FirstName", "Ім'я:", 125, 135, MapGet(mInitialValues, "ім’я"), 70, 100)
    FieldTemplate(oDialogModel, "Patronymic", "По батькові:", 240, 135, MapGet(mInitialValues, "по батькові"), 70, 100)
    ' ==== Контактна інформація ====
    FieldTemplate(oDialogModel, "Phone", "Телефон:", 10, 165, MapGet(mInitialValues, "телефон"), 70, 100)
    FieldTemplate(oDialogModel, "BirthDate", "Дата народження:", 125, 165, MapGet(mInitialValues, "дата народження"), 100, 100)
    ' ==== Документ ====
    FieldTemplate(oDialogModel, "Passport", "Паспортні дані:", 10, 195, MapGet(mInitialValues, "паспортні дані"), 100, 330)
    
    CalculatePaidFieldWithPlace(oDialog)
    
    If sAction = ACTION_CREATE Then
        UpdatePlaceCombo(oDialog, ACTION_CREATE)
    ElseIf sAction = ACTION_EDIT Then
        UpdatePlaceCombo(oDialog, ACTION_EDIT, MapGet(mInitialValues, "місце"))
    End If 
      
    ' ==== Кнопка вставки ====    
    AddButton(oDialogModel, "InsertButton", ChoiceButtonName(sAction), 150, 225, 60, 14)
    
    oDialog.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null) 
    ' ==== Повертаємо ====
    CreateDialog = oDialog
End Function

' =====================================================
' === Функція FormInitialization ======================
' =====================================================
' → Ініціалізує дані для форми залежно від дії (`ACTION_CREATE` або `ACTION_EDIT`).
' → Якщо `CREATE` — заповнює дефолтні значення (новий запис).
' → Якщо `EDIT` — зчитує існуючі дані з таблиці й обчислює зсув, тривалість тощо.
' → Повертає Map зі всіма полями для заповнення форми.
Function FormInitialization(sAction As String) As Variant
    Dim data      As Variant
    Dim Fields    As Variant
    Dim sPrice    As String   ' ціна з аркуша
    Dim sDateTime As String   ' поточна дата і час
    Dim sOffset   As String   ' зсув
    Dim sDuration As String   ' тривалість
    Dim sReason   As String   ' причина зсуву
    Dim sCode     As String   ' код
    Dim sPlace    As String   ' місце
    Dim sPaid     As String   ' сплачено
    Dim sExpense  As String   ' видаток
    Dim sIncome   As String   ' прихід
    Dim sComment  As String   ' коментар
    Dim sLast     As String   ' прізвище
    Dim sFirst    As String   ' ім’я
    Dim sPatr     As String   ' по батькові
    Dim sPhone    As String   ' телефон
    Dim sBirth    As String   ' дата народження
    Dim sPass     As String   ' паспортні дані

    data = CreateMap()
        
    sPrice = ThisComponent.Sheets.getByName("price1").getCellByPosition(1, 1).getValue()

    If sAction = ACTION_CREATE Then
        sDateTime = Format(Now, "DD.MM.YYYY HH:MM")
        sOffset = "0"
        sDuration = "1"
        sReason = ""
        sCode = "1"
        sPlace = "1"
        sPaid = sPrice
        sExpense = "0"
        sIncome = "0"
        sComment = ""
        sLast = ""
        sFirst = ""
        sPatr = "_"
        sPhone = ""
        sBirth = ""
        sPass = ""

    ElseIf sAction = ACTION_EDIT Then 
        Fields = ReadFromTable()
        sDateTime = Format(CDate(MapGet(Fields, "заселення")) - Val(MapGet(Fields, "зсув")), "DD.MM.YYYY")
        sOffset = MapGet(Fields, "зсув")

        ' обчислюємо тривалість
        Dim dCheckIn As Date, dCheckOut As Date, nDuration As Long
        dCheckIn = CDate(MapGet(Fields, "заселення"))
        dCheckOut = CDate(MapGet(Fields, "виселення"))
        nDuration = DateDiff("d", dCheckIn, dCheckOut)
        sDuration = CStr(nDuration)

        sReason = MapGet(Fields, "причина зсуву")
        sCode = MapGet(Fields, "код")
        sPlace = MapGet(Fields, "місце")
        sPaid = MapGet(Fields, "сплачено")
        sExpense = MapGet(Fields, "видаток")
        sIncome = MapGet(Fields, "прихід")
        sComment = MapGet(Fields, "коментар")
        sLast = MapGet(Fields, "прізвище")

        ' Ділемо ім’я та по батькові
        Dim aNameParts As Variant
        aNameParts = Split(MapGet(Fields, "ім'я по батькові"), " ")
        If UBound(aNameParts) >= 0 Then sFirst = aNameParts(0)
        If UBound(aNameParts) >= 1 Then sPatr = aNameParts(1) Else sPatr = ""

        sPhone = MapGet(Fields, "телефон")
        sBirth = MapGet(Fields, "дата народження")
        sPass = MapGet(Fields, "паспортні дані")
    End If

    ' === записуємо в Map ===
    MapPut data, "створено", sDateTime
    MapPut data, "зсув", sOffset
    MapPut data, "тривалість", sDuration
    MapPut data, "причина зсуву", sReason
    MapPut data, "код", sCode
    MapPut data, "місце", sPlace
    MapPut data, "сплачено", sPaid
    MapPut data, "видаток", sExpense
    MapPut data, "прихід", sIncome
    MapPut data, "коментар", sComment
    MapPut data, "прізвище", sLast
    MapPut data, "ім’я", sFirst
    MapPut data, "по батькові", sPatr
    MapPut data, "телефон", sPhone
    MapPut data, "дата народження", sBirth
    MapPut data, "паспортні дані", sPass

    FormInitialization = data
End Function

' =====================================================
' === Функція ChoiceButtonName ========================
' =====================================================
' → Вибирає текст для кнопки діалогу залежно від дії (`CREATE`, `EDIT`).
' → Повертає рядок: "Вставити", "Зберегти", або "OK".
Function ChoiceButtonName(sAction As String) As String
    Dim sAct As String

    Select Case sAction
        Case ACTION_CREATE
            sAct = "Вставити"

        Case ACTION_EDIT
            sAct = "Зберегти"

        Case Else
            sAct = "OK" 
    End Select

    ChoiceButtonName = sAct
End Function

