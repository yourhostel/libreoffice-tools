REM  *****  BASIC  *****

' Utilities.bas

' ==== Службові функції, утиліти, структура Map ======
' =====================================================

' =====================================================
' === Функція Capitalize =============================
' =====================================================
' → Перетворює рядок у формат:
'   Перша літера велика, інші — маленькі.
' → Якщо рядок порожній — повертає порожній.
Function Capitalize(sText As String) As String
    If Len(sText) = 0 Then
        Capitalize = ""
    Else
        Capitalize = UCase(Left(sText, 1)) & LCase(Mid(sText, 2))
    End If
End Function

' =====================================================
' === Функція LockFields ==============================
' =====================================================
' → Блокує поля на формі (робить ReadOnly).
' → Параметр sFieldNames — перелік імен полів через крапку з комою.
' → Наприклад:
'    LockFields(oEvent, "DurationCombo;OffsetField")
Sub LockFields(oEvent As Object, sFieldNames As String)
	Dim FieldNames() As String
	FieldNames = Split(sFieldNames, ";")

    ' Отримуємо діалог з контексту події
    Dim oDialog As Object
    oDialog = oEvent.Source.getContext()

    ' Цикл по масиву імен полів
    Dim i As Integer
    For i = LBound(FieldNames) To UBound(FieldNames)
        oDialog.getControl(FieldNames(i)).Model.ReadOnly = True
    Next i
End Sub

' =====================================================
' ================ Структура Map ======================
' =====================================================
' ==== Простий аналог Map/Dictionary для LibreOffice ==
' ==== Реалізовано на базі масиву пар Key-Value =======
' =====================================================

' ==== Оголошення структури Key-Value =====
Type KeyValuePair
    Key As String
    Value As String
End Type

' =====================================================
' === Функція CreateMap ===============================
' =====================================================
' → Створює пусту мапу (масив пар Key-Value).
' → Повертає Variant-масив.
Function CreateMap() As Variant
    CreateMap = Array()
End Function

' =====================================================
' === Процедура MapPut ================================
' =====================================================
' → Додає пару (Key, Value) в мапу.
' → Якщо ключ існує — оновлює його значення.
' → Якщо ключа нема — додає новий.
Sub MapPut(ByRef Map As Variant, ByVal Key As String, ByVal Value As String)
    Dim i As Integer
    For i = LBound(Map) To UBound(Map)
        If Map(i).Key = Key Then
            Map(i).Value = Value
            Exit Sub
        End If
    Next i

    ' Додаємо новий ключ
    Dim NewPair As KeyValuePair
    NewPair.Key = Key
    NewPair.Value = Value

    If IsEmpty(Map) Then
        Map = Array(NewPair)
    Else
        Map = AppendArray(Map, NewPair)
    End If
End Sub

' =====================================================
' === Функція MapGet ==================================
' =====================================================
' → Повертає значення за ключем.
' → Якщо ключ не знайдено — повертає порожній рядок "".
Function MapGet(ByVal Map As Variant, ByVal Key As String) As String
    Dim i As Integer
    For i = LBound(Map) To UBound(Map)
        If Map(i).Key = Key Then
            MapGet = Map(i).Value
            Exit Function
        End If
    Next i
    MapGet = "" ' Якщо не знайдено
End Function

' =====================================================
' === Функція MapHasKey ===============================
' =====================================================
' → Перевіряє наявність ключа в мапі.
' → Повертає True — якщо ключ є.
' → Повертає False — якщо ключа нема.
Function MapHasKey(ByVal Map As Variant, ByVal Key As String) As Boolean
    Dim i As Integer
    For i = LBound(Map) To UBound(Map)
        If Map(i).Key = Key Then
            MapHasKey = True
            Exit Function
        End If
    Next i
    MapHasKey = False
End Function

' =====================================================
' === Функція MapGetByIndex ===========================
' =====================================================
' → Повертає пару Key і Value за індексом.
' → Повертає Variant(0)=Key, Variant(1)=Value.
' → Якщо індекс невалідний — повертає ("","").
Function MapGetByIndex(ByVal Map As Variant, ByVal Index As Integer) As Variant
    Dim result(1) As String

    If Index >= LBound(Map) And Index <= UBound(Map) Then
        result(0) = Map(Index).Key
        result(1) = Map(Index).Value
    Else
        result(0) = ""
        result(1) = ""
    End If

    MapGetByIndex = result
End Function

' =====================================================
' === Функція AppendArray =============================
' =====================================================
' → Внутрішня допоміжна функція.
' → Додає новий елемент до масиву.
Function AppendArray(arr As Variant, item As Variant) As Variant
    Dim l As Long
    l = -1
    On Error Resume Next
    l = UBound(arr) + 1
    On Error GoTo 0

    If l = 0 And IsEmpty(arr) Then
        AppendArray = Array(item)
    Else
        Dim temp() As Variant
        ReDim temp(0 To l)
        Dim i As Integer
        For i = 0 To l - 1
            temp(i) = arr(i)
        Next i
        temp(l) = item
        AppendArray = temp
    End If
End Function

Sub MapClear(ByRef Map As Variant)
    Map = CreateMap()
End Sub

' =====================================================
' === Процедура CalculatePaidFieldWithPlace ==========
' =====================================================
' → Обчислює вартість проживання за кодом, місцем і кількістю днів.
' → Дані беруться з листа price[N], де N — код.
' → Встановлює обчислену суму у поле Paid.
Sub CalculatePaidFieldWithPlace(oDialog As Object)

    ' ==== Отримуємо значення з ComboBox ====
    Dim sDuration As String
    Dim sCode As String
    Dim sPlace As String

    sDuration = oDialog.getControl("DurationCombo").getText()
    sCode = oDialog.getControl("CodeCombo").getText()
    sPlace = oDialog.getControl("PlaceCombo").getText()

    ' ==== Перетворення у числа ====
    Dim nDuration As Long
    Dim nCode As Long
    Dim nPlace As Long

    nDuration = Val(sDuration)
    nCode = Val(sCode)
    nPlace = Val(sPlace)

    ' ==== Перевірка на чітність місця ====
    Dim bEven As Boolean
    bEven = (nPlace Mod 2 = 0) ' True — нижнє місце (парне), False — верхнє (непарне)

    ' ==== Відкриваємо лист відповідно до коду ====
    Dim oDoc As Object, oSheet As Object
    Dim sSheetName As String
    oDoc = ThisComponent
    sSheetName = "price" & nCode

    ' ==== Якщо листа з таким кодом немає → підставляємо price8 ====
    If Not oDoc.Sheets.hasByName(sSheetName) Then
        sSheetName = "price8"
    End If

    On Error GoTo ErrorHandler
    Set oSheet = oDoc.Sheets.getByName(sSheetName)
    On Error GoTo 0

    ' ==== Пошук ціни ====
    Dim dPrice As Double
    dPrice = 0

    If bEven Then
        ' ==== Нижнє місце → таблиця A:B ====
        For iRow = 1 To MAX_SEARCH_RANGE_IN_PRICE
            If oSheet.getCellByPosition(0, iRow).getValue() = nDuration Then
                dPrice = oSheet.getCellByPosition(1, iRow).getValue()
                Exit For
            End If
        Next iRow
    Else
        ' ==== Верхнє місце → таблиця D:E ====
        For iRow = 1 To MAX_SEARCH_RANGE_IN_PRICE
            If oSheet.getCellByPosition(3, iRow).getValue() = nDuration Then
                dPrice = oSheet.getCellByPosition(4, iRow).getValue()
                Exit For
            End If
        Next iRow
    End If

    ' ==== Встановлюємо ціну у поле Paid ====
    oDialog.getControl("PaidField").setText(CStr(dPrice))

    Exit Sub

ErrorHandler:
	ShowDialog "Помилка", "Не знайдено лист 'price" & nCode & "'. Перевірте правильність коду."
End Sub

' =====================================================
' === Процедура UpdatePlaceCombo ======================
' =====================================================
' → Оновлює список місць у ComboBox "PlaceCombo" згідно з кодом.
' → Витягує значення з колонки G на листі price[N].
' → Встановлює перше доступне місце за замовчуванням.
Sub UpdatePlaceCombo(oDialog As Object, _
            Optional sAction As String, _
            Optional sPlace As String)
    ' ==== Ініціалізуємо опційні аргументи ====
    If IsMissing(sAction) Then sAction = ""
    If IsMissing(sPlace) Then sPlace = ""

    Dim oDoc As Object, oSheet As Object
    Dim oCombo As Object
    Dim iRow As Long
    Dim aPlaces() As String    ' масив місць (рядки)
    Dim iCount As Integer      ' кількість знайдених місць
    Dim nCode As Long          ' код, вибраний у CodeCombo
    Dim sSheetName As String   ' назва аркуша з цінами

    iCount = 0
    ' ==== Отримуємо вибраний код та формуємо ім’я аркуша ====
    nCode = Val(oDialog.getControl("CodeCombo").getText())
    oDoc = ThisComponent
    sSheetName = "price" & nCode

    ' ==== Якщо аркуш з таким кодом не знайдено — fallback на price8 ====
    If Not oDoc.Sheets.hasByName(sSheetName) Then
        sSheetName = "price8"
    End If

    On Error GoTo ErrorHandler
    Set oSheet = oDoc.Sheets.getByName(sSheetName)
    ' ==== Збираємо всі непорожні значення зі стовпця G ====
    For iRow = 1 To MAX_SEARCH_RANGE_IN_PRICE
        Dim nPlace As Long
        nPlace = oSheet.getCellByPosition(6, iRow).getValue()

        If nPlace <> 0 Then
            ReDim Preserve aPlaces(iCount)    ' розширюємо масив
            aPlaces(iCount) = CStr(nPlace)    ' зберігаємо як рядок
            iCount = iCount + 1
        End If
    Next iRow
    ' ==== Оновлюємо модель ComboBox ====
    Set oCombo = oDialog.getControl("PlaceCombo")
    oCombo.Model.StringItemList = aPlaces
	' ==== Вибираємо значення для ComboBox ====
	If iCount > 0 Then
        If sAction = ACTION_CREATE Or sPlace = "" Then
            ' якщо створення або значення не задано — ставимо перший
            oCombo.Model.Text = aPlaces(0)
        Else
            ' якщо редагування і передано sPlace — ставимо його
            oCombo.Model.Text = sPlace
        End If
    End If
	' ==== Встановлюємо висоту випадаючого списку залежно від кількості ====
    If iCount > 12 Then
    	PLACE_COMBO_HEIGHT = 110
	Else
    	PLACE_COMBO_HEIGHT = 15 + iCount * 9
	End If

    Exit Sub

ErrorHandler:
    ShowDialog "Помилка", "Не вдалося завантажити список місць з аркуша 'price" & nCode & "'"
End Sub

' =====================================================
' === Процедура SelectFirstEmptyInA ===================
' =====================================================
' → Знаходить першу порожню комірку в колонці A, починаючи з A4.
' → Виділяє її та прокручує вікно так, щоб комірка була видима.
' → Використовується для швидкої навігації до наступної доступної позиції.
Sub SelectFirstEmptyInA()
    Dim oDoc As Object, oSheet As Object
    Dim oCell As Object
    Dim iRow As Long

    oDoc = ThisComponent
    oSheet = oDoc.CurrentController.ActiveSheet

    iRow = 3 ' починаючи з A4
    Do While oSheet.getCellByPosition(0, iRow).getString() <> ""
        iRow = iRow + 1
    Loop

    oCell = oSheet.getCellByPosition(0, iRow)
    oDoc.CurrentController.select(oCell)
End Sub

' =====================================================
' === Функція FilterPlace =============================
' =====================================================
' → Повертає діапазон (XCellRange), де місце = nPlace
' → Працює в памʼяті, таблицю не чіпає.
Function FilterPlace(nPlace As Long) As Object
    Dim oRange As Object
    Dim oSheet As Object
    Dim oDoc As Object
    Dim nRowCount As Long
    Dim iRelRow As Long
    Dim oCellR As Object ' Місце
    Dim oCellD As Object ' Чорний список
    Dim iAbsRow As Long

    Dim oResultRanges As Object

    Set oRange = GetRecordsRange()
    nRowCount = oRange.Rows.getCount()

    Set oDoc = ThisComponent
    Set oSheet = oDoc.CurrentController.ActiveSheet

    ' створюємо об'єкт для збирання діапазонів через документ
    Set oResultRanges = oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")

    For iRelRow = 0 To nRowCount - 1
        Set oCellR = oRange.getCellByPosition(17, iRelRow)    ' колонка R
        If Val(oCellR.getValue()) = nPlace Then
            Set oCellD = oRange.getCellByPosition(3, iRelRow) ' колонка D
            If Val(oCellD.getValue()) <> 28 Then
                iAbsRow = iRelRow + 3                         ' бо A4
                oResultRanges.addRangeAddress _
                oSheet.getCellRangeByPosition(0, iAbsRow, 20, iAbsRow).RangeAddress, False
            End If
        End If
    Next iRelRow

    If oResultRanges.getCount() = 0 Then
        Set FilterPlace = Nothing
    Else
        Set FilterPlace = oResultRanges
    End If
End Function

' =====================================================
' === Функція FilterCompetitors =======================
' =====================================================
' → Повертає XSheetCellRanges з усіма рядками, що перетинаються з нашим ренжем.
' → Наш ренж: [dTargetStart, dTargetEnd]
' → Чужі ренжі: [checkIn, checkOut]
Function FilterCompetitors(oFoundRows As Object, oDialog As Object, sAction As String) As Object
    Dim oDoc As Object, oSheet As Object
    Dim oFiltered As Object
    Dim i As Long, nCount As Long
    Dim dTargetStart As Date, dTargetEnd As Date
    Dim oRow As Object
    Dim dCheckIn As Date, dCheckOut As Date
    Dim sCheckIn As String, sCheckOut As String

    ' === Документ та аркуш ===
    Set oDoc = ThisComponent
    Set oSheet = oDoc.CurrentController.ActiveSheet

    ' === Параметри форми ===
    Dim dCurrentDate As Date
    Dim nOffset As Long, nDuration As Long
    Dim nIndexRow As Long

    nIndexRow = ThisComponent.CurrentSelection.RangeAddress.StartRow ' індекс рядка

    If (sAction = ACTION_EDIT Or sAction = ACTION_CREATE) And IsObject(oDialog) Then
        nPlace = Val(oDialog.getControl("PlaceCombo").getText())
        dCurrentDate = CDate(oDialog.getControl("CurrentDateField").getText())
        nOffset = Val(oDialog.getControl("OffsetField").getText())
        nDuration = Val(oDialog.getControl("DurationCombo").getText())
    End If

    If sAction = ACTION_CHECK_ROW Then
        nPlace = oSheet.getCellByPosition(17, nIndexRow).getValue()             ' R
        dCurrentDate = CDate(oSheet.getCellByPosition(0, nIndexRow).getValue()) ' A
        nOffset = oSheet.getCellByPosition(16, nIndexRow).getValue()            ' Q
        nDuration = oSheet.getCellByPosition(19, nIndexRow).getValue()          ' T
    End If

    ' MsgBox "nPlace: " & nPlace & " | dCurrentDate: " & dCurrentDate & " | nOffset: " & nOffset & " | nDuration: " & nDuration

    ' === Наш ренж ===
    dTargetStart = dCurrentDate + nOffset
    dTargetEnd = dTargetStart + nDuration

    ' === Порожній результат ===
    Set oFiltered = oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")

    nCount = oFoundRows.getCount()

    For i = 0 To nCount - 1
        Set oRow = oFoundRows.getByIndex(i)

        sCheckIn = oRow.getCellByPosition(0, 0).String
        sCheckOut = oRow.getCellByPosition(4, 0).String

        If IsDate(sCheckIn) And IsDate(sCheckOut) Then
            dCheckIn = CDate(sCheckIn)
            dCheckOut = CDate(sCheckOut)

            ' === Перетин ренжів ===
            If dTargetStart < dCheckOut And dTargetEnd > dCheckIn Then
                oFiltered.addRangeAddress oRow.RangeAddress, False
            End If
        End If
    Next i

    ' === Результат ===
    If oFiltered.getCount() = 0 Then
        Set FilterCompetitors = Nothing
    Else
        Set FilterCompetitors = oFiltered
    End If
End Function

' =====================================================
' === Функція CheckOccupiedPlace ======================
' =====================================================
' → Перевіряє, чи місце зайняте на вказані дати.
' → Отримує номер місця з форми, знаходить усі рядки по цьому місцю.
' → Якщо редагуємо — виключає поточний рядок з перевірки.
' → Далі фільтрує рядки по перетину дат з нашим ренжем (FilterCompetitors).
' → Повертає True, якщо місце вільне, інакше False та показує зайняті рядки.
Function CheckOccupiedPlace(oDialog As Object, sAction As String) As Boolean
    Dim nPlace As Long
    Dim oFoundRows As Object
    Dim nIndexRow As Long
    Dim dCurrentDate As Date
    Dim nOffset As Long
    Dim dTargetDate As Date

    nIndexRow = ThisComponent.CurrentSelection.RangeAddress.StartRow ' індекс рядка

    If (sAction = ACTION_EDIT Or sAction = ACTION_CREATE) And IsObject(oDialog) Then
        nPlace = Val(oDialog.getControl("PlaceCombo").getText())
        dCurrentDate = CDate(oDialog.getControl("CurrentDateField").getText())
        nOffset = Val(oDialog.getControl("OffsetField").getText())
    End If

    If sAction = ACTION_CHECK_ROW Then
        oSheet = ThisComponent.CurrentController.ActiveSheet

        nPlace = oSheet.getCellByPosition(17, nIndexRow).getValue()             ' R
        dCurrentDate = CDate(oSheet.getCellByPosition(0, nIndexRow).getValue()) ' A
        nOffset = oSheet.getCellByPosition(16, nIndexRow).getValue()            ' Q
    End If

    dTargetDate = dCurrentDate + nOffset
    Set oFoundRows = FilterPlace(nPlace)

    ' MsgBox "StartRow" & oFoundRows.getRangeAddresses()(0).StartRow & Chr(10) & _
           ' "EndRow" & oFoundRows.getRangeAddresses()(0).EndRow

    ' === якщо редагуємо — виключаємо поточний рядок ===
    If sAction = ACTION_EDIT Then
        oFoundRows = ExcludeRow(oFoundRows, nIndexRow)
    End If

    If oFoundRows Is Nothing Then
        CheckOccupiedPlace = True
        Exit Function
    End If

    Set oFoundRows = FilterCompetitors(oFoundRows, oDialog, sAction)

    ' If oFoundRows Is Nothing Then MsgBox "FilterPlace не знайшов жодного рядка", 48, "FilterCompetitors"

    ' MsgBox "StartRow" & oFoundRows.getRangeAddresses()(0).StartRow & Chr(10) & _
           ' "EndRow" & oFoundRows.getRangeAddresses()(0).EndRow, 48, "FilterCompetitors"

    If oFoundRows Is Nothing Then
        CheckOccupiedPlace = True
        Exit Function
    End If

    ShowFields oFoundRows, "Можливі перетени діапазонів на цьому місці"

    ' ShowFields GetOccupiedRows(dtargetDate), "Тестування реньджу"
    If sAction = ACTION_EDIT Or sAction = ACTION_CREATE Then
        MsgDlg "Вільні місця: ", GetVacantPlacesString(GetOccupiedRows(dTargetDate), ", "), False, 65
    End If

    CheckOccupiedPlace = False
End Function

' =====================================================
' === Функція ExcludeRow ==============================
' =====================================================
' → Приймає XSheetCellRanges (`oRanges`) і номер рядка (`nRowToExclude`).
' → Повертає новий XSheetCellRanges без зазначеного рядка.
' → Якщо `oRanges` порожній або після виключення немає жодного рядка — повертає Nothing.
Function ExcludeRow(oRanges As Object, nRowToExclude As Long) As Object
    If oRanges Is Nothing Then
        Set ExcludeRow = Nothing
        Exit Function
    End If

    Dim oDoc As Object, oSheet As Object, oResult As Object, i As Long
    Dim oRow As Object
    Set oDoc = ThisComponent
    Set oSheet = oDoc.CurrentController.ActiveSheet
    Set oResult = oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")

    For i = 0 To oRanges.getCount() - 1
        Set oRow = oRanges.getByIndex(i)
        If oRow.RangeAddress.StartRow <> nRowToExclude Then
            oResult.addRangeAddress oRow.RangeAddress, False
        End If
    Next i

    If oResult.getCount() = 0 Then
        Set ExcludeRow = Nothing
    Else
        Set ExcludeRow = oResult
    End If
End Function

' =====================================================
' === Процедура ShowFields ============================
' =====================================================
' → Виводить дані по кожному рядку діапазону oFoundRange.
' → Виводить A, B, С, E, O, R (заселення, прізвище, ім'я по батькові, виселення, створено, місце)
Sub ShowFields(oFoundRange As Object, sTitle As String)
    Dim nRangeCount As Long
    Dim i As Long
    Dim sOutput As String

    nRangeCount = oFoundRange.getCount()
    sOutput = ""

    For i = 0 To nRangeCount - 1
        Dim oRow As Object
        Dim checkIn As String, lastName As String
        Dim checkOut As String, created As String, place As String, id As String

        Set oRow = oFoundRange.getByIndex(i)

        checkIn = oRow.getCellByPosition(0, 0).String      ' A
        lastName = oRow.getCellByPosition(1, 0).String     ' B
        sFullName = oRow.getCellByPosition(2, 0).String    ' С
        checkOut = oRow.getCellByPosition(4, 0).String     ' E
        Term = oRow.getCellByPosition(19, 0).String        ' O
        place = oRow.getCellByPosition(17, 0).String       ' R
        id = oRow.getCellByPosition(20, 0).String          ' U

        sOutput = sOutput & _
               lastName & " " & sFullName & Chr(10) & _
               "Місце: " & place & "      Тривалість " & Term & " " & DayWord(Term) & Chr(10) & _
               "Заселення      |      " & "Виселення" & Chr(10) & _
               checkIn & "              " & checkOut & Chr(10) & _
               Chr(10) & _
               "id: " & id & Chr(10) & String(50, "-") & Chr(10)
    Next i

    MsgDlg sTitle, sOutput, True
End Sub

Function DayWord(n As Long) As String
    n = Abs(n) Mod 100
    Dim n1 As Long
    n1 = n Mod 10

    Select Case True
        Case n > 10 And n < 20
            DayWord = "днів"
        Case n1 = 1
            DayWord = "день"
        Case n1 >= 2 And n1 <= 4
            DayWord = "дні"
        Case Else
            DayWord = "днів"
    End Select
End Function

' =====================================================
' === Функція ShowPasswordDialog ======================
' =====================================================
' → Виводить діалог для введення пароля.
' → Порівнює введений пароль з очікуваним та повертає True, якщо вони збігаються.
Function ShowPasswordDialog(sExpectedPassword As String) As Boolean

    Dim oDialog      As Object
    Dim oDialogModel As Object
    Dim oField       As Object
    Dim oButton      As Object
    Dim sInput       As String
    Dim bResult      As Boolean

    ' створюємо діалог і модель
    oDialog = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    oDialogModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    oDialog.setModel(oDialogModel)

    With oDialogModel
        .PositionX = 100
        .PositionY = 100
        .Width = 160
        .Height = 60
        .Title = "Введіть пароль"
    End With

    ' поле пароля
    FieldTemplate oDialogModel, "Password", "Пароль:", 30, 15, "", 40, 100

    ' кнопка OK
    AddButton oDialogModel, "OkButton", "OK", 55, 40, 50, 14, 1 ' 1 = OK кнопка

    ' створюємо peer
    oDialog.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)

    oDialog.getControl("PasswordField").Model.EchoChar = Asc("*")

    ' виконуємо діалог
    If oDialog.execute() = 1 Then
        oField = oDialog.getControl("PasswordField")
        sInput = oField.getModel().Text
        bResult = (sInput = sExpectedPassword)
    End If

    oDialog.dispose()

    ShowPasswordDialog = bResult
End Function

' =====================================================
' === Функція GetAfterLastEncashRange =================
' =====================================================
' → Визначає діапазон рядків для інкасації.
' → Шукає останній запис інкасації та порожній рядок у колонці E.
' → Повертає масив [start, end].
' → Якщо немає що інкасувати — повертає [0, 0].
Function GetAfterLastEncashRange() As Variant
    Dim oSheet      As Object
    Dim lStartRow   As Long
    Dim lEndRow     As Long
    Dim lCheckRow   As Long

    oSheet = ThisComponent.Sheets(0)

    lStartRow = 0
    lEndRow = 0

    ' === знаходимо перший порожній рядок у колонці E (вниз від A4) ===
    For lCheckRow = 3 To oSheet.Rows.Count - 1
        If Trim(oSheet.getCellByPosition(4, lCheckRow).String) = "" Then
            lEndRow = lCheckRow - 1
            Exit For
        End If
    Next

    ' якщо порожнього так і не знайшли — беремо останній рядок аркуша
    If lEndRow = 0 Then lEndRow = oSheet.Rows.Count - 1

    ' === знаходимо останню інкасацію від lEndRow вгору ===
    For lCheckRow = lEndRow To 3 Step -1
        If Trim(oSheet.getCellByPosition(4, lCheckRow).String) = ENCASH Then
            lStartRow = lCheckRow + 1
            Exit For
        End If
    Next

    ' якщо інкасацій не знайшли — стартуємо з A4
    If lStartRow = 0 Then lStartRow = 3

    ' перевіряємо коректність діапазону
    If lEndRow < lStartRow Then
        GetAfterLastEncashRange = Array(0, 0) ' немає діапазону
    Else
        GetAfterLastEncashRange = Array(lStartRow, lEndRow)
    End If
End Function

' =====================================================
' === Функція GetOccupiedRows =========================
' =====================================================
' → Повертає SheetCellRanges із зайнятими місцями на вказану дату.
' → Перевіряє діапазон [CheckIn, CheckOut] і виключає ENCASH.
' → Результат: XSheetCellRanges.
Function GetOccupiedRows(targetDate As Date) As Object
    Dim oSheet As Object
    Dim oRanges As Object
    Dim oRange As Object
    Dim nRows As Long
    Dim i As Long
    Dim sCheckOut As String
    Dim dCheckIn As Date, dCheckOut As Date

    oRange = GetRecordsRange()
    oSheet = oRange.Spreadsheet
    nRows = oRange.Rows.Count

    oRanges = ThisComponent.createInstance("com.sun.star.sheet.SheetCellRanges")

    For i = 0 To nRows - 1
        On Error Resume Next
        dCheckIn  = CDate(oSheet.getCellByPosition(0, i + 3).String)
        sCheckOut = oSheet.getCellByPosition(4, i + 3).String
        On Error GoTo 0

        If sCheckOut <> ENCASH Then
            dCheckOut = CDate(sCheckOut)
            If dCheckIn <= targetDate And targetDate <= dCheckOut Then
                oRanges.addRangeAddress _
                    oSheet.getCellRangeByPosition(0, i + 3, 20, i + 3).RangeAddress, False
            End If
        End If
    Next i

    GetOccupiedRows = oRanges
End Function

' =====================================================
' === Функція GetVacantPlacesString ===================
' =====================================================
' → Повертає рядок вільних місць через роздільник (за замовчуванням ";").
' → Порівнює ALL_PLACES та зайняті місця.
' → Результат: рядок вільних місць.
Function GetVacantPlacesString(oOccupiedRows As Object, Optional sSeparator As String) As String
    If IsMissing(sSeparator) Then sSeparator = ";"

    Dim aAllPlaces() As String
    Dim aOccupiedPlaces() As String
    Dim aVacantPlaces() As String
    Dim i As Long

    ' Всі місця
    aAllPlaces = Split(ALL_PLACES, ";")

    ' Зайняті місця
    If oOccupiedRows.getCount() > 0 Then
        ReDim aOccupiedPlaces(oOccupiedRows.getCount() - 1)
        For i = 0 To oOccupiedRows.getCount() - 1
            aOccupiedPlaces(i) = Trim(oOccupiedRows.getByIndex(i).getCellByPosition(17, 0).String)
        Next i
    Else
        aOccupiedPlaces = Array()
    End If

    ' Вільні = A - B
    aVacantPlaces = DiffArrays(aAllPlaces, aOccupiedPlaces)

    If UBound(aVacantPlaces) >= 0 Then
        GetVacantPlacesString = Join(aVacantPlaces, sSeparator & " ")
    Else
        GetVacantPlacesString = ""
    End If
End Function

' =====================================================
' === Процедура SetNextId =============================
' =====================================================
' → Процедура для встановлення наступного id у колонці 20.
' → Якщо попередній рядок — другий, то встановлює id = 1.
' → Інакше бере значення з попереднього рядка і додає 1.
' → Записує нове значення у поточну клітинку (колонка 20).
' → Не повертає значення (процедура).
Sub SetNextId(oSheet As Object, oSel As Object)
    Dim id As Long
    Dim currentId As String
    Dim oCursorAddress As Object

    oCursorAddress = oSel.CellAddress

    ' ==== Перевіряємо, чи вже є id у поточній клітинці ====
    currentId = Trim(oSheet.getCellByPosition(20, oCursorAddress.Row).String)
    If Len(currentId) > 0 And Val(currentId) > 0 Then
        Exit Sub
    End If

    id = 0

    ' ==== Визначаємо попередній id ====
    If oCursorAddress.Row - 1 <> 2 Then
        id = Val(oSheet.getCellByPosition(20, oCursorAddress.Row - 1).String)
    End If

    ' ==== Встановлюємо наступний id ====
    oSheet.getCellByPosition(20, oSel.CellAddress.Row).setValue(id + 1)
End Sub

' =====================================================
' === Функція DiffArrays ==============================
' =====================================================
' → Повертає масив A–B (елементи A, яких немає в B).
' → Використовується для визначення вільних місць.
' → Результат: масив строк.
Function DiffArrays(A As Variant, B As Variant) As Variant
    Dim C() As String
    Dim i As Long, j As Long, n As Long
    Dim isInB As Boolean

    ReDim C(UBound(A))
    n = 0

    For i = 0 To UBound(A)
        isInB = False
        For j = 0 To UBound(B)
            If Trim(A(i)) = Trim(B(j)) Then
                isInB = True
                Exit For
            End If
        Next j
        If Not isInB Then
            C(n) = A(i)
            n = n + 1
        End If
    Next i

    If n > 0 Then
        ReDim Preserve C(n - 1)
    Else
        C = Array()
    End If

    DiffArrays = C
End Function

' =====================================================
' === Процедура ShowArray =============================
' =====================================================
' → Виводить елементи масиву через кому в MsgBox.
' → Якщо масив порожній — повідомляє про це.
Sub ShowArray(arr As Variant)
    Dim s As String
    Dim i As Long

    If IsEmpty(arr) Then
        MsgBox "Масив порожній", 64, "Результат"
        Exit Sub
    End If

    s = ""
    For i = LBound(arr) To UBound(arr)
        s = s & arr(i)
        If i < UBound(arr) Then s = s & ", "
    Next i

    MsgBox s, 64, "Результат"
End Sub
