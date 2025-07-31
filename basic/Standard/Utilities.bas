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
' =================== Структура Map ===================
' =====================================================
' ==== Простий аналог Map/Dictionary ==================
' ==== Реалізовано як масив пар Array(Key, Value) =====
' =====================================================

' =====================================================
' === Функція CreateMap ===============================
' =====================================================
' → Створює порожню мапу (масив пар).
' → Кожен елемент: Array(Key As String, Value As String)
' → Повертає Variant-масив.
Function CreateMap() As Variant
    CreateMap = Array()
End Function

' =====================================================
' === Процедура MapPut ================================
' =====================================================
' → Додає або оновлює пару (Key, Value) в мапі.
' → Якщо ключ уже існує — оновлює його значення.
' → Якщо ключ відсутній — додає нову пару.
Sub MapPut(ByRef Map As Variant, ByVal Key As String, ByVal Value As String)
    Dim i As Long
    For i = LBound(Map) To UBound(Map)
        If Map(i)(0) = Key Then
            Map(i)(1) = Value
            Exit Sub
        End If
    Next i

    Dim newEntry(1) As Variant
    newEntry(0) = Key
    newEntry(1) = Value

    If IsEmpty(Map) Then
        Map = Array(newEntry)
    Else
        Map = AppendArray(Map, newEntry)
    End If
End Sub

' =====================================================
' === Функція MapGet ==================================
' =====================================================
' → Повертає значення за ключем.
' → Якщо ключ не знайдено — повертає "" (порожній рядок).
Function MapGet(ByVal Map As Variant, ByVal Key As String) As String
    Dim i As Long
    For i = LBound(Map) To UBound(Map)
        If Map(i)(0) = Key Then
            MapGet = Map(i)(1)
            Exit Function
        End If
    Next i
    MapGet = ""
End Function

' =====================================================
' === Функція MapHasKey ===============================
' =====================================================
' → Перевіряє, чи є ключ у мапі.
' → Повертає True — якщо ключ знайдено.
' → Повертає False — якщо ключ відсутній.
Function MapHasKey(ByVal Map As Variant, ByVal Key As String) As Boolean
    Dim i As Long
    For i = LBound(Map) To UBound(Map)
        If Map(i)(0) = Key Then
            MapHasKey = True
            Exit Function
        End If
    Next i
    MapHasKey = False
End Function

' =====================================================
' === Функція MapGetByIndex ===========================
' =====================================================
' → Повертає пару Array(Key, Value) за індексом.
' → Якщо індекс поза межами — повертає Array("", "").
Function MapGetByIndex(ByVal Map As Variant, ByVal Index As Long) As Variant
    If Index >= LBound(Map) And Index <= UBound(Map) Then
        MapGetByIndex = Map(Index)
    Else
        MapGetByIndex = Array("", "")
    End If
End Function

' =====================================================
' === Функція AppendArray =============================
' =====================================================
' → Допоміжна функція.
' → Додає елемент `item` до кінця масиву `arr`.
Function AppendArray(arr As Variant, item As Variant) As Variant
    Dim l As Long
    l = -1
    On Error Resume Next
    l = UBound(arr) + 1
    On Error GoTo 0

    Dim temp() As Variant
    ReDim temp(0 To l)
    Dim i As Long
    For i = 0 To l - 1
        temp(i) = arr(i)
    Next i
    temp(l) = item
    AppendArray = temp
End Function

' =====================================================
' === Процедура MapClear ==============================
' =====================================================
' → Очищає мапу (скидає до порожнього масиву).
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
    Dim sCode     As String
    Dim sPlace    As String

    sDuration = oDialog.getControl("DurationCombo").getText()
    sCode     = oDialog.getControl("CodeCombo").getText()
    sPlace    = oDialog.getControl("PlaceCombo").getText()

    ' ==== Перетворення у числа ====
    Dim nDuration As Long
    Dim nCode     As Long
    Dim nPlace    As Long

    nDuration = Val(sDuration)
    nCode     = Val(sCode)
    nPlace    = Val(sPlace)

    ' ==== Перевірка на чітність місця ====
    Dim bEven As Boolean
    bEven = (nPlace Mod 2 = 0) ' True — нижнє місце (парне), False — верхнє (непарне)

    ' ==== Відкриваємо лист відповідно до коду ====
    Dim oDoc       As Object
    Dim oSheet     As Object
    Dim sSheetName As String
    
    oDoc       = ThisComponent
    sSheetName = "price" & nCode
    
    ' ==== Якщо листа з таким кодом немає → підставляємо price8 ====
    If Not oDoc.Sheets.hasByName(sSheetName) Then
        sSheetName = "price8"
    End If

    On Error GoTo ErrorHandler
    oSheet = oDoc.Sheets.getByName(sSheetName)
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
    MsgDlg "Помилка", "Не знайдено лист 'price" & nCode & "'. Перевірте правильність коду.", False, 50  
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
    
    Dim oDoc       As Object
    Dim oSheet     As Object
    Dim oCombo     As Object
    Dim iRow       As Long
    Dim aPlaces()  As String    ' масив місць (рядки)
    Dim iCount     As Integer   ' кількість знайдених місць
    Dim nCode      As Long      ' код, вибраний у CodeCombo
    Dim sSheetName As String    ' назва аркуша з цінами
    
    iCount = 0
    ' ==== Отримуємо вибраний код та формуємо ім’я аркуша ====
    nCode      = Val(oDialog.getControl("CodeCombo").getText())
    oDoc       = ThisComponent
    sSheetName = "price" & nCode
    
    ' ==== Якщо аркуш з таким кодом не знайдено — fallback на price8 ====
    If Not oDoc.Sheets.hasByName(sSheetName) Then
        sSheetName = "price8"
    End If

    On Error GoTo ErrorHandler
    oSheet = oDoc.Sheets.getByName(sSheetName)
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
    oCombo = oDialog.getControl("PlaceCombo")
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
    MsgDlg "Помилка", "Не вдалося завантажити список місць з аркуша 'price" & nCode & "'", False, 50      
End Sub

' =====================================================
' === Процедура SelectFirstEmptyInA ===================
' =====================================================
' → Знаходить першу порожню комірку в колонці A, починаючи з A4.
' → Виділяє її та прокручує вікно так, щоб комірка була видима.
' → Використовується для швидкої навігації до наступної доступної позиції.
Sub SelectFirstEmptyInA()
    Dim oDoc   As Object
    Dim oSheet As Object
    Dim oCell  As Object
    Dim iRow   As Long

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
Function FilterPlace(nPlace  As Long, _
                     sAction As String) As Object
    Dim oRange    As Object
    Dim oSheet    As Object
    Dim oDoc      As Object
    Dim nRowCount As Long
    Dim iRelRow   As Long
    Dim oCellR    As Object ' Місце
    Dim oCellD    As Object ' Чорний список
    Dim iAbsRow   As Long

    Dim oResultRanges As Object

    oRange    = GetRecordsRange()
    nRowCount = oRange.Rows.getCount()
    oDoc      = ThisComponent
    oSheet    = oDoc.CurrentController.ActiveSheet

    ' створюємо об'єкт для збирання діапазонів через документ
    oResultRanges = oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")
       
    For iRelRow = 0 To nRowCount - 1
    
        oCellR = oRange.getCellByPosition(16, iRelRow)     ' колонка Q місце
        
        If Val(oCellR.getValue()) = nPlace Then
        
            oCellD = oRange.getCellByPosition(18, iRelRow) ' колонка S код
            
            If InStr(EXCLUDED_CODES, " " & Val(oCellD.getValue()) & " ") = 0 Then
                iAbsRow = iRelRow + 3                          ' бо A4
                oResultRanges.addRangeAddress _
                oSheet.getCellRangeByPosition(0, iAbsRow, 19, iAbsRow).RangeAddress, False
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
Function FilterCompetitors(oFoundRows As Object, _
                           dTargetStart As Date, _
                           dTargetEnd As Date, _
                           sAction As String) As Object

    Dim oDoc       As Object
    Dim oSheet     As Object
    Dim oFiltered  As Object
    Dim i          As Long
    Dim nCount     As Long
    Dim oRow       As Object
    Dim dCheckIn   As Date
    Dim dCheckOut  As Date
    Dim sCheckIn   As String
    Dim sCheckOut  As String

    ' === Документ та аркуш ===
    oDoc = ThisComponent
    oSheet = oDoc.CurrentController.ActiveSheet
    
    ' === Порожній результат ===
    oFiltered = oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")

    nCount = oFoundRows.getCount()

    For i = 0 To nCount - 1
        Set oRow = oFoundRows.getByIndex(i)

        sCheckIn  = oRow.getCellByPosition(0, 0).String
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
Function CheckOccupiedPlace(oDialog As Object, _
                            sAction As String) As Boolean
    Dim nPlace        As Long 
    Dim oFoundRows    As Object
    Dim nIndexRow     As Long
    Dim oCheckInDate  As Object
    Dim cI            As Object
    Dim dCheckInDate  As Date
    Dim dCheckOutDate As Date    
    Dim dTargetDate   As Date
    
    nIndexRow = ThisComponent.CurrentSelection.RangeAddress.StartRow ' індекс рядка
    
    If (sAction = ACTION_EDIT Or sAction = ACTION_CREATE) And IsObject(oDialog) Then
        nPlace        = Val(oDialog.getControl("PlaceCombo").getText())
        oCheckInDate  = oDialog.getControl("CheckInDate")
        cI            = oCheckInDate.Date        
        ' Якщо застосували у календар кнопку "Немає" порожнє поле
        cI = IniUtilDate(cI)
        
        dCheckInDate  = DateSerial(cI.Year, cI.Month, cI.Day)
        dCheckOutDate = CDate(oDialog.getControl("CheckOutField").getText())
        ' MsgBox "dCheckInDate: " & dCheckInDate & "; dCheckOutDate: " & dCheckOutDate
    End If
    
    If sAction = ACTION_CHECK_ROW or sAction = ACTION_CHECK_CANCEL Then
        oSheet = ThisComponent.CurrentController.ActiveSheet      
 
        nPlace = oSheet.getCellByPosition(16, nIndexRow).getValue()              ' Q
        dCheckInDate  = CDate(oSheet.getCellByPosition(0, nIndexRow).getValue()) ' A
        dCheckOutDate = CDate(oSheet.getCellByPosition(4, nIndexRow).getValue()) ' E
    End If
     
    oFoundRows = FilterPlace(nPlace, sAction)
    
    ' === якщо редагуємо — виключаємо поточний рядок ===
    If sAction = ACTION_EDIT or sAction = ACTION_CHECK_CANCEL Then
        oFoundRows = ExcludeRow(oFoundRows, nIndexRow)
    End If
    
    If oFoundRows Is Nothing Then
        CheckOccupiedPlace = True
        Exit Function
    End If

    oFoundRows = FilterCompetitors(oFoundRows, dCheckInDate, dCheckOutDate, sAction) 
        
    If oFoundRows Is Nothing Then
        CheckOccupiedPlace = True
        Exit Function
    End If

    ShowFields oFoundRows, "Можливі перетени діапазонів на цьому місці"
    
     
    If sAction = ACTION_EDIT Or sAction = ACTION_CREATE Then
        Dim aOccupied As Variant : aOccupied = GetOccupiedRows(dCheckInDate)
    
        MsgDlg "Місця " & dCheckInDate, _
            "Вільні місця на " & dCheckInDate & ":" & Chr(10) & _
            GetPlacesString(aOccupied(0), ", ", "vacant") & Chr(10) & Chr(10) & _
            IsNowMessage(dCheckInDate) & Chr(10) & _
            Chr(10) & _
            GetPlacesString(aOccupied(1), ", ", "occupied"), _
            True, 110, 300
    End If
        
    CheckOccupiedPlace = False
End Function

Function IsNowMessage(targetDate As Date) As String
    If targetDate = Date Then
        IsNowMessage = DeadlineMessage()
    Else
        IsNowMessage = "Закінчився термін у цей день (" & Format(targetDate, "DD.MM.YYYY") & "):"
    End If
End Function

Function DeadlineMessage() As String
    If TimeValue(Now) < TimeValue("12:00") Then
        DeadlineMessage = "Сьогодні закінчується термін у 12:00:"
    Else
        DeadlineMessage = "Сьогодні закінчився термін у 12:00:"
    End If
End Function

' =====================================================
' === Функція ExcludeRow ==============================
' =====================================================
' → Приймає XSheetCellRanges (`oRanges`) і номер рядка (`nRowToExclude`).
' → Повертає новий XSheetCellRanges без зазначеного рядка.
' → Якщо `oRanges` порожній або після виключення немає жодного рядка — повертає Nothing.
Function ExcludeRow(oRanges       As Object, _
                    nRowToExclude As Long) As Object

    If oRanges Is Nothing Then
        ExcludeRow = Nothing
        Exit Function
    End If
    
    Dim oDoc    As Object
    Dim oSheet  As Object
    Dim oResult As Object
    Dim i       As Long
    Dim oRow    As Object
    
    oDoc = ThisComponent
    oSheet = oDoc.CurrentController.ActiveSheet
    oResult = oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")   

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
Sub ShowFields(oFoundRange As Object, _
               sTitle      As String)
    
    Dim nRangeCount As Long
    Dim i           As Long
    Dim sOutput     As String

    nRangeCount = oFoundRange.getCount()
    sOutput = ""

    For i = 0 To nRangeCount - 1
        Dim oRow     As Object
        Dim checkIn  As String
        Dim lastName As String
        Dim checkOut As String
        Dim created  As String
        Dim place    As String
        Dim id       As String

        Set oRow = oFoundRange.getByIndex(i)

        lastName  = oRow.getCellByPosition(1, 0).String     ' B
        sFullName = oRow.getCellByPosition(2, 0).String     ' С
        place     = oRow.getCellByPosition(16, 0).String    ' Q
        Term      = oRow.getCellByPosition(3, 0).String     ' D
        checkIn   = oRow.getCellByPosition(0, 0).String     ' A
        checkOut  = oRow.getCellByPosition(4, 0).String     ' E
        id = oRow.getCellByPosition(19, 0).String           ' T

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

' =====================================================
' === Функція DayWord =================================
' =====================================================
' → Повертає правильну форму слова "день" українською залежно від числа.
' → Використовується для відображення тривалості (1 день, 2 дні, 5 днів тощо).
' → Враховує особливості української граматики (числівники).
Function DayWord(n As Long) As String
    Dim n1 As Long
    
    n = Abs(n) Mod 100
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
' === Функція GetOccupiedRows =========================
' =====================================================
' → Повертає SheetCellRanges із зайнятими місцями на вказану дату.
' → Перевіряє діапазон [CheckIn, CheckOut] і виключає ENCASH.
' → Результат: XSheetCellRanges.
Function GetOccupiedRows(targetDate As Date) As Variant
    Dim oSheet        As Object
    Dim oFreePlaces   As Object
    Dim oEvictionList As Object    
    Dim oRange        As Object
    Dim nRows         As Long
    Dim i             As Long
    Dim sCheckOut     As String
    Dim dCheckIn      As Date
    Dim dCheckOut     As Date
    Dim sCode         As String
    
    oRange        = GetRecordsRange()
    oSheet        = oRange.Spreadsheet
    nRows         = oRange.Rows.Count
    oFreePlaces   = ThisComponent.createInstance("com.sun.star.sheet.SheetCellRanges")
    oEvictionList = ThisComponent.createInstance("com.sun.star.sheet.SheetCellRanges")

    For i = 0 To nRows - 1
        On Error Resume Next
        dCheckIn  = CDate(oSheet.getCellByPosition(0, i + 3).String)
        sCheckOut = oSheet.getCellByPosition(4, i + 3).String
        sCode     = Trim(oSheet.getCellByPosition(18, i + 3).String)       
        On Error GoTo 0
           
        If InStr(EXCLUDED_CODES, " " & sCode & " ") = 0 Then
            dCheckOut = CDate(sCheckOut)
               
            If dCheckIn <= targetDate And targetDate <= dCheckOut Then            
                ' TestOccupiedRows(dCheckIn, targetDate, dCheckOut)                
                oFreePlaces.addRangeAddress _
                    oSheet.getCellRangeByPosition(0, i + 3, 20, i + 3).RangeAddress, False
            End If
            
            If targetDate = dCheckOut Then            
                
                oEvictionList.addRangeAddress _
                    oSheet.getCellRangeByPosition(0, i + 3, 20, i + 3).RangeAddress, False
            End If
        End If
    Next i

    GetOccupiedRows = Array(oFreePlaces, oEvictionList)
End Function

Sub TestOccupiedRows(dCheckIn As Date, targetDate As Date, dCheckOut As Date)
    MsgDlg "Test", "dCheckIn(" & dCheckIn & ") <= " & "targetDate(" & targetDate & _
    ") < dCheckOut(" & dCheckOut & ")", False, 50, 260 
End Sub

' === Функція GetPlacesString ===========================
' =====================================================
' → Формує текстове представлення списку місць (всі, зайняті або вільні).
' → Параметри:
'     oOccupiedRows — набір рядків із зайнятими місцями;
'     sSeparator — роздільник між місцями (за замовчуванням: ";");
'     sMode — режим: "vacant" (за замовчуванням), "all", "occupied".
' → Повертає:
'     - рядок місць (для режимів "vacant" або "all");
'     - або табличку з інфо про кожне зайняте місце ("occupied").
Function GetPlacesString(oOccupiedRows As Object, _
                         Optional sSeparator As String, _
                         Optional sMode As String) As String

    If IsMissing(sSeparator) Then sSeparator = ";"
    If IsMissing(sMode) Then sMode = "vacant"  ' за замовчуванням — вільні

    Dim aAllPlaces()      As String
    Dim aOccupiedPlaces() As String
    Dim aVacantPlaces()   As String
    Dim i                 As Long

    aAllPlaces = Split(ALL_PLACES, ";")

    ' Зайняті місця
    If oOccupiedRows.getCount() > 0 Then
        ReDim aOccupiedPlaces(oOccupiedRows.getCount() - 1)
        For i = 0 To oOccupiedRows.getCount() - 1
            aOccupiedPlaces(i) = Trim(oOccupiedRows.getByIndex(i).getCellByPosition(16, 0).String)
        Next i
    Else
        aOccupiedPlaces = Array()
    End If

    Select Case LCase(sMode)
        Case "all"
        
            GetPlacesString = Join(aAllPlaces, sSeparator & " ")
            
        Case "vacant"
            aVacantPlaces = DiffArrays(aAllPlaces, aOccupiedPlaces)
            If UBound(aVacantPlaces) >= 0 Then
                GetPlacesString = Join(aVacantPlaces, sSeparator & " ")
            Else
                GetPlacesString = ""
            End If        
            
        Case "occupied"
                           
                Dim sResult As String
                sResult = Fr("заселення", 10) & "| " & _
                          Fr("виселення", 10) & "| " & _
                          Fr("місце", 6)      & "| " & _
                          Fr("Id", 7)         & "| " & _
                          Fr("прізвище, ім'я по батькові", 36) & Chr(10) & _
                          String(40, Chr(8212)) & Chr(10)
                          

                Dim oRange As Object
                Dim oAddress As Object
                Dim nRow As Long

                For i = 0 To oOccupiedRows.getCount() - 1
                oRange   = oOccupiedRows.getByIndex(i)
                oAddress = oRange.RangeAddress
                nRow     = oAddress.StartRow

                With oRange.Spreadsheet
                    Dim sCheckIn  As String : sCheckIn  = .getCellByPosition(0, nRow).String  ' A
                    Dim sCheckOut As String : sCheckOut = .getCellByPosition(4, nRow).String  ' E
                    Dim sPlace    As String : sPlace    = .getCellByPosition(16, nRow).String ' Q
                    Dim sId       As String : sId       = .getCellByPosition(19, nRow).String ' T
                    Dim sLast     As String : sLast     = .getCellByPosition(1, nRow).String  ' B
                    Dim sPatr     As String : sPatr     = .getCellByPosition(2, nRow).String  ' C
                End With

                sResult = sResult & _
                    Fr(sCheckIn, 11)      & "| " & _
                    Fr(sCheckOut, 11)     & " | " & _
                    Fr(sPlace, 6)         & "| " & _
                    Fr(sId, 6)            & " | " & _
                    Fr(sLast & " " & sPatr, 36) & Chr(10)    
                Next i

                GetPlacesString = sResult
            
        Case Else
            GetPlacesString = "" ' або MsgBox "невідомий режим"
    End Select
End Function

' === Функція Fr ========================================
' =====================================================
' → Форматує рядок до фіксованої ширини, додаючи заповнення.
' → Параметри:
'     s — рядок для форматування;
'     w — цільова ширина;
'     d — символ заповнення (за замовчуванням: вузький пробіл).
' → Повертає: вирівняний текст ширини `w`.
Function Fr(s, w, Optional d As Variant) As string
    If IsMissing(d) Then d = Chr(8194)
    If Len(s) > w Then w = Len(s)
    Fr = s & String(w - Len(s), d)
End Function
' =====================================================
' === Процедура SetNextId =============================
' =====================================================
' → Процедура для встановлення наступного id у колонці 20.
' → Якщо попередній рядок — другий, то встановлює id = 1.
' → Інакше бере значення з попереднього рядка і додає 1.
' → Записує нове значення у поточну клітинку (колонка 20).
' → Не повертає значення (процедура).
Sub SetNextId(oSheet As Object, _
              oSel   As Object)
              
    Dim id             As Long
    Dim currentId      As String
    Dim oCursorAddress As Object

    oCursorAddress = oSel.CellAddress

    ' ==== Перевіряємо, чи вже є id у поточній клітинці ====
    currentId = Trim(oSheet.getCellByPosition(19, oCursorAddress.Row).String)
    If Len(currentId) > 0 And Val(currentId) > 0 Then
        Exit Sub
    End If

    id = 0

    ' ==== Визначаємо попередній id ====
    If oCursorAddress.Row - 1 <> 2 Then
        id = Val(oSheet.getCellByPosition(19, oCursorAddress.Row - 1).String)
    End If

    ' ==== Встановлюємо наступний id ====
    oSheet.getCellByPosition(19, oSel.CellAddress.Row).setValue(id + 1)
End Sub

' =====================================================
' === Функція DiffArrays ==============================
' =====================================================
' → Повертає масив A–B (елементи A, яких немає в B).
' → Використовується для визначення вільних місць.
' → Результат: масив строк.
Function DiffArrays(A As Variant, _
                    B As Variant) As Variant
    Dim C()   As String
    Dim i     As Long
    Dim j     As Long
    Dim n     As Long
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
 
    On Error GoTo EmptyArray
    If UBound(arr) < LBound(arr) Then GoTo EmptyArray
    
    s = ""
    For i = LBound(arr) To UBound(arr)
        s = s & arr(i)
        If i < UBound(arr) Then s = s & ", "
    Next i
    MsgDlg "Результат", s, False, 50, 140
    
    Exit Sub

EmptyArray:
    MsgDlg "Результат", String(27, " ") & "Масив порожній", False, 50, 140
End Sub

' =====================================================
' === Процедура CalculationTerm =======================
' =====================================================
' → Обчислює дату виселення на основі дати заселення та тривалості.
' → Записує результат у відповідне поле діалогу.
' → Використовується для автоматичного розрахунку кінцевої дати проживання.
Sub CalculationTerm(oDialog As Object)
    Dim oDuration As Long
    Dim oCheckIn  As Object
    Dim oCheckOut As Object
    Dim dStart    As Date
    Dim dEnd      As Date
    Dim cI        As Object
    
    oDuration   = CLng(oDialog.getControl("DurationCombo").Text) 
    oCheckIn    = oDialog.getControl("CheckInDate")
    oCheckOut   = oDialog.getControl("CheckOutField")
    cI          = oCheckIn.Date
    
    ' Якщо застосували у календар кнопку "Немає" порожнє поле
    cI = IniUtilDate(cI)
    
    ' стартова дата
    dStart = DateSerial(cI.Year, cI.Month, cI.Day)
    ' додаємо тривалість
    dEnd   = DateAdd("d", oDuration, dStart)
    ' записуємо
    oCheckOut.Text = dEnd
End Sub

' =====================================================
' === Функція IniUtilDate =============================
' =====================================================
' → Перевіряє, чи дата порожня (рік=0), і якщо так — підставляє сьогоднішню.
' → Повертає структуру дати з валідними значеннями.
' → Використовується для уникнення нульових дат при ініціалізації.
Function IniUtilDate(cI As Object)
    If  cI.Year = 0 Then    
        cI.Year  = Year(Now)
        cI.Month = Month(Now)
        cI.Day   = Day(Now)         
    End If
    IniUtilDate = cI
End Function

' =====================================================
' === Процедура FinCommentSwitch ======================
' =====================================================
' → Відображає або ховає поле коментаря залежно від того, чи введені витрати/надходження.
' → Використовується для зручності заповнення фінансових даних.
' → Якщо обидва значення = 0 — поле й мітка ховаються.
Sub FinCommentSwitch(oDialog As Object)
    Dim oExpenseControl       As Object
    Dim oIncomeControl        As Object
    Dim oCommentControl       As Object
    Dim oCommentLabelControl  As Object
    Dim dExpense              As Double
    Dim dIncome               As Double
    Dim bVisible              As Boolean
    
    oExpenseControl      = oDialog.getControl("ExpenseField")
    oIncomeControl       = oDialog.getControl("IncomeField")
    oCommentControl      = oDialog.getControl("CommentField")
    oCommentLabelControl = oDialog.getControl("CommentLabel")
    dExpense             = Val(oExpenseControl.getText())
    dIncome              = Val(oIncomeControl.getText())

    bVisible = (dExpense <> 0 Or dIncome <> 0)

    oCommentControl.Visible      = bVisible
    oCommentLabelControl.Visible = bVisible
End Sub

' =====================================================
' === Процедура OffsetCommentSwitch ===================
' =====================================================
' → Вмикає або вимикає відображення поля та мітки причини зміщення дати.
' → Перевіряє, чи вибрана дата заселення відрізняється від сьогоднішньої.
' → Якщо дата порожня — підставляє поточну дату та показує повідомлення.
' → Використовується для контролю коректності дати та пояснення змін.
Sub OffsetCommentSwitch(oDialog    As Object, _
                        bInit      As Boolean, _
               Optional sAction    As String)
               
    If IsMissing(sAction) Then sAction = ""
          
    Dim oReasonLabelControl   As Object    
    Dim oReasonFieldControl   As Object   
    Dim oCheckInDateControl   As Object          
    Dim oTargetDateControl    As Object        
    Dim cI                    As Object           
    Dim sTarget               As String
    Dim dTarget               As Date      
    Dim bVisible              As Boolean
       
    oReasonLabelControl   = oDialog.getControl("ReasonLabel")
    oReasonFieldControl   = oDialog.getControl("ReasonField")   
    oCheckInDateControl   = oDialog.getControl("CheckInDate")
    oTargetDateControl    = oDialog.getControl("CurrentDateLabel")    
    cI                    = oCheckInDateControl.Date       
    ' Якщо застосували у календар кнопку "Немає" порожнє поле
    cI                    = IniUtilDate(cI)
    sTarget               = oTargetDateControl.Text
    dTarget               = UnformattingDate(sTarget)
    
    ' Перевіряємо, що поле заповнене
    If Trim(oCheckInDateControl.Text) = "" And Not bInit Then
        MsgDlg "Помилка", "Дата не вибрана.", False, 50, 120
        
        ' вставляємо значення
        If sAction = ACTION_EDIT Then            
            Dim oSheet   As Object
            Dim oSel     As Object
            Dim row      As Long
            Dim sCellValue As String
            
            oSheet     = ThisComponent.Sheets(0)
            oSel       = ThisComponent.CurrentSelection
            row        = oSel.RangeAddress.StartRow
            sCellValue = oSheet.getCellByPosition(0,  row).String
                                                      
            oCheckInDateControl.Text = Format(CDate(sCellValue), "DD.MM.YYYY")
        Else
            oCheckInDateControl.Text = Format(Now, "DD.MM.YYYY")
        End If              
        Exit Sub    
    End If
    
    bVisible = (DateSerial(cI.Year, cI.Month, cI.Day) <> dTarget)
    
    ' Логіка відображення або приховування
    oReasonLabelControl.Visible = bVisible
    oReasonFieldControl.Visible = bVisible
End Sub

' =====================================================
' === Функція UnformattingDate ========================
' =====================================================
' → Виділяє тільки дату з рядка, що містить дату й час.
' → Відкидає час та перенос рядків, залишаючи лише частину з датою.
' → Якщо рядок не містить коректної дати — повертає 0 (Null Date).
Function UnformattingDate(sTarget As String, Optional bKeepTime As Variant) As Date
    If IsMissing(bKeepTime) Then bKeepTime = False
    
    If Not bKeepTime then
        If InStr(sTarget, Chr(10)) > 0 Then
            sTarget = Trim(Left(sTarget, InStr(sTarget, Chr(10)) - 1))
        ElseIf InStr(sTarget, " ") > 0 Then
            sTarget = Trim(Left(sTarget, InStr(sTarget, " ") - 1))
        End If

        If IsDate(sTarget) Then
            UnformattingDate = CDate(sTarget)
        Else
            UnformattingDate = 0
        End If
    Else
        ' замінити перенос на пробіл і обрізати
        sTarget = Replace(sTarget, Chr(10), " ")
        sTarget = Trim(sTarget)
    End If
    
    
    If IsDate(sTarget) Then
        UnformattingDate = CDate(sTarget)
    Else
        UnformattingDate = 0
    End If  
End Function

' === Процедура PersonalDataById ========================
' =====================================================
' → Шукає в таблиці користувача за введеним ID і заповнює відповідні поля діалогу.
' → Виводить помилки у разі порожнього ID, відсутності ID в таблиці або некоректного коду.
' → Повертає: нічого (Sub), працює через побічні ефекти (UI).
Sub PersonalDataById(oEvent As Object)
    Dim oDialog     As Object
    Dim oSheet      As Object
    Dim oSel        As Object
    Dim sId         As String
    Dim i           As Long    : i = 3
    Dim nRows       As Long
    Dim found       As Boolean : found = False
    Dim lValueS     As Long
    Dim sAddTitle   As String
    Dim sHalfMsg    As String
    Dim aName       As Variant
    Dim sFirst      As String  : sFirst = ""
    Dim sPatr       As String  : sPatr  = ""

    oDialog   = oEvent.Source.getContext()
    oSheet    = ThisComponent.Sheets(0)
    nRows     = oSheet.getRows().getCount()      
    sId       = Trim(oDialog.getControl("IdField").Text)
    sAddTitle = "ID " & sId
    
    sHalfMsg  = " він не може бути використаний як джерело персональних даних."
    
    If sId = "" Then
        MsgDlg "Помилка", "ID порожній", False, 50, 120
        Exit Sub
    End If

    ' шукаємо рядок   
    Do While Trim(oSheet.getCellByPosition(19, i).String) <> ""
        If Trim(oSheet.getCellByPosition(19, i).String) = sId Then
            found = True
            Exit Do
        End If
        i = i + 1
    Loop
    
    lValueS   = Val(oSheet.getCellByPosition(18, i).String)
    
    If Not found Then
        MsgDlg "Помилка", "ID " & sId & " не знайдено", False, 50, 120
        Exit Sub
    End If 
    
    If Not CheckCode(sAddTitle, sHalfMsg, lValueS)  Then Exit Sub
    
    aName = Split(oSheet.getCellByPosition(2, i).String, " ")
    If UBound(aName) >= 0 Then sFirst = aName(0)
    If UBound(aName) >= 1 Then sPatr  = aName(1)

    ' заповнюємо контроли
    oDialog.getControl("LastNameField").Text    = oSheet.getCellByPosition(1, i).String
    oDialog.getControl("FirstNameField").Text   = sFirst
    oDialog.getControl("PatronymicField").Text  = sPatr
    oDialog.getControl("PhoneField").Text       = oSheet.getCellByPosition(9, i).String
    oDialog.getControl("PassportField").Text    = oSheet.getCellByPosition(10, i).String
    oDialog.getControl("BirthDate").Text        = oSheet.getCellByPosition(11, i).String

    MsgDlg "Успіх", "Дані заповнено з ID " & sId, False, 50, 120 
End Sub

' === Функція CheckCode =================================
' =====================================================
' → Перевіряє статус коду події у рядку таблиці та блокує обробку, якщо він заборонений.
' → Повертає False — якщо код дорівнює одному з небажаних значень: 0, 7, 28, 30, 70.
' → Виводить відповідне повідомлення з уточненням причини (чорний список, баланс, інкасація тощо).
' → Повертає True — якщо код допустимий для подальшої обробки.
Function CheckCode(sAddTitle As String, sHalfMsg As String, lValueS As Long) As Boolean
    CheckCode = True
    
    Select Case lValueS
        Case 28
            MsgDlg "Відмова: " & sAddTitle, _
                   "Цей рядок належить до ""чорного списку""," & sHalfMsg, _
                   False, 60
            CheckCode = False        
            Exit Function
            
        Case 30
            MsgDlg "Відмова: " & sAddTitle, _
                   "Цей рядок належить до ""балансу""," & sHalfMsg, _
                   False, 60
            CheckCode = False       
            Exit Function
            
        Case 7
            MsgDlg "Відмова: " & sAddTitle, _
                   "Цей рядок — ""інкасація""," & sHalfMsg, _
                   False, 60
            CheckCode = False       
            Exit Function
            
        Case 70
            MsgDlg "Відмова: " & sAddTitle, _
                   "Цей рядок зі статусом ""інше""," & sHalfMsg, _
                   False, 60
            CheckCode = False       
            Exit Function
        Case 0
            MsgDlg "Відмова: " & sAddTitle, _
                   "Цей рядок зі статусом ""Пусто""," & sHalfMsg, _
                   False, 60
            CheckCode = False       
            Exit Function    
    End Select
End Function 

' =====================================================
' === Функція Obfuscate ================================
' =====================================================
' → Виконує просте XOR-шифрування рядка з ключем &HAA.
' → Кожен символ перетворюється в 2-символьне hex-представлення.
' → Повертає шифрований рядок у шістнадцятковому форматі.
Function Obfuscate(s As String) As String
    Dim i As Integer, res As String
    For i = 1 To Len(s)
        res = res & Right("00" & Hex(Asc(Mid(s, i, 1)) Xor &HAA), 2)
    Next i
    Obfuscate = res
End Function

' =====================================================
' === Функція Deobfuscate ==============================
' =====================================================
' → Дешифрує hex-рядок, отриманий через Obfuscate.
' → Кожні 2 hex-символи конвертує назад у символ із XOR-дешифрацією (&HAA).
' → Повертає оригінальний текст.
Function Deobfuscate(hexString As String) As String
    Dim i As Integer, s As String
    For i = 1 To Len(hexString) Step 2
        s = s & Chr(CInt("&H" & Mid(hexString, i, 2)) Xor &HAA)
    Next i
    Deobfuscate = s
End Function

' Sub tObf ()
  'MsgDlg "Test Pass", Obfuscate(), False, 50 
  'MsgDlg "Test Pass", Deobfuscate(), False, 50
' End Sub

' =====================================================
' === Функція IsCancelCode =============================
' =====================================================
' → Перевіряє, чи є код скасуванням (тобто один із: 20, 21, 22, 23).
' → Повертає True — якщо код входить до списку, інакше False.
' → Зручний спосіб перевірки без Select Case або масиву.
Function IsCancelCode(nCode As Long) As Boolean
    IsCancelCode = InStr(" 20 21 22 23 ", " " & nCode & " ") > 0
End Function
