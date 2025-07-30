REM  *****  BASIC  *****

' Search.bas

' =====================================================
' === Процедура ShowFormByLast ========================
' =====================================================
' → Створює й відображає діалог пошуку зі спадним списком типів полів.
' → Прив’язує кнопку "Пошук" до відповідного обробника.
' → Після закриття діалогу — очищає слухачі та ресурси.
Sub ShowFormByLast()
    Dim oDialog         As Object  
    Dim oButtonSearch   As Object
    Dim oListenerSearch As Object
    Dim sResult         As String
    
    oDialog = CreateDlgSearch()
    
    ' === Кнопка "Пошук" ===
    oButtonSearch    = oDialog.getControl("SearchButton")
    
    ' === Обробник кнопки oButtonCancel ===
    oListenerSearch = CreateUnoListener("SearchButton_", "com.sun.star.awt.XActionListener")
    oButtonSearch.addActionListener(oListenerSearch)
    
    oDialog.execute()   

    ' === Очищення ===
    oButtonSearch.removeActionListener(oListenerSearch)
    oDialog.dispose() 
End Sub

' =====================================================
' === Процедура SearchButton_actionPerformed ==========
' =====================================================
' → Обробляє подію натискання кнопки "Пошук".
' → Запускає пошук за обраним полем і введеним значенням.
' → Виводить відформатований результат у поле ResultEdit.
Sub SearchButton_actionPerformed(oEvent As Object)
    Dim oDialog   As Object
    Dim oResSmpl  As Object
    
    oDialog   = oEvent.Source.getContext()
    oResSmpl  = oDialog.getControl("ResultEdit")
    
    oResSmpl.setText(GetSearchResult(oDialog))         
End Sub

' =====================================================
' === Функція GetSearchResult =========================
' =====================================================
' → Отримує введене значення та тип пошуку з діалогу.
' → Визначає відповідний стовпець та діапазон рядків (за останні 2 місяці).
' → Знаходить збіги й формує результат для виводу.
Function GetSearchResult(oDialog As Object) As String
    Dim sSearch          As String
    Dim sSearchType      As String
    Dim iNumCol          As Integer
    Dim aMonthRange      As Variant : aMonthRange = GetSearchRangeRows()
    Dim aMatchingNumRows As Variant
    
    ' Отримуємо текст для пошуку
    sSearch = Trim(oDialog.getControl("SearchField").getText())

    ' Отримуємо обране значення з комбо-боксу (тип пошуку)
    sSearchType = Trim(oDialog.getControl("TypeCombo").getText())
    
    iNumCol = NumberColumnInit(sSearchType)
    
    aMatchingNumRows = FindMatchingRows(aMonthRange, sSearchType, sSearch)
    
    GetSearchResult = FormatSearchResults(aMatchingNumRows)
End Function

' =====================================================
' === Функція NumberColumnInit ========================
' =====================================================
' → Перетворює назву типу пошуку (з комбо-боксу) на номер відповідного стовпця в таблиці.
' → Повертає: Integer — номер колонки або -1, якщо тип невідомий.
Function NumberColumnInit(sSearchType As String) As Integer
    Select Case sSearchType
        Case "за місцем"               ' Q
            NumberColumnInit = 16
        Case "по прізвищу"             ' B
            NumberColumnInit  = 1
        Case "за iм'ям по-батькові"    ' C
            NumberColumnInit  = 2
        Case "по id"                   ' T
            NumberColumnInit = 19
        Case "за номером телефону"     ' J
            NumberColumnInit  = 9
        Case "по адміністратору"       ' U
            NumberColumnInit = 20
        Case "за датою заселення"      ' A
            NumberColumnInit  = 0
        Case "за датою виселення"      ' E
            NumberColumnInit  = 4
        Case "за датою народження"     ' L
            NumberColumnInit = 11
        Case "за терміном"             ' D
            NumberColumnInit  = 3
        Case "за кодом"                ' S
            NumberColumnInit = 18            
        Case Else
            NumberColumnInit = -1
            Exit Function
    End Select
End Function

' =====================================================
' === Функція GetSearchRangeRows ======================
' =====================================================
' → Обмежує діапазон пошуку лише останніми рядками за N місяці (колонка O — дата створення).
' → Повертає масив [startRow, endRow] для подальшого сканування.
Function GetSearchRangeRows() As Variant
    Dim oSheet       As Object : oSheet  = ThisComponent.Sheets.getByName("Data")
    Dim lastRow      As Long   : lastRow = FindLastRow()
    Dim startRow     As Long
    Dim dCreated     As Date
    Dim i            As Long
    Dim twoMonthsAgo As Date
    Dim sDateStr     As String

    twoMonthsAgo = DateAdd("m", SEARCH_MONTHS_AGO, Now)
    
    For i = lastRow To 3 Step -1
        sDateStr = Trim(oSheet.getCellByPosition(14, i).String) ' колонка O = 14

        If sDateStr <> "" And IsDate(sDateStr) Then
            dCreated = CDate(sDateStr)
            If dCreated < twoMonthsAgo Then
                Exit For
            End If
        End If
    Next i

    startRow = i + 1 ' перший, що ≥ twoMonthsAgo
    If startRow > lastRow Then startRow = lastRow

    GetSearchRangeRows = Array(startRow, lastRow)
End Function

' =====================================================
' === Функція FindMatchingRows ========================
' =====================================================
' → Ітерує заданий діапазон рядків і шукає ті, де значення в колонці відповідає критерію.
' → Повертає масив рядків, які задовольняють умові, або порожній масив.
Function FindMatchingRows(aMonthRange As Variant, _
                          sSearchType As String, _
                          sSearch As String) As Variant
                          
    Dim oSheet     As Object  : oSheet = ThisComponent.Sheets.getByName("Data")
    Dim iStart     As Long    : iStart = aMonthRange(0)
    Dim iEnd       As Long    : iEnd   = aMonthRange(1)
    Dim iCol       As Integer : iCol   = NumberColumnInit(sSearchType)
    Dim aResult()  As Long
    Dim iCount     As Integer : iCount = 0
    Dim i          As Long
    Dim sCellVal   As String
    Dim nCode      As Long    : nCode  = 0
    Dim bEx        As Boolean : bEx    = False
    Dim bCd        As Boolean : bCd    = False 
    
    For i = iStart To iEnd
        sCellVal = Trim(oSheet.getCellByPosition(iCol, i).String)
        nCode    = oSheet.getCellByPosition(18, i).getValue
        
        bEx = InStr(EXCLUDED_CODES, " " & nCode & " ") > 0
        bCd = sSearchType <> "за кодом"
        
        If Not (bEx And bCd) Then
            If CompareCellValue(sSearchType, sCellVal, sSearch) Then
                ReDim Preserve aResult(iCount)
                aResult(iCount) = i
                iCount = iCount + 1
            End If
        End If    
    Next i

    If iCount = 0 Then
        FindMatchingRows = Array() ' порожній
    Else
        FindMatchingRows = aResult
    End If
End Function

' =====================================================
' === Функція CompareCellValue ========================
' =====================================================
' → Порівнює вміст клітинки з шуканим значенням залежно від типу пошуку.
' → Для числових і точних — повна відповідність.
' → Для текстових — пошук входження.
' → Для дат — через NormalizeDate.
Function CompareCellValue(sSearchType As String, sCellVal As String, sSearch As String) As Boolean
    Dim sVal    As String : sVal    = LCase(Trim(sCellVal))
    Dim sNeedle As String : sNeedle = LCase(Trim(sSearch))

    Select Case sSearchType
        Case "по id", "за місцем", "за терміном", "за номером телефону", "за кодом"
            CompareCellValue = (sVal = sNeedle)

        Case "по прізвищу", "за iм'ям по-батькові"
            CompareCellValue = (InStr(sVal, sNeedle) > 0)

        Case "по адміністратору"
            Dim aParts() As String : aParts = Split(sVal, " ")
            If UBound(aParts) >= 0 Then
                CompareCellValue = (InStr(aParts(0), sNeedle) > 0)
            Else
                CompareCellValue = False
            End If

        Case "за датою заселення", "за датою виселення", "за датою народження"           
            CompareCellValue = (NormalizeDate(sVal) = NormalizeDate(sNeedle))

        Case Else
            CompareCellValue = False
    End Select
End Function

' =====================================================
' === Функція NormalizeDate ===========================
' =====================================================
' → Уніфікує формат дати з текстового значення.
' → Замінює роздільники на крапки, розбиває на частини й формує у форматі "DD.MM.YYYY".
' → У разі помилки повертає "00.00.0000".
Function NormalizeDate(sInput As String) As String
    Dim sClean As String
    Dim aParts() As String
    Dim d As Date
    Dim dayPart As Integer, monthPart As Integer, yearPart As Integer

    ' Замінюємо роздільники на крапку
    sClean = Replace(sInput, "/", ".")
    sClean = Replace(sClean, ",", ".")
    
    ' Розбиваємо
    aParts = Split(sClean, ".")

    If UBound(aParts) <> 2 Then
        NormalizeDate = "00.00.0000"
        Exit Function
    End If

    On Error GoTo Fail
    dayPart   = CInt(aParts(0))
    monthPart = CInt(aParts(1))
    yearPart  = CInt(aParts(2))

    d = DateSerial(yearPart, monthPart, dayPart)
    NormalizeDate = Format(d, "DD.MM.YYYY")
    Exit Function

Fail:
    NormalizeDate = "00.00.0000"
End Function

' =====================================================
' === Функція FormatSearchResults =====================
' =====================================================
' → Формує фінальне текстове представлення знайдених записів у вигляді таблиці.
' → Включає поля: id, місце, дати заселення/виселення, термін, ПІБ.
' → Використовує Fr для вирівнювання колонок.
Function FormatSearchResults(aMatchingRows As Variant) As String
    Dim oSheet As Object : oSheet = ThisComponent.Sheets.getByName("Data")
    Dim sOut As String
    Dim i As Integer, row As Long

    ' ───── Шапка таблиці з фіксованою шириною ─────
    sOut = Fr("id", 7) & "| " & _
           Fr("місце", 7) & "| " & _
           Fr("заселення", 11) & "| " & _
           Fr("виселення", 11) & "| " & _
           Fr("термін", 6) & "| " & _
           Fr("прізвище ім'я по батькові", 36) & Chr(10)

    sOut = sOut & String(40, Chr(8212)) & Chr(10) ' горизонтальна лінія

    ' ───── Цикл по знайдених рядках ─────
    For i = 0 To UBound(aMatchingRows)
        row = aMatchingRows(i)

        Dim sId    As String : sId    = oSheet.getCellByPosition(19, row).String
        Dim sPlace As String : sPlace = oSheet.getCellByPosition(16, row).String
        Dim sIn    As String : sIn    = oSheet.getCellByPosition(0, row).String
        Dim sOutD  As String : sOutD  = oSheet.getCellByPosition(4, row).String
        Dim sTerm  As String : sTerm  = oSheet.getCellByPosition(3, row).String
        Dim sFio   As String : sFio   = oSheet.getCellByPosition(1, row).String & " " & _
                                        oSheet.getCellByPosition(2, row).String

        sOut = sOut & Fr(sId, 7)      & "| " & _
                       Fr(sPlace, 7)  & "| " & _
                       Fr(sIn, 12)    & "| " & _
                       Fr(sOutD, 12)  & "| " & _
                       Fr(sTerm, 6)   & "| " & _
                       Fr(sFio, 36)   & Chr(10)
    Next i

    FormatSearchResults = sOut
End Function

' =====================================================
' === Функція CreateDlgSearch =========================
' =====================================================
' → Створює діалог пошуку з полем для введення, комбобоксом для типу та зоною виводу результату.
' → Додає кнопку "Пошук", тло й шаблонні елементи.
' → Повертає готовий об'єкт діалогу.
Function CreateDlgSearch()    
    Dim oDialog      As Object
    Dim oDialogModel As Object
    
    oDialog      = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    oDialogModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
      
    oDialog.setModel(oDialogModel)
    
    ' ==== Параметри діалогу ====    
    With oDialogModel
        .PositionX = 100
        .PositionY = 100
        .Width     = 350
        .Height    = 250
        .Title     = "Пошук"
    End With
    
    Dim gX As Long, gY As Long
    gX = 10 : gY = 15
    
    AddBackground(oDialogModel, BACKGROUND) 
       
    FieldTemplate(oDialogModel,    "Search", "Шукане значення:", gx, gY, "", 70, 100)
    
    ComboBoxTemplate(oDialogModel, "Type",              "Поле:", gx + 115, gY, "за місцем", 50,  100, SEARCH_LIST)
    
    AddEditTemplate(oDialogModel,  "Result", gx, 20 + gY, 330, 180, "Тут буде результат пошуку", True)

    AddButton(oDialogModel,  "SearchButton", "Пошук", 135 + gx, 215 + gY, 60, 14)
    
    oDialog.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)
       
    CreateDlgSearch = oDialog
End Function

