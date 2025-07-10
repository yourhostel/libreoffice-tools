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
    oDoc = ThisComponent

    On Error GoTo ErrorHandler
    Set oSheet = oDoc.Sheets.getByName("price" & nCode)
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
Sub UpdatePlaceCombo(oDialog As Object)
    Dim oDoc As Object, oSheet As Object
    Dim oCombo As Object
    Dim iRow As Long
    Dim aPlaces() As String
    Dim iCount As Integer
    Dim nCode As Long

    iCount = 0

    nCode = Val(oDialog.getControl("CodeCombo").getText())
    oDoc = ThisComponent

    On Error GoTo ErrorHandler
    Set oSheet = oDoc.Sheets.getByName("price" & nCode)

    For iRow = 1 To MAX_SEARCH_RANGE_IN_PRICE
        Dim nPlace As Long
        nPlace = oSheet.getCellByPosition(6, iRow).getValue()

        If nPlace <> 0 Then
            ReDim Preserve aPlaces(iCount)
            aPlaces(iCount) = CStr(nPlace)
            iCount = iCount + 1
        End If
    Next iRow

    Set oCombo = oDialog.getControl("PlaceCombo")
    oCombo.Model.StringItemList = aPlaces

    If iCount > 0 Then
    	oCombo.Model.Text = aPlaces(0)
	End If

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

    iRow = 3 ' начиная с A4
    Do While oSheet.getCellByPosition(0, iRow).getString() <> ""
        iRow = iRow + 1
    Loop

    oCell = oSheet.getCellByPosition(0, iRow)
    oDoc.CurrentController.select(oCell)
End Sub

' =====================================================
' === Функція IsPlaceOccupiedToday ====================
' =====================================================
' → Перевіряє, чи зайняте вказане місце на певний день.
' → Враховує дату заселення, дату виселення і виключені коди.
' → Повертає масив:
'   (0) — True/False (зайняте чи ні)
'   (1) — масив усіх знайдених місць на цей день
Function IsPlaceOccupiedToday(nPlace As Long, nOffset As Long) As Variant
    Dim oDoc As Object, oSheet As Object

    Dim iRow As Long
        iRow = 3

    Dim dToday As Double
        dToday = Int(Now())

    Dim dTargetDay As Double
        dTargetDay = dToday + nOffset
    Dim foundPlaces() As Long
    Dim foundCount As Long
    foundCount = 0

    Set oDoc = ThisComponent
    Set oSheet = oDoc.Sheets.getByName("data")

    Dim excludedCodes() As Variant
    excludedCodes = Split(EXCLUDED_CODES, ";")

    Dim isOccupied As Boolean
    isOccupied = False

    Do While oSheet.getCellByPosition(0, iRow).getValue() <> 0
        Dim checkIn As Double
        Dim checkOut As Double
        Dim code As Long
        Dim place As Long

        checkIn = oSheet.getCellByPosition(0, iRow).getValue()
        checkOut = oSheet.getCellByPosition(4, iRow).getValue()
        code = oSheet.getCellByPosition(3, iRow).getValue()
        place = oSheet.getCellByPosition(17, iRow).getValue()

        sFullName = oSheet.getCellByPosition(1, iRow).string
        sName = oSheet.getCellByPosition(2, iRow).string

        Dim isDateOk As Boolean
        isDateOk = (checkIn <= dTargetDay) And (checkOut >= dTargetDay)

        Dim isExcluded As Boolean
        isExcluded = False

        ' DebugGun(isDateOk, Array(iRow, checkIn, checkOut, code, place, dTargetDay, sFullName, sName))

        Dim i As Integer
        For i = LBound(excludedCodes) To UBound(excludedCodes)
            If code = CLng(excludedCodes(i)) Then
                isExcluded = True
                Exit For
            End If
        Next

        If isDateOk And Not isExcluded Then
            ' Заносимо місце у масив знайдених
            ReDim Preserve foundPlaces(foundCount)
            foundPlaces(foundCount) = place
            foundCount = foundCount + 1

            If place = nPlace Then
                isOccupied = True
            End If
        End If

        iRow = iRow + 1
    Loop

    IsPlaceOccupiedToday = Array(isOccupied, foundPlaces)
End Function

Sub DebugGun(isDateOk As Boolean, Item As Variant)
    If isDateOk Then
        MsgBox "⚠ isDateOk = True" & Chr(10) & _
                    "Row: " & Item(0) & Chr(10) & _
                "checkIn: " & Format(Item(1), "DD.MM.YYYY") & Chr(10) & _
               "checkOut: " & Format(Item(2), "DD.MM.YYYY") & Chr(10) & _
                   "code: " & Item(3) & Chr(10) & _
                  "place: " & Item(4) & Chr(10) & _
             "dTargetDay: " & Format(Item(5), "DD.MM.YYYY") & Chr(10) & _
              "sFullName: " & Item(6) & Chr(10) & _
                  "sName: " & Item(7)
    End If
End Sub

' =====================================================
' === Процедура CheckOccupiedPlace ====================
' =====================================================
' → Використовує IsPlaceOccupiedToday для перевірки зайнятості місця.
' → Якщо місце зайняте — показує повідомлення з переліком вільних місць.
Function CheckOccupiedPlace(oDialog As Object) As Boolean
    Dim nPlace As Long
    Dim nOffset As Long
    nPlace = Val(oDialog.getControl("PlaceCombo").getText())
    nOffset = Val(oDialog.getControl("OffsetField").getText())

    Dim result As Variant
    result = IsPlaceOccupiedToday(nPlace, nOffset)

    Dim isOccupied As Boolean
    isOccupied = result(0)

    Dim occupiedPlaces As Variant
    occupiedPlaces = result(1)

    Dim allPlaces As Variant
    allPlaces = Split(ALL_PLACES, ";")

    Dim freePlaces As String
    freePlaces = ""

    Dim i As Integer, j As Integer
    Dim isBusy As Boolean

    For i = LBound(allPlaces) To UBound(allPlaces)
        isBusy = False
        For j = LBound(occupiedPlaces) To UBound(occupiedPlaces)
            If CLng(allPlaces(i)) = CLng(occupiedPlaces(j)) Then
                isBusy = True
                Exit For
            End If
        Next j
        If Not isBusy Then
            If Len(freePlaces) > 0 Then
                freePlaces = freePlaces & ";"
            End If
            freePlaces = freePlaces & allPlaces(i)
        End If
    Next i

    If isOccupied Then
        ShowDialog "Місце № " & nPlace & " зайнято", "Вільні: ", freePlaces
        CheckOccupiedPlace = False
        Exit Function
    End If

    CheckOccupiedPlace = True
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

