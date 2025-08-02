REM  *****  BASIC  *****

' Events.bas

' =====================================================
' === Процедура AddDropdownToCombo ====================
' =====================================================
' → Підключає до ComboBox три слухачі:
'     • XItemListener   — для зміни елемента
'     • XMouseListener  — для реагування на кліки
'     • XTextListener   — для зміни тексту вручну
' → Зберігає назву елемента в .Tag для подальшого використання.
Sub AddDropdownToCombo(oDialog As Object, sControlName As String)
    Dim oCombo         As Object
    Dim oItemListener  As Object
    Dim oMouseListener As Object
    
    oCombo = oDialog.getControl(sControlName)
    oCombo.Model.Tag = sControlName
    
    oItemListener   = CreateUnoListener("DropdownShared_", "com.sun.star.awt.XItemListener")
    oMouseListener  = CreateUnoListener("DropdownShared_", "com.sun.star.awt.XMouseListener")
    oTextListener   = CreateUnoListener("DropdownShared_", "com.sun.star.awt.XTextListener") 

    oCombo.addItemListener(oItemListener)
    oCombo.addMouseListener(oMouseListener)
    oCombo.addTextListener(oTextListener)
End Sub

' =====================================================
' === Процедура DropdownShared_textChanged =============
' =====================================================
' → Реагує на зміну тексту у ComboBox.
' → Викликає AddComboReadonlyEnforcer із захистом від рекурсії через флаг isHandlingComboTextChange.
Global isHandlingComboTextChange As Boolean
Sub DropdownShared_textChanged(oEvent)
    If isHandlingComboTextChange Then Exit Sub
    isHandlingComboTextChange = True    
        AddComboReadonlyEnforcer(oEvent)
    isHandlingComboTextChange = False
End Sub

' =====================================================
' === Процедура DropdownShared_mousePressed ============
' =====================================================
' → Спрацьовує при натисканні на ComboBox.
' → Збільшує висоту елемента (наприклад, до 45), щоб імітувати розкритий список.
Sub DropdownShared_mousePressed(oEvent)
    DropdownToCombo(oEvent, 50)
End Sub

' =====================================================
' === Процедура DropdownShared_itemStateChanged ========
' =====================================================
' → Спрацьовує при виборі нового елемента у ComboBox.
' → Зменшує висоту поля назад до стандартної (наприклад, 15).
Sub DropdownShared_itemStateChanged(oEvent)
    DropdownToCombo(oEvent, 15)
End Sub

' =====================================================
' === Порожні методи-обгортки XMouseListener ===========
' =====================================================
' → Реалізовані як no-op: mouseReleased, mouseEntered, mouseExited, disposing.
' → Не виконують дій, але потрібні для повної реалізації інтерфейсу.
Sub DropdownShared_mouseReleased(oEvent)
End Sub

Sub DropdownShared_mouseEntered(oEvent)
End Sub

Sub DropdownShared_mouseExited(oEvent)
End Sub

Sub DropdownShared_disposing(oEvent)
End Sub

' =====================================================
' === Процедура AddTextFieldsOffsetListener ===========
' =====================================================
' → Додає слухача змін у полі OffsetField.
' → Відображає або ховає поле ReasonField залежно від значення OffsetField.
' → Використовується для динамічного контролю видимості додаткових полів.
Sub AddCheckInDataListener(oDialog As Object, sAction As String)
    Dim oListener           As Object
    Dim oTargetDateControl  As Object
    
    OffsetCommentSwitch(oDialog, True, sAction)
       
    oTargetDateControl   = oDialog.getControl("CheckInDate")
    oTargetDateControl.Model.Tag = sAction
    ' ==== Створення слухача ====
    oListener = CreateUnoListener("CheckInData_", "com.sun.star.awt.XTextListener")

    ' ==== Підключення слухача ====
    oTargetDateControl.addTextListener(oListener)
End Sub

' =====================================================
' === Процедура OffsetField_textChanged ===============
' =====================================================
' → Викликається при зміні тексту у полі OffsetField.
' → Відображає або ховає поле ReasonField та його мітку залежно від значення.
' → Реалізує метод інтерфейсу com.sun.star.awt.XTextListener.
Sub CheckInData_textChanged(oEvent)  
    Dim oDialog As Object
    oDialog = oEvent.Source.getContext()    
    sAction = oEvent.Source.Model.Tag
    
    OffsetCommentSwitch(oDialog, False, sAction)
    CalculationTerm(oDialog)
    CalculatePaidFieldWithPlace(oDialog)    
End Sub

' =====================================================
' === Процедура OffsetField_disposing =================
' =====================================================
' → Викликається при видаленні слухача OffsetField.
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub CheckInData_disposing(oEvent)
End Sub

' =====================================================
' === Процедура AddCodeComboListeners =================
' =====================================================
' → Додає слухачі подій до ComboBox "CodeCombo".
' → Відкриває список при кліку та згортає після вибору.
' → Також викликає оновлення залежних полів.
Sub AddCodeComboListeners(oDialog As Object)
    Dim oCombo         As Object
    Dim oItemListener  As Object
       
    ' Отримання екземпляру ComboBox за іменем
    oCombo = oDialog.getControl("CodeCombo")

    ' === Слухач вибору елементу зі списку ===
    ' Відповідає за згортання списку після вибору елементу
    oItemListener = CreateUnoListener("CodeCombo_", "com.sun.star.awt.XItemListener")
    oCombo.addItemListener(oItemListener)
End Sub

' =====================================================
' === Процедура CodeCombo_mousePressed ================
' =====================================================
' → Викликається при натисканні на ComboBox "CodeCombo".
' → Збільшує висоту для візуального відкриття списку.
Sub CodeCombo_mousePressed(oEvent)
End Sub

' =====================================================
' === Процедура CodeCombo_mouseReleased ===============
' =====================================================
' → Викликається при відпусканні кнопки миші з ComboBox "CodeCombo".
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub CodeCombo_mouseReleased(oEvent)
End Sub

' =====================================================
' === Процедура CodeCombo_mouseEntered ================
' =====================================================
' → Викликається при наведенні курсора на ComboBox "CodeCombo".
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub CodeCombo_mouseEntered(oEvent)
End Sub

' =====================================================
' === Процедура CodeCombo_mouseExited =================
' =====================================================
' → Викликається при виведенні курсора з ComboBox "CodeCombo".
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub CodeCombo_mouseExited(oEvent)
End Sub

' =====================================================
' === Процедура CodeCombo_itemStateChanged ============
' =====================================================
' → Викликається при зміні вибору в ComboBox "CodeCombo".
' → Згортає список, оновлює пов’язані поля та перевіряє зайнятість місця.
Sub CodeCombo_itemStateChanged(oEvent)
    'Dim oControl As Object
    Dim oDialog  As Object
    
    oDialog  = oEvent.Source.getContext()
    'oControl = oEvent.Source.getModel()
    'oControl.Height = 15 ' Скриваємо список
    CalculatePaidFieldWithPlace(oDialog)
    UpdatePlaceCombo(oDialog)
End Sub

' =====================================================
' === Процедура CodeCombo_disposing ===================
' =====================================================
' → Викликається при видаленні слухача з ComboBox "CodeCombo".
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub CodeCombo_disposing(oEvent)
End Sub

' =====================================================
' === Процедура AddDurationComboListeners =============
' =====================================================
' → Додає слухачі подій до ComboBox "DurationCombo".
' → Відкриває список при кліку та згортає після вибору.
' → Перераховує вартість у полі PaidField.
Sub AddDurationComboListeners(oDialog As Object)
    Dim oCombo           As Object
    Dim oItemListener    As Object
     
    oCombo = oDialog.getControl("DurationCombo")

    ' ==== Слухач вибору елементу ====
    oItemListener = CreateUnoListener("Duration_", "com.sun.star.awt.XItemListener")
    oCombo.addItemListener(oItemListener)
End Sub

' =====================================================
' === Процедура DurationCombo_mousePressed ============
' =====================================================
' → Викликається при натисканні на ComboBox "DurationCombo".
Sub Duration_mousePressed(oEvent)
End Sub

' =====================================================
' === Процедура DurationCombo_mouseReleased ===========
' =====================================================
' → Викликається при відпусканні кнопки миші з ComboBox "DurationCombo".
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub Duration_mouseReleased(oEvent)
End Sub

' =====================================================
' === Процедура DurationCombo_mouseEntered ============
' =====================================================
' → Викликається при наведенні курсора на ComboBox "DurationCombo".
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub Duration_mouseEntered(oEvent)
End Sub

' =====================================================
' === Процедура DurationCombo_mouseExited =============
' =====================================================
' → Викликається при виведенні курсора з ComboBox "DurationCombo".
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub Duration_mouseExited(oEvent)
End Sub

' =====================================================
' === Процедура DurationCombo_itemStateChanged ========
' =====================================================
' → Викликається при зміні вибору в ComboBox "DurationCombo".
' → Згортає список та перераховує вартість у полі PaidField.
Sub Duration_itemStateChanged(oEvent)
    Dim oDialog   As Object
        oDialog     = oEvent.Source.getContext()

    CalculationTerm(oDialog)
    CalculatePaidFieldWithPlace(oDialog)
End Sub

' =====================================================
' === Процедура DurationCombo_disposing ===============
' =====================================================
' → Викликається при видаленні слухача з ComboBox "DurationCombo".
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub Duration_disposing(oEvent)
End Sub

' =====================================================
' === Процедура AddFinListeners =======================
' =====================================================
' → Додає слухачі змін у полях "ExpenseField" та "IncomeField".
' → Викликає FinCommentSwitch для початкового налаштування.
' → Слухачі оновлюють видимість поля коментаря залежно від введених сум.
' → Використовується для контролю коректності фінансових полів.
Sub AddFinListeners(oDialog As Object)
    Dim oListener As Object
    Dim oExpense  As Object
    Dim oIncome   As Object

    oListener = CreateUnoListener("Fin_", "com.sun.star.awt.XTextListener")
    
    FinCommentSwitch(oDialog)
    
    oExpense = oDialog.getControl("ExpenseField")
    oIncome  = oDialog.getControl("IncomeField")

    oExpense.addTextListener(oListener)
    oIncome.addTextListener(oListener)
End Sub

' =====================================================
' === Процедура Fin_textChanged =======================
' =====================================================
' → Викликається при зміні тексту у полях "ExpenseField" або "IncomeField".
' → Перевіряє значення полів і показує/ховає поле коментаря.
' → Реалізує метод XTextListener.
Sub Fin_textChanged(oEvent)
    Dim oDialog As Object
    oDialog = oEvent.Source.getContext()

    FinCommentSwitch(oDialog)
End Sub

' =====================================================
' === Процедура Fin_disposing =========================
' =====================================================
' → Викликається при видаленні слухача з фінансових полів.
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub Fin_disposing(oEvent)
End Sub

' =====================================================
' === Процедура AddPlaceComboListeners ================
' =====================================================
' → Додає слухачі подій до ComboBox "PlaceCombo".
' → Відкриває список при кліку та згортає після вибору.
' → Виконує перевірку зайнятості місця.
Sub AddPlaceComboListeners(oDialog As Object)
    Dim oCombo         As Object
    Dim oItemListener  As Object
    
    oCombo = oDialog.getControl("PlaceCombo")

    ' ==== Слухач вибору елементу ====
    oItemListener = CreateUnoListener("PlaceCombo_", "com.sun.star.awt.XItemListener")
    oCombo.addItemListener(oItemListener)
End Sub

' =====================================================
' === Процедура PlaceCombo_mousePressed ===============
' =====================================================
' → Викликається при натисканні на ComboBox "PlaceCombo".
' → Збільшує висоту для візуального відкриття списку.
Sub PlaceCombo_mousePressed(oEvent)
End Sub

' =====================================================
' === Процедура PlaceCombo_mouseReleased ==============
' =====================================================
' → Викликається при відпусканні кнопки миші з ComboBox "PlaceCombo".
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub PlaceCombo_mouseReleased(oEvent)
End Sub

' =====================================================
' === Процедура PlaceCombo_mouseEntered ===============
' =====================================================
' → Викликається при наведенні курсора на ComboBox "PlaceCombo".
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub PlaceCombo_mouseEntered(oEvent)
End Sub

' =====================================================
' === Процедура PlaceCombo_mouseExited ================
' =====================================================
' → Викликається при виведенні курсора з ComboBox "PlaceCombo".
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub PlaceCombo_mouseExited(oEvent)
End Sub

' =====================================================
' === Процедура PlaceCombo_itemStateChanged ===========
' =====================================================
' → Викликається при зміні вибору в ComboBox "PlaceCombo".
' → Згортає список, оновлює вартість та перевіряє зайнятість місця.
Sub PlaceCombo_itemStateChanged(oEvent)
    Dim oDialog As Object
  
    oDialog = oEvent.Source.getContext()
    CalculatePaidFieldWithPlace(oDialog)
End Sub

' =====================================================
' === Процедура PlaceCombo_disposing ==================
' =====================================================
' → Викликається при видаленні слухача з ComboBox "PlaceCombo".
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub PlaceCombo_disposing(oEvent)
End Sub

Sub OnSelectionChanged(oEvent)
  Dim ctrl: ctrl = ThisComponent.CurrentController
  Dim fvRow: fvRow = ctrl.FirstVisibleRow
  Dim limitRow As Long: limitRow = ThisComponent.Sheets(0).UsedRange.EndRow - 5

  If fvRow < limitRow Then
    ctrl.setFirstVisibleRow(limitRow)
  End If
End Sub

