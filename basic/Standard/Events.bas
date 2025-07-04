REM  *****  BASIC  *****

' Events.bas

' =====================================================
' === Процедура AddTextFieldsOffsetListener ===========
' =====================================================
' → Додає слухача змін у полі OffsetField.
' → Відображає або ховає поле ReasonField залежно від значення OffsetField.
' → Використовується для динамічного контролю видимості додаткових полів.
Sub AddTextFieldsOffsetListener(oDialog As Object)

    ' ==== Отримання посилань на елементи ====
    Dim oOffsetControl As Object
    Dim oReasonControl As Object
    Dim oLabelReasonControl As Object

    Set oOffsetControl = oDialog.getControl("OffsetField")
    Set oReasonControl = oDialog.getControl("ReasonField")
    Set oLabelReasonControl = oDialog.getControl("ReasonLabel")

    ' ==== Встановлення дефолтного стану ====
    If Val(oOffsetControl.getText()) = 0 Then
        oReasonControl.Visible = False
        oLabelReasonControl.Visible = False
    Else
        oReasonControl.Visible = True
        oLabelReasonControl.Visible = True
    End If

    ' ==== Створення слухача ====
    Dim oListener As Object
    oListener = CreateUnoListener("OffsetField_", "com.sun.star.awt.XTextListener")

    ' ==== Підключення слухача ====
    oOffsetControl.addTextListener(oListener)
End Sub

' =====================================================
' === Процедура OffsetField_textChanged ===============
' =====================================================
' → Викликається при зміні тексту у полі OffsetField.
' → Відображає або ховає поле ReasonField та його мітку залежно від значення.
' → Реалізує метод інтерфейсу com.sun.star.awt.XTextListener.
Sub OffsetField_textChanged(oEvent)
    Dim oDialog As Object
    Set oDialog = oEvent.Source.getContext()

    Dim oOffset As Object
    Dim oReason As Object
    Dim oLabelReason As Object

    Set oOffset = oDialog.getControl("OffsetField")
    Set oReason = oDialog.getControl("ReasonField")
    Set oLabelReason = oDialog.getControl("ReasonLabel")

    ' Логіка відображення або приховування
    If Val(oOffset.getText()) <> 0 Then
        oReason.Visible = True
        oLabelReason.Visible = True
    Else
        oReason.Visible = False
        oLabelReason.Visible = False
    End If
End Sub

' =====================================================
' === Процедура OffsetField_disposing =================
' =====================================================
' → Викликається при видаленні слухача OffsetField.
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub OffsetField_disposing(oEvent)
End Sub

' =====================================================
' === Процедура AddCodeComboListeners =================
' =====================================================
' → Додає слухачі подій до ComboBox "CodeCombo".
' → Відкриває список при кліку та згортає після вибору.
' → Також викликає оновлення залежних полів.
Sub AddCodeComboListeners(oDialog As Object)

    ' Отримання екземпляру ComboBox за іменем
    Dim oCombo As Object
    oCombo = oDialog.getControl("CodeCombo")

    ' === Слухач миші для події натискання ===
    ' Відповідає за відкриття списку шляхом зміни висоти
    Dim oMouseListener As Object
    oMouseListener = CreateUnoListener("CodeCombo_", "com.sun.star.awt.XMouseListener")
    oCombo.addMouseListener(oMouseListener)

    ' === Слухач вибору елементу зі списку ===
    ' Відповідає за згортання списку після вибору елементу
    Dim oItemListener As Object
    oItemListener = CreateUnoListener("CodeCombo_", "com.sun.star.awt.XItemListener")
    oCombo.addItemListener(oItemListener)
End Sub

' =====================================================
' === Процедура CodeCombo_mousePressed ================
' =====================================================
' → Викликається при натисканні на ComboBox "CodeCombo".
' → Збільшує висоту для візуального відкриття списку.
Sub CodeCombo_mousePressed(oEvent)
    Dim oControl As Object
    oControl = oEvent.Source.getModel()
    oControl.Height = 110 ' Підняли висоту
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
    Dim oControl As Object
    Dim oDialog As Object
    oDialog = oEvent.Source.getContext()
    oControl = oEvent.Source.getModel()
    oControl.Height = 15 ' Скриваємо список
    Call CalculatePaidFieldWithPlace(oDialog)
    Call UpdatePlaceCombo(oDialog)
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
    Dim oCombo As Object
    oCombo = oDialog.getControl("DurationCombo")

    ' ==== Слухач миші для відкриття списку ====
    Dim oMouseListener As Object
    oMouseListener = CreateUnoListener("DurationCombo_", "com.sun.star.awt.XMouseListener")
    oCombo.addMouseListener(oMouseListener)

    ' ==== Слухач вибору елементу ====
    Dim oItemListener As Object
    oItemListener = CreateUnoListener("DurationCombo_", "com.sun.star.awt.XItemListener")
    oCombo.addItemListener(oItemListener)
End Sub

' =====================================================
' === Процедура DurationCombo_mousePressed ============
' =====================================================
' → Викликається при натисканні на ComboBox "DurationCombo".
' → Збільшує висоту для візуального відкриття списку.
Sub DurationCombo_mousePressed(oEvent)
    Dim oControl As Object
    oControl = oEvent.Source.getModel()
    oControl.Height = 110 ' Відкриваємо список
End Sub

' =====================================================
' === Процедура DurationCombo_mouseReleased ===========
' =====================================================
' → Викликається при відпусканні кнопки миші з ComboBox "DurationCombo".
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub DurationCombo_mouseReleased(oEvent)
End Sub

' =====================================================
' === Процедура DurationCombo_mouseEntered ============
' =====================================================
' → Викликається при наведенні курсора на ComboBox "DurationCombo".
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub DurationCombo_mouseEntered(oEvent)
End Sub

' =====================================================
' === Процедура DurationCombo_mouseExited =============
' =====================================================
' → Викликається при виведенні курсора з ComboBox "DurationCombo".
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub DurationCombo_mouseExited(oEvent)
End Sub

' =====================================================
' === Процедура DurationCombo_itemStateChanged ========
' =====================================================
' → Викликається при зміні вибору в ComboBox "DurationCombo".
' → Згортає список та перераховує вартість у полі PaidField.
Sub DurationCombo_itemStateChanged(oEvent)
    Dim oControl As Object
    Dim oDialog As Object
    oDialog = oEvent.Source.getContext()
    oControl = oEvent.Source.getModel()
    oControl.Height = 15 ' Згортаємо ComboBox
    Call CalculatePaidFieldWithPlace(oDialog)
End Sub

' =====================================================
' === Процедура DurationCombo_disposing ===============
' =====================================================
' → Викликається при видаленні слухача з ComboBox "DurationCombo".
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub DurationCombo_disposing(oEvent)
End Sub

' =====================================================
' === Процедура AddPlaceComboListeners ================
' =====================================================
' → Додає слухачі подій до ComboBox "PlaceCombo".
' → Відкриває список при кліку та згортає після вибору.
' → Виконує перевірку зайнятості місця.
Sub AddPlaceComboListeners(oDialog As Object)
    Dim oCombo As Object
    oCombo = oDialog.getControl("PlaceCombo")

    ' ==== Слухач миші для відкриття списку ====
    Dim oMouseListener As Object
    oMouseListener = CreateUnoListener("PlaceCombo_", "com.sun.star.awt.XMouseListener")
    oCombo.addMouseListener(oMouseListener)

    ' ==== Слухач вибору елементу ====
    Dim oItemListener As Object
    oItemListener = CreateUnoListener("PlaceCombo_", "com.sun.star.awt.XItemListener")
    oCombo.addItemListener(oItemListener)
End Sub

' =====================================================
' === Процедура PlaceCombo_mousePressed ===============
' =====================================================
' → Викликається при натисканні на ComboBox "PlaceCombo".
' → Збільшує висоту для візуального відкриття списку.
Sub PlaceCombo_mousePressed(oEvent)
    Dim oControl As Object
    oControl = oEvent.Source.getModel()
    oControl.Height = PLACE_COMBO_HEIGHT ' Відкриваємо список
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
    Dim oControl As Object
    Dim oDialog As Object
    oControl = oEvent.Source.getModel()
    oDialog = oEvent.Source.getContext()
    oControl.Height = 15 ' Згортаємо список після вибору
    Call CalculatePaidFieldWithPlace(oDialog)
End Sub

' =====================================================
' === Процедура PlaceCombo_disposing ==================
' =====================================================
' → Викликається при видаленні слухача з ComboBox "PlaceCombo".
' → Порожня. Забезпечує відповідність інтерфейсу UNO.
Sub PlaceCombo_disposing(oEvent)
End Sub
