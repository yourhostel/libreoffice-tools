REM  *****  BASIC  *****

' Events.bas

' ==== Слухач для зміни тексту у полі OffsetField ====
' Реалізує інтерфейс com.sun.star.awt.XTextListener
' Відповідає за динамічне відображення або приховування поля ReasonField
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

' === Обробка зміни тексту у полі OffsetField ===
' Реалізує метод інтерфейсу com.sun.star.awt.XTextListener
' Відповідає за динамічне відображення або приховування поля ReasonField і мітки ReasonLabel
' Логіка:
'   — Якщо значення Offset ≠ 0 → поле ReasonField і мітка ReasonLabel стають видимими
'   — Якщо значення Offset = 0 → поле ReasonField і мітка ReasonLabel приховуються
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

' ==== Обов’язковий метод для інтерфейсу XTextListener ====
' Викликається при видаленні слухача або знищенні елемента
' Не виконує дій. Забезпечує відповідність UNO API.
Sub OffsetField_disposing(oEvent)
    ' Порожньо. Необхідно для відповідності інтерфейсу.
End Sub


' === Додає обробники подій для ComboBox ===
' oDialog — екземпляр діалогу, у якому знаходиться ComboBox
Sub AddComboListeners(oDialog As Object)

    ' Отримання екземпляру ComboBox за іменем
    Dim oCombo As Object
    oCombo = oDialog.getControl("CodeCombo")

    ' === Слухач миші для події натискання ===
    ' Відповідає за відкриття списку шляхом зміни висоти
    Dim oMouseListener As Object
    oMouseListener = CreateUnoListener("Combo_", "com.sun.star.awt.XMouseListener")
    oCombo.addMouseListener(oMouseListener)

    ' === Слухач вибору елементу зі списку ===
    ' Відповідає за згортання списку після вибору елементу
    Dim oItemListener As Object
    oItemListener = CreateUnoListener("Combo_", "com.sun.star.awt.XItemListener")
    oCombo.addItemListener(oItemListener)

End Sub

' === Обробка натискання кнопки миші на ComboBox ===
' Збільшує висоту ComboBox для візуального розкриття списку
Sub Combo_mousePressed(oEvent)
    Dim oControl As Object
    ' Отримання моделі елемента керування
    oControl = oEvent.Source.getModel()
    ' Зміна висоти на 60 для розкриття списку
    oControl.Height = 110 ' Підняли висоту
End Sub

' === Обробка відпускання кнопки миші ===
' Не використовується, необхідно для відповідності інтерфейсу
Sub Combo_mouseReleased(oEvent)
    ' Можна нічого
End Sub

' === Обробка наведення курсора на ComboBox ===
' Порожньо. Метод обов’язковий для XMouseListener.
Sub Combo_mouseEntered(oEvent)
End Sub

' === Обробка виведення курсора за межі ComboBox ===
' Порожньо. Метод обов’язковий для XMouseListener.
Sub Combo_mouseExited(oEvent)
End Sub

' === Обробка зміни вибраного елемента в ComboBox ===
' Після вибору елемента список згортається шляхом зменшення висоти
Sub Combo_itemStateChanged(oEvent)
    Dim oControl As Object
    Dim oDialog As Object
    ' Отримуємо діалог з контексту події
    oDialog = oEvent.Source.getContext()
    ' Отримання моделі елемента керування
    oControl = oEvent.Source.getModel()
    ' Зміна висоти на 15 для згортання списку
    oControl.Height = 15 ' Скрываем список

    CalculatePaidFieldByDuration(oEvent)
    LockFields(oEvent, "DurationField;OffsetField")
End Sub

' === Звільнення ресурсів для ComboBox ===
' Порожній метод. Необхідний для відповідності інтерфейсу.
Sub Combo_disposing(oEvent)
End Sub

' === Підключає слухач для поля DurationField ===
Sub AddTextFieldsDurationListener(oDialog As Object)

    Dim oDurationControl As Object
    Set oDurationControl = oDialog.getControl("DurationField")

    ' ==== Створюємо слухача ====
    Dim oListener As Object
    Set oListener = CreateUnoListener("DurationField_", "com.sun.star.awt.XTextListener")

    ' ==== Підключаємо ====
    oDurationControl.addTextListener(oListener)

End Sub

' === Обробка зміни тексту у полі DurationField ===
Sub DurationField_textChanged(oEvent)

    ' Отримуємо діалог
    Dim oDialog As Object
    Set oDialog = oEvent.Source.getContext()

    ' Отримуємо значення Duration
    Dim nDuration As Long
    nDuration = Val(oDialog.getControl("DurationField").getText())

    ' ==== Відкриваємо лист з цінами ====
    Dim oDoc As Object, oSheet As Object
    Set oDoc = ThisComponent
    Set oSheet = oDoc.Sheets.getByName("price1")

    ' ==== Пошук ціни ====
    Dim iRow As Long
    Dim dPrice As Double
    dPrice = 0

    For iRow = 1 To 10 ' Рядки з 2 по 11 (індексація з 0)
        If oSheet.getCellByPosition(0, iRow).getValue() = nDuration Then
            dPrice = oSheet.getCellByPosition(1, iRow).getValue()
            Exit For
        End If
    Next iRow

    ' ==== Запис у поле Paid ====
    oDialog.getControl("PaidField").setText(CStr(dPrice))

End Sub

' === Звільнення ресурсів ===
Sub DurationField_disposing(oEvent)
    ' Порожньо. Необхідно для відповідності інтерфейсу.
End Sub