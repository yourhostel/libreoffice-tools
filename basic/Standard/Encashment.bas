REM  *****  BASIC  *****

' Encashment.bas

' =====================================================
' === Процедура DoEncashment ==========================
' =====================================================
' → Запускає інкасацію: перевіряє пароль, знаходить діапазон, підраховує та вставляє запис.
' → Якщо помилка (порожня таблиця, немає записів для інкасації) — показує повідомлення й виходить.
Sub DoEncashment()
    If Not ShowPasswordDialog(NEGET_RULES) Then
        MsgBox "Операцію скасовано.", 48, "Відмова"
        Exit Sub
    End If

    Dim oDocument        As Object
    Dim oSheet           As Object
    Dim aRange            As Variant
    Dim lStartRow         As Long
    Dim lEndRow           As Long
    Dim nTotalEncash      As Double

    ' === ініціалізація документа та аркуша ===
    oDocument = ThisComponent
    oSheet = oDocument.Sheets(0)

    ' === перевірка: таблиця порожня ===
    If Trim(oSheet.getCellByPosition(0, 3).String) = "" Then
        ShowDialog "Помилка", "Таблиця порожня. Інкасувати нема чого."
        Exit Sub
    End If

    ' === знаходимо діапазон для інкасації ===
    aRange = GetAfterLastEncashRange()
    lStartRow = aRange(0)
    lEndRow = aRange(1)

    ' === перевіряємо чи є що інкасувати ===
    If lStartRow = 0 And lEndRow = 0 Then
        ShowDialog "Інкасація не потрібна", "Після останньої інкасації немає нових записів."
        Exit Sub
    End If

    ' === підрахунок суми інкасації ===
    nTotalEncash = CalculateEncashment(lStartRow, lEndRow)

    ' === вставка інкасації ===
    InsertEncashment nTotalEncash

End Sub

' =====================================================
' === Функція CalculateEncashment ====================
' =====================================================
' → Підраховує загальну суму інкасації за діапазоном рядків.
' → Рахує: оплачено - витрати + доходи.
Function CalculateEncashment(lStartRow As Long, lEndRow As Long) As Double
    Dim oDocument     As Object
    Dim oSheet        As Object
    Dim lRowNumber    As Long
    Dim nSumPaid      As Double
    Dim nSumExpense   As Double
    Dim nSumIncome    As Double

    oDocument = ThisComponent
    oSheet = oDocument.Sheets(0)

    nSumPaid = 0
    nSumExpense = 0
    nSumIncome = 0

    For lRowNumber = lStartRow To lEndRow
        nSumPaid = nSumPaid + oSheet.getCellByPosition(5, lRowNumber).getValue()
        nSumExpense = nSumExpense + oSheet.getCellByPosition(6, lRowNumber).getValue()
        nSumIncome = nSumIncome + oSheet.getCellByPosition(7, lRowNumber).getValue()
    Next lRowNumber

    CalculateEncashment = nSumPaid - nSumExpense + nSumIncome
End Function

' =====================================================
' === Процедура InsertEncashment =====================
' =====================================================
' → Вставляє запис про інкасацію в обраний рядок.
' → Записує дату, тип, суму й показує повідомлення про успіх.
Sub InsertEncashment(nTotalEncash As Double)
    Dim oDocument        As Object
    Dim oSheet           As Object
    Dim oFormats         As Object
    Dim oLocale          As New com.sun.star.lang.Locale
    Dim lFormatDate      As Long
    Dim lFormatDateTime  As Long
    Dim oCell            As Object
    Dim dToday           As Date
    Dim oSelectedCell    As Object
    Dim lRowNumber       As Long

    oDocument = ThisComponent
    oSheet = oDocument.Sheets(0)

    dToday = Now()

    ' === позиціонування на перший порожній рядок ===
    SelectFirstEmptyInA()

    oSelectedCell = oDocument.CurrentSelection
    lRowNumber = oSelectedCell.RangeAddress.StartRow

    ' формати
    oFormats = oDocument.getNumberFormats()
    oLocale.Language = "uk"
    oLocale.Country = "UA"

    lFormatDate = oFormats.queryKey("DD.MM.YYYY", oLocale, True)
    If lFormatDate = -1 Then lFormatDate = oFormats.addNew("DD.MM.YYYY", oLocale)

    lFormatDateTime = oFormats.queryKey("DD.MM.YYYY HH:MM", oLocale, True)
    If lFormatDateTime = -1 Then lFormatDateTime = oFormats.addNew("DD.MM.YYYY HH:MM", oLocale)

    ' запис
    oCell = oSheet.getCellByPosition(0, lRowNumber)
    oCell.setValue(dToday)
    oCell.NumberFormat = lFormatDate

    oSheet.getCellByPosition(3, lRowNumber).setValue(7)
    oSheet.getCellByPosition(4, lRowNumber).setString(ENCASH)

    oCell = oSheet.getCellByPosition(14, lRowNumber)
    oCell.setValue(dToday)
    oCell.NumberFormat = lFormatDateTime

    oSheet.getCellByPosition(5, lRowNumber).setValue(nTotalEncash)

    ShowDialog "Інкасація виконана", "Сума інкасації: " & nTotalEncash
End Sub
