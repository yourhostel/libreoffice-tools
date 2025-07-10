REM  *****  BASIC  *****

' Encashment.bas

REM  *****  BASIC  *****

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
    Dim oSelectedCell    As Object
    Dim lRowNumber       As Long
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

    ' === позиціонування на перший порожній рядок ===
    SelectFirstEmptyInA()
    oSelectedCell = oDocument.CurrentSelection
    lRowNumber = oSelectedCell.RangeAddress.StartRow

    ' === знаходимо діапазон для інкасації ===
    aRange = FindEncashmentRange(lRowNumber)
    lStartRow = aRange(0)
    lEndRow = aRange(1)

    ' === перевіряємо чи є що інкасувати ===
    If lStartRow = 0 And lEndRow = 0 Then
        ShowDialog "Інкасація не потрібна", "Після останньої інкасації немає повних записів."
        Exit Sub
    End If

    ' === підрахунок суми інкасації ===
    nTotalEncash = CalculateEncashment(lStartRow, lEndRow)

    ' === вставка інкасації ===
    InsertEncashment lRowNumber, nTotalEncash

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
' === Функція FindEncashmentRange ====================
' =====================================================
' → Визначає діапазон рядків для інкасації.
' → Шукає останній запис інкасації й повертає масив [start, end].
' → Якщо немає що інкасувати — повертає [0, 0].
Function FindEncashmentRange(lRowNumber As Long) As Variant
    Dim oSheet      As Object
    Dim lStartRow   As Long
    Dim lEndRow     As Long
    Dim lCheckRow   As Long

    oSheet = ThisComponent.Sheets(0)

    ' шукаємо останню інкасацію від курсора вгору
    For lCheckRow = lRowNumber - 1 To 3 Step -1
        If Trim(oSheet.getCellByPosition(4, lCheckRow).String) = ENCASH Then
            lStartRow = lCheckRow + 1
            Exit For
        End If
    Next

    If lStartRow = 0 Then lStartRow = 3

    lEndRow = lRowNumber - 1

    If lEndRow < lStartRow Then
        FindEncashmentRange = Array(0, 0) ' немає діапазону
    Else
        FindEncashmentRange = Array(lStartRow, lEndRow)
    End If
End Function

' =====================================================
' === Процедура InsertEncashment =====================
' =====================================================
' → Вставляє запис про інкасацію в обраний рядок.
' → Записує дату, тип, суму й показує повідомлення про успіх.
Sub InsertEncashment(lRowNumber As Long, nTotalEncash As Double)
    Dim oDocument        As Object
    Dim oSheet           As Object
    Dim oFormats         As Object
    Dim oLocale          As New com.sun.star.lang.Locale
    Dim lFormatDate      As Long
    Dim lFormatDateTime  As Long
    Dim oCell            As Object
    Dim dToday           As Date

    oDocument = ThisComponent
    oSheet = oDocument.Sheets(0)

    dToday = Now()

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
