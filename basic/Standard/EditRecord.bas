REM  *****  BASIC  *****

' EditRecord.bas

Sub StartEdit()
    ' False - не шукає першу порожню комірку по стовпцю А. (SelectFirstEmptyInA)
    ' ResetPeopleTodayFilter(False) 
    If Not IsInAllowedColumns("редагування.") Then Exit Sub
    If Not IsInAllowedRows("редагування.") Then Exit Sub
    ' Позиціонує на стовпець А в тому ж рядку
    ' MoveCursorToColumnA()
    ' ACTION_EDIT - змінює логіку роботи форми
    ' Змінює назву кнопки
    ' Ініціалізує поля даними, які зчитані з рядка, що редагується 
    ShowForm(ACTION_EDIT)
End Sub

' =====================================================
' === Функція ReadFromTable ===========================
' =====================================================
' → Зчитує дані з активного рядка таблиці у Map.
' → Історію не додає.
Function ReadFromTable() As Variant
    Dim oSheet As Object
    Dim oSel   As Object
    Dim row    As Long
    Dim data   As Variant
    
    oSheet = ThisComponent.Sheets(0)
    oSel = ThisComponent.CurrentSelection
    row = oSel.RangeAddress.StartRow

    data = CreateMap()

    ' Заповнюємо Map
    MapPut data, "заселення", oSheet.getCellByPosition(0, row).String
    MapPut data, "прізвище", oSheet.getCellByPosition(1, row).String
    MapPut data, "ім'я по батькові", oSheet.getCellByPosition(2, row).String
    MapPut data, "код", oSheet.getCellByPosition(3, row).String
    MapPut data, "виселення", oSheet.getCellByPosition(4, row).String
    MapPut data, "сплачено", oSheet.getCellByPosition(5, row).String
    MapPut data, "видаток", oSheet.getCellByPosition(6, row).String
    MapPut data, "прихід", oSheet.getCellByPosition(7, row).String
    MapPut data, "коментар", oSheet.getCellByPosition(8, row).String
    MapPut data, "телефон", oSheet.getCellByPosition(9, row).String
    MapPut data, "паспортні дані", oSheet.getCellByPosition(10, row).String
    MapPut data, "дата народження", oSheet.getCellByPosition(11, row).String
    MapPut data, "чорний список", oSheet.getCellByPosition(12, row).String
    MapPut data, "хостел", oSheet.getCellByPosition(13, row).String
    MapPut data, "створено", oSheet.getCellByPosition(14, row).String
    MapPut data, "причина зсуву", oSheet.getCellByPosition(15, row).String
    MapPut data, "зсув", oSheet.getCellByPosition(16, row).String
    MapPut data, "місце", oSheet.getCellByPosition(17, row).String
    MapPut data, "історія", oSheet.getCellByPosition(18, row).String
    
    ' MsgBOx MapGet(data, "місце")

    ReadFromTable = data
End Function

' =====================================================
' === Процедура ShowFieldsInMsgBox ====================
' =====================================================
' → Виводить дані з Map у MsgBox (крім поля "історія")
Sub ShowFieldsInMsgBox(Fields As Variant)
    Dim i     As Integer
    Dim pair  As Variant
    Dim sText As String
    text = ""

    For i = LBound(Fields) To UBound(Fields)
        pair = MapGetByIndex(Fields, i)
        'If pair(0) <> "історія" Then
            sText = sText & pair(0) & ": " & pair(1) & Chr(10)
        'End If
    Next i

    MsgBox text, 64, "Дані з рядка"
End Sub

' =====================================================
' === Функція IsInAllowedColumns ======================
' =====================================================
' → Перевіряє, чи курсор знаходиться в межах стовпців A–S.
' → Повертає True, якщо так, інакше False та показує повідомлення.
Function IsInAllowedColumns(Optional sAddTitle As Variant) As Boolean
    If IsMissing(sAddTitle) Or Len(Trim(sAddTitle)) = 0 Then sAddTitle = ""
    Dim oDoc As Object
    Dim oSel As Object
    Dim col  As Long
    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection
    col = oSel.RangeAddress.StartColumn
    
    If col < 0 Or col > 20 Then
        MsgDlg "Помилка " & sAddTitle, "Курсор має бути в межах стовпців A–U.", False, 50, 140
        IsInAllowedColumns = False
    Else
        IsInAllowedColumns = True
    End If
End Function

' =====================================================
' === Функція IsInAllowedRows =========================
' =====================================================
' → Перевіряє, чи курсор знаходиться в дозволеному діапазоні рядків.
' → Викликає GetAfterLastEncashRange для отримання меж діапазону.
' → Якщо курсор всередині — OK.
' → Якщо поза й треба пароль — просить пароль.
' → Якщо неможливо редагувати — виводить повідомлення.
' → Повертає True, якщо редагування дозволено.
Function IsInAllowedRows(Optional sAddTitle As Variant) As Boolean
    If IsMissing(sAddTitle) Or Len(Trim(sAddTitle)) = 0 Then sAddTitle = ""
    Dim lStartRow As Long
    Dim lEndRow   As Long
    Dim lRow      As Long
    Dim lValueD   As Long
    Dim sValueS   As String
    Dim aRange    As Variant
    Dim oDoc      As Object
    Dim oSel      As Object
    Dim oSheet    As Object


    oSheet = ThisComponent.Sheets(0)

    
    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection
    lRow = oSel.RangeAddress.StartRow
    lValueD = CLng(oSheet.getCellByPosition(3, lRow).getValue)
    sValueS = Trim(oSheet.getCellByPosition(3, lRow).getValue)
    
    aRange = GetAfterLastEncashRange()
    lStartRow = aRange(0)
    lEndRow = aRange(1)
    
    Select Case lValueD
        Case 28
            MsgDlg "Відмова " & sAddTitle, "Строка ""чорний список"" не може бути отредаговане.", False, 50, 180
            IsInAllowedRows = False
            Exit Function

        Case 30
            MsgDlg "Відмова " & sAddTitle, "Строка ""баланс"" не може бути отредаговане.", False, 50, 160
            IsInAllowedRows = False
            Exit Function

        Case 7
            MsgDlg "Відмова " & sAddTitle, "Строка ""інкасація"" не може бути отредаговане.", False, 50, 160
            IsInAllowedRows = False
            Exit Function
            
        Case 20
            MsgDlg "Відмова " & sAddTitle, "Строка ""скасовано"" не може бути отредаговане.", False, 50, 170
            IsInAllowedRows = False
            Exit Function
    End Select
    
    ' якщо курсор всередині діапазону — ОК
    If lRow >= lStartRow And lRow <= lEndRow Then
        IsInAllowedRows = True
        Exit Function
    End If
    
    Dim bOutOfRange As Boolean, bHasData As Boolean
    bOutOfRange = (lStartRow = 0 And lEndRow = 0) Or (lRow < lStartRow)
    bHasData = lRow > 2 And sValueE <> ""

    ' якщо діапазон = [0,0], курсор >2, E не порожня, просимо пароль
    If bOutOfRange And bHasData Then
        If Not ShowPasswordDialog(NEGET_RULES) Then
            MsgDlg "Відмова " & sAddTitle, "Операцію скасовано.", False, 55, 130
            IsInAllowedRows = False
            Exit Function
        End If

        IsInAllowedRows = True
        Exit Function
    End If
    
    IsInAllowedRows = isBelowEncashRange(Array(lRow, lStartRow, lEndRow))
End Function

' =====================================================
' === Функція isBelowEncashRange ======================
' =====================================================
' → Перевіряє, чи курсор не нижче дозволеного діапазону.
' → Повертає True, якщо курсор у допустимих межах.
' → Показує повідомлення про помилку, якщо ні.
Function isBelowEncashRange(aRange As Variant) As Boolean
    Dim lRow As Long,lStartRow As Long, lEndRow As Long
    lRow = aRange(0) 
    lStartRow = aRange(1)
    lEndRow = aRange(2)
    
    isBelowEncashRange = False
    
    ' якщо діапазону немає
    If lStartRow = 0 And lEndRow = 0 Then
        MsgDlg "Помилка редагування.", "Діапазон дозволених рядків відсутній.", False, 55
        Exit Function
    End If
    
    ' якщо тільки один рядок доступний
    If lStartRow = lEndRow Then
        MsgDlg "Помилка редагування.", "До редагування доступен лише один рядок: " & (lStartRow + 1) & ".", False, 55
        Exit Function
    End If
    
    ' якщо курсор поза межами
    If lRow < lStartRow Or lRow > lEndRow Then
        MsgDlg "Помилка редагування.", "Курсор поза дозволеним діапазоном рядків: " & (lStartRow + 1) & "–" & (lEndRow + 1) & ".", False, 55
        Exit Function
    End If
    
    isBelowEncashRange = True
End Function

' =====================================================
' === Процедура MoveCursorToColumnA ===================
' =====================================================
' → Переносить курсор у стовпець A поточного рядка.
Sub MoveCursorToColumnA()
    Dim oDoc, oCtrl, oSel, row
    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection
    row = oSel.RangeAddress.StartRow

    oCtrl = oDoc.CurrentController
    oCtrl.select(oDoc.Sheets(0).getCellByPosition(0, row))
End Sub

' =====================================================
' === Функція FormatHistoryLine =======================
' =====================================================
' → Приймає `data` (Map із полями рядка).
' → Формує рядок історії у вигляді: значення1 | значення2 | … | значенняN.
' → Використовується для збереження змін у полі "історія" (S).
Function FormatHistoryLine(data As Variant) As String
    Dim keys As Variant
    Dim i As Long, sLine As String

    keys = Array( _
        "заселення", "прізвище", "ім'я по батькові", "код", "виселення", _
        "сплачено", "видаток", "прихід", "коментар", "телефон", _
        "паспортні дані", "дата народження", "чорний список", "хостел", _
        "створено", "причина зсуву", "зсув", "місце" _
    )

    sLine = ""
    For i = LBound(keys) To UBound(keys)
        If i > 0 Then sLine = sLine & " | "
        sLine = sLine & MapGet(data, keys(i))
    Next i

    FormatHistoryLine = sLine
End Function

' =====================================================
' === Функція AppendHistory ===========================
' =====================================================
' → Зчитує всі поля поточного рядка та формує новий рядок історії.
' → Додає його на початок існуючої історії (якщо вона є).
' → Повертає `True`, якщо запис історії успішний.
' → Використовується під час редагування для відстеження змін.
Function AppendHistory(row As Long) As Boolean
    On Error GoTo ErrHandler
    Dim oSheet As Object, data As Variant
    Dim sOldHistory As String, sNewHistory As String, sLine As String

    oSheet = ThisComponent.Sheets(0)
    data = ReadFromTable()
    sLine = FormatHistoryLine(data)
    sOldHistory = oSheet.getCellByPosition(18, row).String

    If Len(Trim(sOldHistory)) > 0 Then
        sNewHistory = sLine & Chr(10) & sOldHistory
    Else
        sNewHistory = sLine
    End If

    oSheet.getCellByPosition(18, row).String = sNewHistory
    AppendHistory = True
    Exit Function

ErrHandler:
    MsgDlg "Помилка редагування.", "Помилка запису історії: " & Err.Description, False, 55
    AppendHistory = False
End Function

