REM  *****  BASIC  *****

' EditRecord.bas

Sub StartEdit()
    Dim aAllRows As Variant : aAllRows = IsInAllowedRows("редагування.")
    
    If Not IsInAllowedColumns("редагування.") Then Exit Sub    
    If Not aAllRows(0) Then Exit Sub
      
    ' Позиціонує на стовпець А в тому ж рядку
    ' MoveCursorToColumnA()
    ' ACTION_EDIT - змінює логіку роботи форми
    ' Змінює назву кнопки
    ' Ініціалізує поля даними, які зчитані з рядка, що редагується 
    ShowForm(ACTION_EDIT, aAllRows(1))
End Sub

' =====================================================
' === Функція ReadFromTable ===========================
' =====================================================
' → Зчитує дані з активного рядка таблиці у Map.
Function ReadFromTable(Optional row As Variant) As Variant
    Dim oSheet As Object
    Dim oSel   As Object
    Dim data   As Variant
    
    oSheet = ThisComponent.Sheets(0)
    data   = CreateMap()
    
    If IsMissing(row) Then
        oSel = ThisComponent.CurrentSelection
        row = oSel.RangeAddress.StartRow
    End If

    ' Заповнюємо Map
    MapPut data, "заселення",        oSheet.getCellByPosition(0,  row).String
    MapPut data, "прізвище",         oSheet.getCellByPosition(1,  row).String
    MapPut data, "ім'я по батькові", oSheet.getCellByPosition(2,  row).String
    MapPut data, "термін",           oSheet.getCellByPosition(3,  row).String
    MapPut data, "виселення",        oSheet.getCellByPosition(4,  row).String
    MapPut data, "сплачено",         oSheet.getCellByPosition(5,  row).String
    MapPut data, "видаток",          oSheet.getCellByPosition(6,  row).String
    MapPut data, "прихід",           oSheet.getCellByPosition(7,  row).String
    MapPut data, "коментар",         oSheet.getCellByPosition(8,  row).String
    MapPut data, "телефон",          oSheet.getCellByPosition(9,  row).String
    MapPut data, "паспортні дані",   oSheet.getCellByPosition(10, row).String
    MapPut data, "дата народження",  oSheet.getCellByPosition(11, row).String  
    MapPut data, "чорний список",    oSheet.getCellByPosition(12, row).String
    MapPut data, "хостел",           oSheet.getCellByPosition(13, row).String
    MapPut data, "створено",         oSheet.getCellByPosition(14, row).String
    MapPut data, "причина зсуву",    oSheet.getCellByPosition(15, row).String
    MapPut data, "місце",            oSheet.getCellByPosition(16, row).String
    MapPut data, "історія",          oSheet.getCellByPosition(17, row).String
    MapPut data, "код",              oSheet.getCellByPosition(18, row).String
    MapPut data, "id",               oSheet.getCellByPosition(19, row).String
    MapPut data, "адмін",            oSheet.getCellByPosition(20, row).String            
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

    MsgBox sText, 64, "Дані з рядка"
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
    col  = oSel.RangeAddress.StartColumn
    
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
Function IsInAllowedRows(Optional sAddTitle As Variant) As Variant

    If IsMissing(sAddTitle) Or Len(Trim(sAddTitle)) = 0 Then sAddTitle = ""
    
    Dim lStartRow As Long
    Dim lEndRow   As Long
    Dim lRow      As Long
    Dim lValueD   As Long
    Dim sValueS   As String
    Dim sValueE   As String
    Dim aRange    As Variant
    Dim oDoc      As Object
    Dim oSel      As Object
    Dim oSheet    As Object

    oSheet    = ThisComponent.Sheets(0)
    oDoc      = ThisComponent
    oSel      = oDoc.CurrentSelection
    lRow      = oSel.RangeAddress.StartRow
    lValueD   = CLng(oSheet.getCellByPosition(18, lRow).getValue)
    sValueS   = Trim(oSheet.getCellByPosition(18, lRow).getValue)
    sValueE   = Trim(oSheet.getCellByPosition(4, lRow).getString) ' E — виселення   
    aRange    = GetAfterLastEncashRange()
    lStartRow = aRange(0)
    lEndRow   = aRange(1)
    
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
        IsInAllowedRows = Array(True, False)
        Exit Function
    End If
    
    Dim bOutOfRange As Boolean
    Dim bHasData    As Boolean
    
    bOutOfRange = (lStartRow = 0 And lEndRow = 0) Or (lRow < lStartRow)
    bHasData = lRow > 2 And sValueE <> ""

    ' якщо діапазон = [0,0], курсор >2, E не порожня, просимо пароль
    If bOutOfRange And bHasData Then
        If Not ShowNegetDialog(NEGET_RULES) Then
            MsgDlg "Відмова " & sAddTitle, "Операцію скасовано.", False, 55, 130
            IsInAllowedRows = Array(False, True)
            Exit Function
        End If

        IsInAllowedRows = Array(True, True)
        Exit Function
    End If
    
    IsInAllowedRows = Array(isBelowEncashRange(Array(lRow, lStartRow, lEndRow)), False)
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
    row  = oSel.RangeAddress.StartRow

    oCtrl = oDoc.CurrentController
    oCtrl.select(oDoc.Sheets(0).getCellByPosition(0, row))
End Sub

' =====================================================
' === Функція FormatHistoryLine =======================
' =====================================================
' → Приймає `data` (Map із полями рядка).
' → Формує рядок історії у вигляді: значення1 | значення2 | … | значенняN.
' → Використовується для збереження змін у полі "історія" (S).
Function FormatHistoryLine(data As Variant, oDialog As Object) As String
    Dim i      As Long
    Dim sLine  As String : sLine = ""
    Dim sAdmin As String
    Dim aParts As Variant

    ' ==== Отримуємо текст з мітки ====
    aParts = Split(oDialog.getControl("AdminLabel").getText(), ":")
    
    If UBound(aParts) >= 1 Then
        sAdmin = Trim(aParts(1))
    Else
        sAdmin = ""
    End If
        
    For i = LBound(data) To UBound(data)
        If i > LBound(data) Then sLine = sLine & " | "
        sLine = sLine & data(i).Value
    Next i
    
    sLine = sLine & " | " & sAdmin
    
    FormatHistoryLine = sLine
End Function

' =====================================================
' === Функція AppendHistory ===========================
' =====================================================
' → Зчитує всі поля поточного рядка та формує новий рядок історії.
' → Додає його на початок існуючої історії (якщо вона є).
' → Повертає `True`, якщо запис історії успішний.
' → Використовується під час редагування для відстеження змін.
Function AppendHistory(row As Long, oDialog As Object) As Boolean
    On Error GoTo ErrHandler
    Dim oSheet      As Object
    Dim data        As Variant
    Dim sOldHistory As String
    Dim sNewHistory As String
    Dim sLine       As String

    oSheet      = ThisComponent.Sheets(0)
    data        = FilterMapByKeys(ReadFromTable(), LIST_OF_HISTORY_FIELDS) ' тількі поля з LIST_OF_HISTORY_FIELDS
    sLine       = FormatHistoryLine(data, oDialog)
    sOldHistory = oSheet.getCellByPosition(17, row).String

    If Len(Trim(sOldHistory)) > 0 Then
        sNewHistory = sOldHistory & Chr(10) & sLine  
    Else
        sNewHistory = LIST_OF_HISTORY_FIELDS & Chr(10) & sLine
    End If

    oSheet.getCellByPosition(17, row).String = sNewHistory
    AppendHistory = True
    Exit Function

ErrHandler:
    MsgDlg "Помилка редагування.", "Помилка запису історії: " & Err.Description, False, 55
    AppendHistory = False
End Function

Function FilterMapByKeys(data As Variant, keys As String) As Variant
    Dim arrKeys() As String
    Dim filtered As Variant
    Dim k As String
    arrKeys = Split(keys, "|")
    filtered = CreateMap()
    
    Dim i As Long
    For i = LBound(arrKeys) To UBound(arrKeys)
        k = Trim(arrKeys(i))
        If k <> "" And MapHasKey(data, k) Then
            MapPut filtered, k, MapGet(data, k)
        End If
    Next i
    
    FilterMapByKeys = filtered
End Function

