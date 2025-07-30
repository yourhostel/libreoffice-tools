REM  *****  BASIC  *****

' Blacklist.bas

' =====================================================
' === Sub BlacklistStart ==============================
' =====================================================
' → Головна точка входу.
' → Визначає контекст: якщо натиснули на шапку — перемикає фільтр чорного списку.
' → Якщо виділений рядок — перевіряє діапазон, чорний список, правильність значень і викликає AddToBlacklist.
Sub BlacklistStart()
    Dim oSel              As Object
    Dim oSheet            As Object
    Dim nRow              As Long
    Dim sVal              As String
    Dim nVal              As Long
    Dim bInvalid          As Boolean
    Dim sCurrentCellValue As String
    Dim sM                As String
    Dim nFirstRow         As Long
    Dim nLastRow          As Long
    
    ' отримуємо поточну виділену клітинку
    oSel = ThisComponent.CurrentSelection
    
    If oSel.supportsService("com.sun.star.sheet.SheetCell") Then
        ' Це одна комірка
        nRow = oSel.CellAddress.Row
        nCol = oSel.CellAddress.Column
        oSheet = oSel.Spreadsheet
    Else
        ShowBlacklistInstructions
        Exit Sub
    End If

    ' якщо взагалі щось не те — теж показати інструкцію
    If nCol < 0 Or nRow < 0 Then
        ShowBlacklistInstructions
        Exit Sub
    End If    
    
    ' ==== Якщо захищений — знімаємо захист ====
    Dim bWasProtected As Boolean
    bWasProtected = oSheet.isProtected()
    If bWasProtected Then oSheet.unprotect(Deobfuscate(NEGET_RULES))
    
    ' отримуємо діапазон записів
    oRange    = GetRecordsRange()
    nFirstRow = oRange.RangeAddress.StartRow
    nLastRow  = oRange.RangeAddress.EndRow
    
    ' Застосовуємо фільтр якщо вибрано назву шапки 
    If Trim(oSel.getString()) = "чорний список" And nFirstRow - 1 = 2 Then
        ResetFilter(False)
        FilterBlacklist()      
        GoTo Cleanup
    End If
    
    ' перевіряємо, чи курсор у діапазоні
    If nRow < nFirstRow Or nRow > nLastRow Then
        MsgDlg "Помилка", "Виділений рядок поза діапазоном записів (" & (nFirstRow+1) & "–" & (nLastRow+1) & ")", False, 65
        ShowBlacklistInstructions
        GoTo Cleanup
    End If
    
    ' перевіряємо колонку M — якщо вже є коментар, повертаємо код з коментаря у колонку D, видаляємо коментар й виходимо
    sM = Trim(oSheet.getCellByPosition(12, nRow).getString())
    If sM <> "" Then
        If RemoveFromBlacklist(oSheet, nRow) Then GoTo Cleanup
    End If
    
    ' читаємо значення з колонки D
    sVal = Trim(oSheet.getCellByPosition(18, nRow).getString()) 
    If IsNumeric(sVal) Then
        nVal = Val(sVal)
    Else
        nVal = -1 ' гарантовано невірне
    End If

    ' перевіряємо одразу тут
    bInvalid = Not IsNumeric(sVal) Or nVal < 1 Or nVal > 19 Or nVal = 7

    If bInvalid Then
        MsgDlg "Помилка вибору запису", _
               "Запис має тримати Людину яка проживала/проживає", False, 65
        ShowBlacklistInstructions 
        GoTo Cleanup       
    End If

    ' якщо все ок — викликаємо AddToBlacklist
    AddToBlacklist(oSel)
    
    ' ==== Повертаємо захист назад ====
Cleanup:
    If bWasProtected Then 
        oSheet.protect(Deobfuscate(NEGET_RULES))
    End If
End Sub

' =====================================================
' === Процедура ShowBlacklistInstructions =============
' =====================================================
' → Виводить інструкцію користувачу щодо роботи з чорним списком.
' → Містить покрокові підказки:
'     1. Вибір однієї клітинки.
'     2. Додавання запису до ч/с.
'     3. Застосування/скидання фільтру через клітинку M3.
Sub ShowBlacklistInstructions()
    Dim sMsg As String
    sMsg = "1. Виберіть лише одну клітинку!" & Chr(10) & _
           Chr(10) & _
           "3. Для додавання до ч/с оберіть рядок з людиною." & Chr(10) & _
           Chr(10) & _
           "4. Щоб застосувати фільтр — оберіть клітинку M3." & Chr(10) & _
           "(чорний список у шапці таблиці)" & Chr(10) & _
           Chr(10) & _
           "5. Прибрати фільтр можна основною кнопкою 'Скинути' ""Скин."""           
    MsgDlg "Інструкція", sMsg, False, 115
End Sub

' =====================================================
' === Sub AddToBlacklist ==============================
' =====================================================
' → Виводить діалог із полем для коментаря.
' → Записує коментар у колонку M (12‑та) для поточного рядка.
' → Якщо рядок уже в чорному списку — видаляє коментар.
' → Виводить повідомлення про результат (додано/скасовано).
Sub AddToBlacklist(oSel As Object)
    Dim oSheet       As Object
    Dim oDialog      As Object
    Dim oDialogModel As Object
    Dim sComment     As String
    Dim nRow         As Long
    Dim nResult      As Long
    Dim sSurname     As String
    
    ' === Отримуємо аркуш і рядок ===
    oSheet = oSel.Spreadsheet
    nRow   = oSel.CellAddress.Row
    
    ' === Читаємо прізвище з колонки B ===
    sSurname    = Trim(oSheet.getCellByPosition(1, nRow).getString())
    sPatronymic = Trim(oSheet.getCellByPosition(2, nRow).getString())

    ' === Створюємо діалог ===
    oDialog = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    oDialogModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    oDialog.setModel(oDialogModel)

    With oDialogModel
        .PositionX = 100
        .PositionY = 100
        .Width     = 200
        .Height    = 70
        .Title     = sSurname & " " & sPatronymic 
    End With
    
    AddBackground(oDialogModel, BACKGROUND)

    ' === Поле введення коментаря ===
    FieldTemplate oDialogModel, "Comment", "Причина додавання в чорний список:" , 10, 20, "", 180, 180
    AddButton oDialogModel, "OkButton", "Додати", 75, 50, 50, 14, 1

    ' === Показуємо діалог ===
    oDialog.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)

    If oDialog.execute() <> 1 Then
        MsgDlg sSurname & " " & sPatronymic, "Скасовано. Запис до чорного списку не додано", False, 50
        oDialog.dispose()
        Exit Sub
    End If

    ' === Отримуємо коментар ===
    sComment = oDialog.getControl("CommentField").getModel().Text
    oDialog.dispose()
    
    ' перевіряємо довжину коментаря
    If Len(sComment) < 5 Then
        MsgDlg "Помилка", sSurname & " " & sPatronymic & " не додано до чорного списку" & Chr(10) & _
               Chr(10) & _
               "Коментар має бути не менше 5 символів!", False, 65
        Exit Sub
    End If
    
    Dim nCode As Long
    nCode = oSheet.getCellByPosition(18, nRow).getString()

    ' === Записуємо коментар у колонку M (12‑та) ===
    oSheet.getCellByPosition(12, nRow).setString(FormatBlacklist(nCode, sComment))
    oSheet.getCellByPosition(18, nRow).setValue(28)
    
    ResetFilterlimited
    
    MsgDlg "Додано", sSurname & " " & sPatronymic & " додано до чорного списку", False, 50
End Sub

' =====================================================
' === Функція FormatBlacklist =========================
' =====================================================
' → Формує текстовий запис для чорного списку.
' → Витягує ПІБ адміністратора з комірки D1 аркуша "admins".
' → Повертає рядок у форматі:
'     Код | {код} | {дата/час} | {адміністратор} | {коментар}
Function FormatBlacklist(nCode As Long, sComment) As String
    Dim oDoc      As Object : oDoc = ThisComponent
    Dim oSheetAdm As Object : oSheetAdm = oDoc.Sheets.getByName("admins")
    Dim sAdm      As String : sAdm = oSheetAdm.getCellByPosition(3, 0).getString()
    
    FormatBlacklist = "Код | " & nCode & " | " & Format(Now, "DD.MM.YYYY HH:MM") & _
        " | " & sAdm & " | " & sComment     
End Function

' =====================================================
' === Function RemoveFromBlacklist ====================
' =====================================================
' → Перевіряє комірку у колонці M на наявність коментаря чорного списку.
' → Якщо коментар знайдено та розібрано успішно —
'     переносить код назад у колонку D, видаляє коментар та показує повідомлення.
' → Повертає True, якщо запис успішно видалено з чорного списку;
'     False, якщо сталася помилка або коментар некоректний.
Function RemoveFromBlacklist(oSheet As Object, nRow As Long) As Boolean
    Dim oBlackListCell As Object
    Dim oSurnameCell   As Object
    Dim oCodeCell      As Object
    Dim parts()        As String
    Dim nParsedCode    As Long
    
    If Not CheckOccupiedPlace(Nothing, ACTION_CHECK_ROW) Then
        MsgDlg "Помилка", "Видалення з чорного списку неможливо." & Chr(10) & _
                          "Місце було зайняте.", False, 65
        RemoveFromBlacklist = True                        
        Exit Function
    End If
    
    oBlackListCell = oSheet.getCellByPosition(12, nRow)
    oSurnameCell   = oSheet.getCellByPosition(1, nRow)
    oPatronymic    = oSheet.getCellByPosition(2, nRow)
    oCodeCell      = oSheet.getCellByPosition(18, nRow)

    parts = Split(Trim(oBlackListCell.getString()), "|")

    If UBound(parts) >= 1 And IsNumeric(Trim(parts(1))) Then
        nParsedCode = CLng(Trim(parts(1)))
    Else
        MsgDlg "Помилка", "Не вдалося розібрати значення комірки «чорний список»!", False, 50
        RemoveFromBlacklist = False
        Exit Function
    End If

    oCodeCell.setValue(nParsedCode)
    oBlackListCell.setString("")

    MsgDlg "Видалено", oSurnameCell.getString() & " " & oPatronymic.getString() & " видалено із чорного списку."  & Chr(10) & _
    "" & Chr(10) & _
    "Код(" & nParsedCode & ") прайсу в колонку «код» повернуто успішно.", False, 65
    
    ResetFilterlimited
    
    RemoveFromBlacklist = True
End Function

' =====================================================
' === Sub FilterBlacklist =============================
' =====================================================
' → Застосовує фільтр для діапазону, щоб залишити тільки ті рядки, де колонка M непорожня.
' → Показує повідомлення, що фільтр застосовано.
Sub FilterBlacklist()
    Dim oRange           As Object
    Dim oFilterDesc      As Object
    Dim oFilterFields(0) As New com.sun.star.sheet.TableFilterField

    ' ==== Отримання діапазону даних ====
    Set oRange = GetRecordsRange()
    
    ' ==== Створення дескриптора фільтру ====
    Set oFilterDesc = oRange.createFilterDescriptor(True)

    ' ==== Оголошення фільтруючих полів (1 умова) ====
    With oFilterFields(0)
        .Field = 12        ' Колонка M (12‑та, індексація з 0)
        .Operator = com.sun.star.sheet.FilterOperator.NOT_EMPTY
        .IsNumeric = False
    End With

    ' ==== Призначаємо поля фільтру ====
    oFilterDesc.setFilterFields(oFilterFields())

    ' ==== Застосовуємо фільтр ====
    oRange.filter(oFilterDesc)

    MsgDlg "Фільтр застосовано", "Показано лише рядки з чорним списком", False, 50
End Sub

