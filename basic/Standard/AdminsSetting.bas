REM  *****  BASIC  *****

' AdminsSetting.bas

Dim FormResultAdmins As Boolean

' =====================================================
' === Функція StartAdmins =============================
' =====================================================
' → Запускає діалог для встановлення адміністратора.
' → Показує діалог із випадаючим списком адміністраторів.
' → Виконує попередню перевірку через ShowNegetDialog().
' → Повертає "OK" або "Cancel" залежно від дій користувача.
Function StartAdmins() As String
    Dim oDialog         As Object  
    Dim oButtonAdmins   As Object
    Dim oListenerAdmins As Object
    Dim sResult         As String
    
    If Not ShowNegetDialog(NEGET_RULES) Then
        MsgDlg "Помилка", String(18, " ") & "Операцію скасовано.", False, 50, 130
        Exit Function
    End If
        
    FormResultAdmins = False
    
    oDialog = CreateDlgAdmins()
    
    ' === Кнопка "Скасувати" ===
    oButtonAdmins    = oDialog.getControl("AdminsButton")
    
    ' === Обробник кнопки oButtonCancel ===
    oListenerAdmins = CreateUnoListener("AdminsButton_", "com.sun.star.awt.XActionListener")
    oButtonAdmins.addActionListener(oListenerAdmins)
    
    ' === Змінна результату ===
    sResult = "Cancel"
    
    ' === Запуск діалогу ===
    If oDialog.execute() = 1 Then    
        ' Натиснута кнопка (будь-яка) — тут можна перевіряти логіку            
        sResult = "OK"
    End If   
    
    If FormResultAdmins Then
        MsgDlg "Готово", String(20, " ") & "Дані збережено", False, 50, 120
	Else
    	MsgDlg "Скасовано", String(22, " ") & "Вихід без змін", False, 50, 120
	End If

    ' === Очищення ===
    oButtonAdmins.removeActionListener(oListenerAdmins)
    oDialog.dispose()

    StartAdmins = sResult 
End Function

' =====================================================
' === Подія AdminsButton_actionPerformed ==============
' =====================================================
' → Обробляє натискання кнопки "Застосувати".
' → Перевіряє, чи вибраний адмін дійсний (зі списку).
' → Записує вибраного адміністратора в таблицю.
' → Закриває діалог, якщо все пройшло успішно.
Sub AdminsButton_actionPerformed(oEvent As Object)
    Dim oDialog As Object
    Dim sAdmin  As String 
    Dim aAdmins As Variant : aAdmins = Split(GetAdmins(), ";")
    
    oDialog = oEvent.Source.getContext()
    sAdmin  = oDialog.getControl("AdminsCombo").getText()
    
    If Not IsValidComboList(aAdmins, sAdmin) Then Exit Sub 
    If Not SetAdmin(sAdmin) Then Exit Sub
    
    FormResultAdmins = True          
    oDialog.endExecute()
End Sub

' =====================================================
' === Функція SetAdmin ================================
' =====================================================
' → Записує вибраного адміністратора в комірку D1 на аркуші "admins".
' → Повертає True, якщо вдалося, або False з повідомленням про помилку.
Function SetAdmin(sAdmin As String) As Boolean
    Dim oDoc    As Object
    Dim oSheet  As Object
    SetAdmin = False
    
    On Error GoTo ErrorHandler
    oDoc   = ThisComponent
    oSheet = oDoc.Sheets.getByName("admins")
    oSheet.getCellByPosition(3, 0).setString(sAdmin) ' D1
    
    SetAdmin = True
    Exit Function

ErrorHandler:
    MsgDlg "Помилка", "Не вдалося застосувати адміністратора: " & sAdmin, False, 50    
End Function    

' =====================================================
' === Функція CreateDlgAdmins =========================
' =====================================================
' → Створює діалог для вибору адміністратора.
' → Додає ComboBox з адміністраторами та кнопку підтвердження.
' → Повертає об'єкт створеного діалогу.
Function CreateDlgAdmins()    
    Dim oDialog      As Object
    Dim oDialogModel As Object
    Dim sAdmins      As String : sAdmins = GetAdmins()
    
    oDialog      = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    oDialogModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
      
    oDialog.setModel(oDialogModel)
    
    ' ==== Параметри діалогу ====    
    With oDialogModel
        .PositionX = 100
        .PositionY = 100
        .Width     = 200
        .Height    = 70
        .Title     = "Встановлення адміністратора"
    End With
    
    Dim gX As Long, gY As Long
    gX = 20 : gY = 15
    
    AddBackground(oDialogModel, BACKGROUND)    
    ComboBoxTemplate(oDialogModel, "Admins",     "Обрати:",      gx,      gY, "Необрано", 70,  160, sAdmins)
    AddButton(oDialogModel,  "AdminsButton", "Застосувати", 50 + gx, 30 + gY, 60, 14)
    
    oDialog.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)
       
    CreateDlgAdmins = oDialog
End Function

' =====================================================
' === Функція GetAdmins ===============================
' =====================================================
' → Зчитує список адміністраторів із аркуша "admins".
' → Читає дані зі стовпців A, B, C (Прізвище, Ім’я, По батькові).
' → Формує список у форматі: "ПІБ;ПІБ;...".
' → Повертає один рядок зі значеннями, розділеними крапкою з комою.
Function GetAdmins() As String
    Dim aAdmins As String
    Dim oDoc    As Object
    Dim oSheet  As Object
    Dim sName   As String
    Dim sPart   As String
    Dim sMiddle As String
    
    oDoc    = ThisComponent
    aAdmins = ""
    
    On Error GoTo ErrorHandler
    oSheet = oDoc.Sheets.getByName("admins")
    
    ' ==== Збираємо всі непорожні значення зі стовпців A B C ====
    For iRow = 0 To MAX_SEARCH_RANGE_IN_ADMINS
       
        sPart  = oSheet.getCellByPosition(0, iRow).String         
        sName  = oSheet.getCellByPosition(1, iRow).String
        sMiddle = oSheet.getCellByPosition(2, iRow).String
        
        If sPart <> "" And sName <> "" And sMiddle <> "" Then
            If aAdmins <> "" Then aAdmins = aAdmins & ";"
            aAdmins = aAdmins & sPart & " " & sName & " " & sMiddle
        End If
    Next iRow
    
    GetAdmins = aAdmins
    
    Exit Function

ErrorHandler:
    MsgDlg "Помилка", "Не вдалося завантажити список адміністраторів з аркуша 'admins'", False, 50
End Function

' =====================================================
' === Функція IsValidComboList ========================
' =====================================================
' → Перевіряє, чи заданий рядок є одним із елементів списку.
' → Використовується для валідації введення в ComboBox.
' → Якщо значення не знайдено — показує повідомлення про помилку.
' → Повертає True або False.
Function IsValidComboList(aValid As Variant, sTest As String)   
    For i = 0 To UBound(aValid)
        If Trim(sTest) = Trim(aValid(i)) Then
            IsValidComboList = True
            Exit Function
        End If
    Next i
    
    MsgDlg "Помилка", "Обране значення ('" & sTest & "') не з переліку!", False, 50
    IsValidComboList = False
End Function    
