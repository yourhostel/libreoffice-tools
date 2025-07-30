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
Function GetAdmins(Optional employee As Variant) As String
    Dim aAdmins As String
    Dim oDoc    As Object
    Dim oSheet  As Object
    Dim sPart   As String
    Dim sName   As String
    Dim sMiddle As String
    Dim p       As Long    
    Dim n       As Long
    Dim m       As Long    
    Dim sRole   As String
    
    oDoc    = ThisComponent
    aAdmins = ""
    
    If IsMissing(employee) Then
        sRole = ADMIN_MANAGER
    Else
        sRole = CStr(employee)
    End If
    
    Select Case sRole
        Case ADMIN_MANAGER
            p = 0 : n = 1 : m = 2
        Case TOP_MANAGER
            p = 5 : n = 6 : m = 7
        Case Else
            MsgDlg "Помилка", "Невідомий тип ролі: ", False, 50 
            Exit Function
    End Select
    
    On Error GoTo ErrorHandler
    oSheet = oDoc.Sheets.getByName("admins")
    
    ' ==== Збираємо всі непорожні значення зі стовпців A B C ====
    For iRow = 0 To MAX_SEARCH_RANGE_IN_ADMINS
       
        sPart   = oSheet.getCellByPosition(p, iRow).String         
        sName   = oSheet.getCellByPosition(n, iRow).String
        sMiddle = oSheet.getCellByPosition(m, iRow).String
        
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

Sub TestRole()
    MsgDlg "TestRole", GetAdmins(TOP_MANAGER), True, 100
End Sub

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

' =====================================================
' === Функція ShowPasswordDialog ======================
' =====================================================
' → Виводить діалог для введення пароля.
' → Порівнює введений пароль з очікуваним та повертає True, якщо вони збігаються.
Function ShowNegetDialog(sExpectedPassword As String) As Boolean

    Dim oDialog      As Object
    Dim oDialogModel As Object
    Dim oPass        As Object
    Dim oMngr        As Object
    Dim oButton      As Object
    Dim sPass        As String
    Dim sMngr        As String
    Dim bResult      As Boolean

    ' створюємо діалог і модель
    oDialog      = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    oDialogModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    
    oDialog.setModel(oDialogModel)

    With oDialogModel
        .PositionX = 100
        .PositionY = 100
        .Width     = 160
        .Height    = 90
        .Title     = "Введіть пароль"
    End With
    
    AddBackground(oDialogModel, BACKGROUND)
    
    ComboBoxTemplate oDialogModel,   "Top", "Manager:",  10, 15, MANAGER, 50, 140, GetAdmins(TOP_MANAGER)    
    FieldTemplate oDialogModel, "Password",  "Пароль:",  10, 45,      "", 50, 140   
    AddButton oDialogModel, "OkButton", "OK", 55, 70, 50, 14, 1 ' 1 = OK кнопка
    oDialog.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)
    oDialog.getControl("PasswordField").Model.EchoChar = Asc("*")

    ' виконуємо діалог
    If oDialog.execute() = 1 Then
        oPass  = oDialog.getControl("PasswordField")
        sPass  = oPass.getModel().Text
        oMngr  = oDialog.getControl("TopCombo")
        sMngr  = oMngr.getModel().Text
        
        bResult = (Obfuscate(sPass) = GetNeget(sMngr))
        
        If bResult Then 
            SetMngr(sMngr)
        Else 
            MsgDlg "Помилка", "Невірний пароль або ім’я менеджера", False, 50
        End If
        ' MsgDlg "ShowNegetDialog", bResult, False, 50
    End If

    oDialog.dispose()

    ShowNegetDialog = bResult
End Function

' =====================================================
' === Функція SetMngr ==================================
' =====================================================
' → Встановлює ім’я менеджера в клітинку K1 листа "admins".
' → Повертає True — якщо успішно, False — якщо виникла помилка (через MsgDlg).
' → Використовується для збереження вибраного менеджера системи.
Function SetMngr(sMngr As String) As Boolean  
    On Error GoTo Fail
    
    Dim oDoc      As Object : oDoc      = ThisComponent
    Dim oSheet    As Object : oSheet    = oDoc.Sheets.getByName("admins")
    Dim oCellMngr As Object : oCellMngr = oSheet.getCellByPosition(10, 0) ' K1
    oCellMngr.setString(sMngr)
    
    SetMngr = True
    Exit Function
    
Fail:
    MsgDlg "Помилка", "Не вдалося встановити менеджера", False, 50
    SetMngr = False     
End Function

' =====================================================
' === Функція GetNeget ================================
' =====================================================
' → Шукає менеджера за ПІБ на листі "admins" (F-H колонки).
' → Якщо знайдено — повертає значення з колонки I (правила NEGET).
' → Якщо не знайдено — повертає порожній рядок.
' → Перевіряється до 20 рядків (nRange = 20).
Function GetNeget(sMngr As String) As String
    Dim oDoc    As Object : oDoc    = ThisComponent
    Dim oSheet  As Object : oSheet  = oDoc.Sheets.getByName("admins")
    Dim nRange  As Long   : nRange  = 20
    Dim i       As Long
    Dim sFull   As String
    Dim sPart   As String
    Dim sName   As String
    Dim sMiddle As String
    
    For i = 0 To nRange
        sPart   = Trim(oSheet.getCellByPosition(5, i).String)   ' Прізвище
        sName   = Trim(oSheet.getCellByPosition(6, i).String)   ' Ім'я
        sMiddle = Trim(oSheet.getCellByPosition(7, i).String)   ' По батькові
        
        sFull = sPart & " " & sName & " " & sMiddle
        
        If sFull = sMngr Then
            GetNeget = oSheet.getCellByPosition(8, i).String
            Exit Function
        End If
    Next i

    GetNeget = "" ' Якщо не знайдено
End Function
