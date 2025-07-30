REM  *****  BASIC  *****

' DeleteRecord.bas

' =====================================================
' === Процедура DeleteRow =============================
' =====================================================
' → Видаляє поточний рядок після запиту пароля
' → За бажанням — переносить курсор у стовпець A
Sub DeleteRow()
    Dim oDoc          As Object
    Dim oSheet        As Object
    Dim oSel          As Object
    Dim row           As Long
    Dim bWasProtected As Boolean

    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection
    row = oSel.RangeAddress.StartRow
    oSheet = oDoc.Sheets(0)

    If Not ShowNegetDialog(NEGET_RULES) Then
        MsgDlg "Помилка", String(18, " ") & "Операцію скасовано.", False, 50, 130
        Exit Sub
    End If
    
    ' === Перевіряємо, чи захищений аркуш ===
    bWasProtected = oSheet.IsProtected
    
    ' ==== Якщо захищений — знімаємо захист ====
    If bWasProtected Then
        oSheet.unprotect(Deobfuscate(NEGET_RULES))
    End If
    
    ' === Видаляємо рядок ===
    oSheet.Rows.removeByIndex(row, 1)
    
    ' ==== Повертаємо захист назад ====
    If bWasProtected Then
        oSheet.protect(Deobfuscate(NEGET_RULES))
    End If

    ' === Переносимо курсор у стовпець ===
    ' MoveCursorToColumnA()
End Sub

Dim FormResultCancel As Boolean

' =====================================================
' === Функція CancelRowSwitch =========================
' =====================================================
' → Основна точка входу для скасування бронювання.
' → Перевіряє код на допустимість скасування.
' → Якщо код — один із скасувальних (20–23), виконує відкочення.
' → Інакше показує діалог скасування, в якому користувач обирає тип.
' → За результатами записує новий код і коментар.
' → Повертає "OK" або "Cancel" залежно від дій користувача.
Function CancelRowSwitch() As String
    Dim aSet            As Variant : aSet     = GetSheetSetting()
    Dim oSheet          As Object  : oSheet   = aSet(0)
    Dim row             As Long    : row      = aSet(1)
    Dim sId             As String  : sId      = oSheet.getCellByPosition(19, row).String   
    Dim oCode           As Object  : oCode    = oSheet.getCellByPosition(18, row)
    Dim nCode           As Long    : nCode    = oCode.getValue()
    Dim oComment        As Object  : oComment = oSheet.getCellByPosition(8, row)
    Dim aCodCom         As Variant
    Dim nCodeOld        As Long
    Dim sCommentOld     As String
    Dim oDialog         As Object  
    Dim oButtonCancel   As Object
    Dim oListenerCancel As Object
    Dim sResult         As String
    
    If Not CheckCode("ID " & sId, " він не може бути скасован.", CLng(nCode)) Then 
        Exit Function
    End If
    
    If Not ShowNegetDialog(NEGET_RULES) Then
        MsgDlg "Помилка", String(18, " ") & "Операцію скасовано.", False, 50, 130
        Exit Function
    End If            
          
    If Not CheckOccupiedPlace(oDialog, ACTION_CHECK_CANCEL) Then
        Exit Function
    End If
       
    Select Case nCode
        Case 20, 21, 22, 23
            aCodCom     = ParseCodeAndComment(oComment.String)
            nCodeOld    = aCodCom(0)
            sCommentOld = aCodCom(1)
            
            If nCodeOld = -1 Then
                MsgDlg "Помилка" , "Невдача парсингу рядка коментаря", False, 50 
                Exit Function
            End If

           oComment.setString(sCommentOld)
           oCode.setValue(nCodeOld)
           Exit Function
    End Select
    
    FormResultCancel = False
    
    oDialog = CreateDlgCancel()
    
    ' === Кнопка "Скасувати" ===
    oButtonCancel    = oDialog.getControl("CancelButton")
    
    ' === Обробник кнопки oButtonCancel ===
    oListenerCancel = CreateUnoListener("CancelButton_", "com.sun.star.awt.XActionListener")
    oButtonCancel.addActionListener(oListenerCancel)
    
    ' === Змінна результату ===
    sResult = "Cancel"   

    ' === Запуск діалогу ===
    If oDialog.execute() = 1 Then    
        ' Натиснута кнопка (будь-яка) — тут можна перевіряти логіку            
        sResult = "OK"
    End If
    
    If FormResultCancel Then
        MsgDlg "Готово", String(20, " ") & "Дані збережено", False, 50, 120
	Else
    	MsgDlg "Скасовано", String(22, " ") & "Вихід без змін", False, 50, 120
	End If

    ' === Очищення ===
    oButtonCancel.removeActionListener(oListenerCancel)
    oDialog.dispose()

    CancelRowSwitch = sResult 
End Function

' =====================================================
' === Процедура CancelButton_actionPerformed ==========
' =====================================================
' → Обробник кнопки "Скасувати" у діалозі Cancel.
' → Зчитує причину і коментар з діалогу.
' → Формує новий рядок коментаря і додає його до попереднього.
' → Залежно від причини встановлює відповідний код (20–23).
' → Ставить прапорець FormResultCancel = True для підтвердження.
' → Закриває діалог через endExecute.
Sub CancelButton_actionPerformed(oEvent As Object)
    Dim oDialog As Object
    
    oDialog = oEvent.Source.getContext()
    
    Dim aSet      As Variant  : aSet        = GetSheetSetting()
    Dim oSheet    As Object   : oSheet      = aSet(0)
    Dim row       As Long     : row         = aSet(1)
    Dim oComment  As Object   : oComment    = oSheet.getCellByPosition(8, row)
    Dim oCode     As Object   : oCode       = oSheet.getCellByPosition(18, row)
    Dim sCode     As String   : sCode       = oCode.String
    Dim sComment  As String   : sComment    = oComment.String

    Dim sReason     As String : sReason     = Trim(oDialog.getControl("CancelCombo").getText())
    Dim sCommentDlg As String : sCommentDlg = oDialog.getControl("CommentField").getText()
    Dim StrLast     As String
    
    If sComment = "" Then
        StrLast = "Код | " & sCode & " | " & sReason & " | " & sCommentDlg
    Else
        StrLast = sComment & Chr(10) & "Код | " & sCode & " | " & sReason & " | " & sCommentDlg 
    End If
          
    oComment.setString(StrLast)
    
    Select Case sReason
        Case "Скасування"
            oCode.setValue(20)
                        
        Case "Пауза"
            oCode.setValue(21)
                       
        Case "Перенесення"
            oCode.setValue(22)
                       
        Case "Часткова оплата"
        
            oCode.setValue(23)             
    End Select
    
    FormResultCancel = True          
    oDialog.endExecute()
End Sub

' =====================================================
' === Функція CreateDlgCancel =========================
' =====================================================
' → Створює діалог для скасування (ComboBox з причиною та поле коментаря).
' → Встановлює заголовок відповідно до імені і частини.
' → Повертає ініціалізований об'єкт діалогу для подальшої роботи.
Function CreateDlgCancel()
    Dim oDialog        As Object
    Dim oDialogModel   As Object
    
    Dim aSet    As Variant : aSet   = GetSheetSetting()
    Dim oSheet  As Object  : oSheet = aSet(0)
    Dim row     As Long    : row    = aSet(1)
   
    Dim sName   As String  : sName  = oSheet.getCellByPosition(1, row).String
    Dim sPart   As String  : sPart  = oSheet.getCellByPosition(2, row).String
    
    oDialog      = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    oDialogModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
      
    oDialog.setModel(oDialogModel)
    
    ' ==== Параметри діалогу ====    
    With oDialogModel
        .PositionX = 100
        .PositionY = 100
        .Width     = 200
        .Height    = 100
        .Title     = sName & " " & sPart
    End With
    
    AddBackground(oDialogModel, BACKGROUND)
       
    ' ==== Групова рамка ====
    Dim gX As Long, gY As Long
    gX = 20 : gY = 15
    
    ComboBoxTemplate(oDialogModel, "Cancel",  "Причина:",      gx,      gY, "Скасування", 70,  160, LIST_OF_CANCELLATIONS)
    FieldTemplate(oDialogModel,   "Comment", "Коментар:",      gx, 30 + gY,           "", 70,  160)    
    AddButton(oDialogModel,  "CancelButton", "Скасувати", 50 + gx, 60 + gY, 60, 14)
    
    oDialog.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)
       
    CreateDlgCancel = oDialog
End Function

' =====================================================
' === Функція GetSheetSetting =========================
' =====================================================
' → Повертає масив з активним листом, рядком і виділенням.
' → Формат повернення: Array(oSheet, row, oSelection)
' → Використовується для стандартизованого доступу до активного контексту.
Function GetSheetSetting() As Variant
    Dim oDoc    As Object : oDoc   = ThisComponent
    Dim oSel    As Object : oSel   = oDoc.CurrentSelection
    Dim row     As Long   : row    = oSel.RangeAddress.StartRow
    Dim oSheet  As Object : oSheet = oDoc.Sheets(0)
    GetSheetSetting = Array(oSheet, row, oSel)
End Function

' =====================================================
' === Функція ParseCodeAndComment =====================
' =====================================================
' → Парсить останній рядок коментаря виду "Код | ...".
' → Витягує числовий код із другої позиції (index 1).
' → Склеює назад усі попередні рядки — як старий коментар.
' → Повертає масив: [код, старийКоментар]
' → Якщо парсинг не вдається, повертає [-1, оригінальний текст]
Function ParseCodeAndComment(sText As String) As Variant
    Dim aLines     As Variant
    Dim sLastLine  As String
    Dim aParts     As Variant
    Dim sCodePart  As String
    Dim sRest      As String
    Dim i          As Long
    
    ' Значення за замовчуванням
    ParseCodeAndComment = Array(-1, sText)
    
    aLines = Split(sText, Chr(10))
    If Not IsArray(aLines) Or UBound(aLines) < 0 Then Exit Function

    sLastLine = Trim(aLines(UBound(aLines)))
    If sLastLine = "" Then Exit Function
    
    aParts = Split(sLastLine, " | ")
    If Not IsArray(aParts) Or UBound(aParts) < 1 Then Exit Function

    sCodePart = Trim(aParts(1))
    If Not IsNumeric(sCodePart) Then Exit Function

    ' Склеюємо назад усі рядки крім останнього
    sRest = ""
    For i = 0 To UBound(aLines) - 1
        If sRest <> "" Then sRest = sRest & Chr(10)
        sRest = sRest & aLines(i)
    Next i

    ' Повертаємо код та старий коментар
    ParseCodeAndComment = Array(CLng(sCodePart), sRest)
End Function
