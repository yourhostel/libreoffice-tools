REM  *****  BASIC  *****

' Encashment.bas

' =====================================================
' === Процедура DoEncashment ==========================
' =====================================================
' → Запускає інкасацію: перевіряє пароль, знаходить діапазон, підраховує та вставляє запис.
' → Якщо помилка (порожня таблиця, немає записів для інкасації) — показує повідомлення й виходить.
Sub DoEncashment()   
    Dim oDocument         As Object
    Dim oSheet            As Object
    Dim aRange            As Variant
    Dim lStartRow         As Long
    Dim lEndRow           As Long
    Dim nTotalEncash      As Double
    
    ' === ініціалізація документа та аркуша ===
    oDocument = ThisComponent
    oSheet = oDocument.Sheets(0)
    
    ' === перевірка: таблиця порожня ===
    If Trim(oSheet.getCellByPosition(18, 3).String) = "" Then
        MsgDlg "Помилка", "Таблиця порожня. Інкасувати нема чого.", False, 50, 160
        Exit Sub
    End If
    
    ' === знаходимо діапазон для інкасації ===
    aRange = GetAfterLastEncashRange()
    lStartRow = aRange(0)
    lEndRow = aRange(1)

    ' === перевіряємо чи є що інкасувати ===
    If lStartRow = 0 And lEndRow = 0 Then
        MsgDlg "Інкасація не потрібна", "Після останньої інкасації немає нових записів.", False, 50, 160
        Exit Sub
    End If
    
    If Not ShowNegetDialog(NEGET_RULES) Then
        MsgDlg "Помилка", String(18, " ") & "Операцію скасовано.", False, 50, 130
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

    nSumPaid    = 0
    nSumExpense = 0
    nSumIncome  = 0
    
    For lRowNumber  = lStartRow To lEndRow
        nSumPaid    = nSumPaid + oSheet.getCellByPosition(5, lRowNumber).getValue()
        nSumExpense = nSumExpense + oSheet.getCellByPosition(6, lRowNumber).getValue()
        nSumIncome  = nSumIncome + oSheet.getCellByPosition(7, lRowNumber).getValue()
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
    Dim oCell            As Object
    Dim dToday           As Date
    Dim oSelectedCell    As Object
    Dim lRowNumber       As Long
    Dim oSheetAdm        As Object

    oDocument = ThisComponent
    oSheet = oDocument.Sheets(0)

    dToday = Now()
    
    ' === позиціонування на перший порожній рядок ===
    SelectFirstEmptyInA()
    
    oSelectedCell = oDocument.CurrentSelection
    lRowNumber = oSelectedCell.RangeAddress.StartRow

    ' запис
    oCell = oSheet.getCellByPosition(0, lRowNumber)
    oCell.setValue(Cdate(Format(dToday, "DD.MM.YYYY")))

    oSheet.getCellByPosition(18, lRowNumber).setValue(7)
    oSheet.getCellByPosition(4, lRowNumber).setString(ENCASH)

    oCell = oSheet.getCellByPosition(14, lRowNumber)
    oCell.setValue(Cdate(Format(dToday, "DD.MM.YYYY HH:MM:SS")))
    
    ' ==== Вставка id у колонку U ====
    SetNextId(oSheet, oSelectedCell)
    
    ' ==== Вставка hostel у колонку N ====
    HostelInsertion oSelectedCell
    
    oSheet.getCellByPosition(5, lRowNumber).setValue(nTotalEncash)
    
    oSheetAdm = oDocument.Sheets.getByName("admins")
    sAdmin    = oSheetAdm.getCellByPosition(4, 0).getString    ' E1 admins
    oSheet.getCellByPosition(20, lRowNumber).setString(sAdmin) ' U  data 

    MsgDlg "Інкасація виконана", "Сума інкасації: " & nTotalEncash, False, 50, 140
End Sub

' =====================================================
' === Функція GetAfterLastEncashRange =================
' =====================================================
' → Визначає діапазон рядків для інкасації.
' → Шукає останній запис інкасації (7) та порожній рядок у колонці D.
' → Повертає масив [lStartRow, lEndRow, lFinancialRow]
' → де lFinancialRow — рядок з 30 або (lEndRow+1), якщо 30 не знайдено.
Function GetAfterLastEncashRange() As Variant
    Dim oSheet        As Object
    Dim lStartRow     As Long
    Dim lEndRow       As Long
    Dim lFinancialRow As Long
    Dim lCheckRow     As Long
    Dim dVal          As Double
    Dim bFound30      As Boolean

    oSheet        = ThisComponent.Sheets(0)
    lStartRow     = -1
    lEndRow       = -1
    lFinancialRow = -1
    bFound30      = False
    
    ' === знаходимо останній заповнений рядок у колонці S ===
    For lCheckRow = 3 To oSheet.Rows.Count - 1
        dVal = oSheet.getCellByPosition(18, lCheckRow).getValue()
        If dVal = 0 Then
            lEndRow = lCheckRow - 1
            Exit For
        End If
    Next
    If lEndRow = -1 Then lEndRow = oSheet.Rows.Count - 1

    ' === знаходимо останній "7" від lEndRow вгору ===
    For lCheckRow = lEndRow To 3 Step -1
        dVal = oSheet.getCellByPosition(18, lCheckRow).getValue()
        If dVal = 7 Then
            lStartRow = lCheckRow + 1
            Exit For
        End If
    Next
    If lStartRow = -1 Then lStartRow = 3

    ' === якщо діапазон некоректний ===
    If lEndRow < lStartRow Then
        GetAfterLastEncashRange = Array(0, 0, lEndRow+1, bFound30)
        Exit Function
    End If

    ' === шукаємо "30" між lStartRow та lEndRow ===
    For lCheckRow = lStartRow To lEndRow
        dVal = oSheet.getCellByPosition(18, lCheckRow).getValue()
        If dVal = 30 Then
            lFinancialRow = lCheckRow
            bFound30 = True
            Exit For
        End If
    Next
    
    If lFinancialRow = -1 Then lFinancialRow = lEndRow + 1
    
    ' === повертаємо результат ===
    GetAfterLastEncashRange = Array(lStartRow, lEndRow, lFinancialRow, bFound30)
End Function

' =====================================================
' === Sub AddFinancialRow =============================
' → Додає рядок з кодом 30 між інкасаціями
' → у визначений діапазон після останнього 7 і до порожнього
' =====================================================
Sub AddFinancialRow()
    Dim aEncashRange()    As Variant
    Dim aRes()            As Variant
    Dim IFinRow           As Long 
    Dim lAmount           As Long
    Dim oDocument         As Object
    Dim oSheet            As Object
    Dim oSheetAdmins      As Object
    Dim oCode             As Object
    Dim oExpense          As Object
    Dim oIncome           As Object
    Dim oComment          As Object
    Dim oNameRow          As Object
    Dim oDate             As Object
    Dim oDateCreate       As Object
    Dim oAdmin            As Object        
    Dim bToggleAddingType As Boolean
    Dim sCurrentAdmin     As String
    Dim sComment          As String
    Dim sCommentCell      As String
    Dim sToggleFinType    As String
    Dim dToday            As Date
            
    dToday            = Now()
    aEncashRange      = GetAfterLastEncashRange()
    IFinRow           = aEncashRange(2)
    bToggleAddingType = aEncashRange(3)

    oDocument         = ThisComponent
    oSheet            = oDocument.Sheets.getByName("data")
    oSheetAdmins      = oDocument.Sheets.getByName("admins")
    sCurrentAdmin     = oSheetAdmins.getCellByPosition(3, 0).getString() ' admins D1
         
    oExpense          = oSheet.getCellByPosition(6, IFinRow)
    oIncome           = oSheet.getCellByPosition(7, IFinRow)
    oComment          = oSheet.getCellByPosition(8, IFinRow)
    oCode             = oSheet.getCellByPosition(18, IFinRow)
    oDate             = oSheet.getCellByPosition(0, IFinRow)
    oDateCreate       = oSheet.getCellByPosition(14, IFinRow)
    oNameRow          = oSheet.getCellByPosition(4, IFinRow)
    oAdmin            = oSheet.getCellByPosition(20, IFinRow)
    
    Do
        aRes = ShowFinDialog()
        If IsNull(aRes) Then
            MsgDlg "Скасовано", "Операція скасована користувачем", False, 50, 140
            Exit Sub
        End If
    Loop Until ValidationFinData(aRes)
    
    ' ==== Видаток, прихід, коментар G, H, I
    lAmount        = aRes(0)
    sComment       = aRes(1)
    sToggleFinType = aRes(2)
    
    Dim sCmt As String
    If bToggleAddingType Then 
        sCommentCell = oComment.getString()
        
        sCmt = sCommentCell & Chr(10) & _
               FormatFinLine(sToggleFinType,lAmount, sComment, sCurrentAdmin)
        oComment.setString(sCmt)
                       
        If sToggleFinType = "видаток" Then oExpense.setValue(lAmount + oExpense.getValue())      
        If sToggleFinType = "прихід" Then oIncome.setValue(lAmount + oIncome.getValue())
        MsgDlg "Баланс", "До записів додано " & sToggleFinType & ": " & lAmount, False, 50, 140      
    Else
        If sToggleFinType = "видаток" Then oExpense.setValue(lAmount)       
        If sToggleFinType = "прихід" Then oIncome.setValue(lAmount)
        sCmt = FormatFinLine(sToggleFinType,lAmount, sComment, sCurrentAdmin)
        
        oComment.setString(sCmt)
        oCode.setValue(30)
        oNameRow.setString(BALANCE)

        oDate.setValue(Cdate(Format(dToday, "DD.MM.YYYY")))   
        oDateCreate.setValue(Cdate(Format(dToday, "DD.MM.YYYY HH:MM:SS")))
        
        ' ==== Вставка id у колонку T ====
        SetNextId oSheet, oDate
        
        ' ==== Вставка hostel у колонку N ====
        HostelInsertion oDate
        
        ' ==== Вставка поточного адміна у колонку U ====
        oAdmin.setString(sCurrentAdmin)
        
        MsgDlg "Баланс", "Створено новий запис. До записів додано " & sToggleFinType & ": " & lAmount, False, 50, 200
    End If  
End Sub

' =====================================================
' === Функція FormatFinLine =====================
' → Повертає красиво вирівняний рядок:
'   видаток  |   100  | Коментар
'   прихід   |   50   | Коментар
' =====================================================
Function FormatFinLine(sType As String, _
                       amount As Double, _
                       sComment As String, _
                       sCurrentAdmin As String) As String
                       
    Dim sLine          As String
    Dim sAmount        As String
    Dim resultAmount   As String
    Dim colAmountWidth As Integer
    Dim cChar          As String
    
    cChar = "   " ' три пробіли
    colAmountWidth = 6

    ' === Перша колонка: sType ===
    If sType = "видаток" Then
        sLine = "видаток"
    ElseIf sType = "прихід" Then
        sLine = "прихід" & cChar
    Else
        sLine = Left(sType & Space(6), 6)
    End If

    ' === Друга колонка: amount з вирівнюванням через cChar ===
    sAmount = CStr(amount)
    resultAmount = ""
    For i = 1 To (colAmountWidth - Len(sAmount))
        resultAmount = resultAmount & cChar
    Next i
    resultAmount = resultAmount & sAmount

    ' === Об’єднання всіх частин ===
    sLine = Format(Now, "DD.MM.YYYY HH:MM") & _
        " | " & sLine & " | " & resultAmount & " | " & sCurrentAdmin & " | " & sComment

    FormatFinLine = sLine
End Function

Function ValidationFinData(aRes As Variant) As boolean
    Dim amount As Double, comment As String
    
    ValidationFinData = True
    ' парсимо
    amount            = Val(aRes(0))
    comment           = aRes(1)

    ' перевірка суми
    If amount <= 0 Then
        MsgDlg "Помилка", "Сума має бути додатним числом", False, 50, 140
        ValidationFinData = False
        Exit Function
    End If

    ' перевірка коментаря
    If Len(comment) < 3 Then
        MsgDlg "Помилка", "Додайте коментар не менше 3 символів", False, 50, 140
        ValidationFinData = False
        Exit Function
    End If
End Function

' =====================================================
' === Функція ShowFinDialog ===========================
' =====================================================
' → Показує діалог фінансового запису.
' → Повертає Array(Сума, Коментар, Тип) або Nothing.
Function ShowFinDialog() As Variant
    Dim oDlg      As Object
    Dim oDlgModel As Object
    Dim result    As Variant
    
    oDlg      = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    oDlgModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    oDlg.setModel(oDlgModel)

    ' ==== Налаштування діалогу ====
    With oDlgModel
        .Title = "Фінансовий запис"
        .Width = 220
        .Height = 120
        .PositionX = 100
        .PositionY = 100
    End With
    
    AddBackground(oDlgModel, BACKGROUND)

    ' ==== Поле для суми ====
    FieldTemplate       oDlgModel, "Amount", "Сума:", 10, 20, "", 40, 50

    ' ==== Поле для коментаря ====
    FieldTemplate       oDlgModel, "Comment", "Коментар:", 10, 55, "", 50, 150

    ' ==== Радіокнопки Видаток/Прихід ====
    OptionGroupTemplate oDlgModel, "FinType", "Видаток", "Прихід", 120, 20, 40, True

    ' ==== Кнопка ОК ====
    AddButton           oDlgModel, "OkButton", "Додати", 85, 90, 50, 14, 1

    ' ==== Показуємо діалог ====
    oDlg.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)
    If oDlg.execute() = 1 Then
        Dim amount As String, comment As String, ftype As String
        amount = Trim(oDlg.getControl("AmountField").getModel().Text)
        comment = Trim(oDlg.getControl("CommentField").getModel().Text)
        If oDlg.getControl("FinTypeLeft").State Then
            ftype = "видаток"
        ElseIf oDlg.getControl("FinTypeRight").State Then
            ftype = "прихід"
        Else
            ftype = "?"
        End If
        result = Array(amount, comment, ftype)
    Else
        result = Nothing
    End If

    oDlg.dispose()
    ShowFinDialog = result
End Function

