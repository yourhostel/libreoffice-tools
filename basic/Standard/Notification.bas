REM  *****  BASIC  *****

' Notification.bas

' =====================================================
' === Процедура ShowDialog ============================
' =====================================================
' → Створює невеликий діалог з заголовком та одним або двома повідомленнями.
' → Виводить повідомлення користувачу та чекає натискання кнопки OK.
' → Додаткове повідомлення (нижнє) є необов’язковим.
' → Автоматично підлаштовує висоту діалогу під кількість повідомлень.
Sub ShowDialog(Title As String, _
			   MessageTop As String, _
			   Optional MessageDown As Variant)

    Dim oDlg As Object, oDlgModel As Object
    Dim PosY_Button As Integer
    Dim Height_Dialog As Integer
    Dim CheckMessageDown As Boolean

    CheckMessageDown = IsMissing(MessageDown) Or Len(Trim(MessageDown)) = 0
    PosY_Button = 30
    Height_Dialog = 50

    If Not CheckMessageDown Then
    	PosY_Button = 40
    	Height_Dialog = 65
    End If

    oDlgModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    oDlg = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    oDlg.setModel(oDlgModel)

    With oDlgModel
        .Title = Title
        .Width = 220
        .Height = Height_Dialog
    End With

    ' ==== FieldTemplate тільки для мітки без поля ====
    ' NamePrefix
    ' LabelText
    ' PositionX
    ' PositionY
    ' vText
    ' WidthLabel
    ' WidthField
    ' ReadOnly
    Call FieldTemplate(oDlgModel, "lineTop", MessageTop, 10, 20, "", 210, 0, True)

    If Not CheckMessageDown Then
        Call FieldTemplate(oDlgModel, "lineDown", MessageDown, 10, 35, "", 210, 0, True)
    	PosX_Button = 50
    End If

    ' ==== Кнопка OK ====
    Call AddButton(oDlgModel, "OKButton", "OK", 90, PosY_Button, 40, 14, 1)

    oDlg.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)
    oDlg.execute()
    oDlg.dispose()
End Sub

' ===============================================
' === MsgDlg ====================================
' → Показує діалог із довгим текстом, що переноситься й має скрол
' ===============================================
Sub MsgDlg(Title As String, _
         Message As String, _
         bVScroll As Boolean, _
         Optional nHeight As Variant)

    Dim oDlg As Object, oDlgModel As Object   ' діалог і його модель

    If IsMissing(nHeight) Or Len(Trim(nHeight)) = 0 Then nHeight = 140

    ' створюємо модель діалогу
    oDlgModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    ' створюємо сам діалог
    oDlg = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    oDlg.setModel(oDlgModel)

    ' налаштовуємо розмір і заголовок діалогу
    With oDlgModel
        .Title = Title         ' заголовок вікна
        .Width = 220           ' ширина діалогу (в пікселях)
        .Height = nHeight          ' висота діалогу (в пікселях)
    End With

    ' ==== Створюємо багаторядкове текстове поле ====
    Dim oTextModel As Object   ' модель текстового поля
    ' правильний клас моделі: UnoControlEditModel
    oTextModel = oDlgModel.createInstance("com.sun.star.awt.UnoControlEditModel")
    With oTextModel
        .Name = "MessageBox"       ' ім'я елемента
        .MultiLine = True          ' багаторядковий режим
        .ReadOnly = True           ' тільки для читання
        .VScroll = bVScroll        ' вертикальний скрол
        .HScroll = False           ' горизонтальний скрол
        .Text = Message            ' текст повідомлення
        .Width = 200               ' ширина текстового поля
        .Height = nHeight - 40     ' висота текстового поля
        .PositionX = 10            ' відступ зліва
        .PositionY = 10            ' відступ зверху
    End With

    ' додаємо текстове поле на діалог
    oDlgModel.insertByName("MessageBox", oTextModel)

    ' ==== Додаємо кнопку OK ====
    ' Викликаємо підпроцедуру, яка додає кнопку (твоя функція AddButton)
    ' Аргументи: модель, ім'я, текст, X, Y, ширина, висота, тип кнопки
    AddButton(oDlgModel, "OKButton", "OK", 90, nHeight - 20, 40, 14, 1)

    ' створюємо peer (реалізацію вікна) і показуємо діалог
    oDlg.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)
    oDlg.execute()      ' чекає на взаємодію користувача
    oDlg.dispose()      ' закриває й звільняє ресурси
End Sub