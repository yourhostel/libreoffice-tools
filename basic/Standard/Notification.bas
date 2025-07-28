REM  *****  BASIC  *****

' Notification.bas

' ===============================================
' === MsgDlg ====================================
' → Показує діалог із довгим текстом, що переноситься й має скрол
' ===============================================
Sub MsgDlg(sTitle   As String, _
           Message  As String, _
           bVScroll As Boolean, _
  Optional nHeight  As Variant, _
  Optional nWidth   As Variant)
         
    Dim oDlg      As Object
    Dim oDlgModel As Object   ' діалог і його модель
    
    If IsMissing(nHeight) Or Len(Trim(nHeight)) = 0 Then nHeight = 140
    If IsMissing(nWidth) Or Len(Trim(nWidth)) = 0 Then nWidth = 220
    
    ' створюємо модель діалогу
    oDlgModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    ' створюємо сам діалог
    oDlg = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    oDlg.setModel(oDlgModel)

    ' налаштовуємо розмір і заголовок діалогу
    With oDlgModel
        .Title = sTitle            ' заголовок вікна
        .Width = nWidth            ' ширина діалогу (в пікселях)
        .Height = nHeight          ' висота діалогу (в пікселях)
        '.BackgroundColor = RGB(0, 0, 170)
    End With

    AddBackground(oDlgModel, BACKGROUND)
    
    ' ==== Створюємо багаторядкове текстове поле ====
    Dim oTextModel As Object   ' модель текстового поля
    ' клас моделі: UnoControlEditModel
    oTextModel = oDlgModel.createInstance("com.sun.star.awt.UnoControlEditModel")
    With oTextModel
        .Name      = "MessageBox"  ' ім'я елемента
        .MultiLine = True          ' багаторядковий режим
        .ReadOnly  = True          ' тільки для читання
        .VScroll   = bVScroll      ' вертикальний скрол
        .HScroll   = False         ' горизонтальний скрол
        .Text      = Message       ' текст повідомлення
        .Width     = nWidth - 20   ' ширина текстового поля
        .Height    = nHeight - 40  ' висота текстового поля
        .PositionX = 10            ' відступ зліва
        .PositionY = 10            ' відступ зверху
        .TextColor = RGB(255, 255, 255)
        .BackgroundColor = RGB(22, 11, 172)
    End With
    
    

    ' додаємо текстове поле на діалог
    oDlgModel.insertByName("MessageBox", oTextModel)
    
    'Dim oLbl
    'oLbl = oDlgModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel")

    'With oLbl
        '.Name = "MyLabel"
        '.Label = Message
        '.PositionX = 10
        '.PositionY = 10
        '.Width = nWidth - 20
        '.Height = nHeight - 40
        '.TextColor = RGB(255,255,255) ' Білий шрифт
    'End With

    'oDlgModel.insertByName(oLbl.Name, oLbl)
    

    ' ==== Додаємо кнопку OK ====
    ' Викликаємо підпроцедуру, яка додає кнопку (твоя функція AddButton)
    ' Аргументи: модель, ім'я, текст, X, Y, ширина, висота, тип кнопки
    AddButton(oDlgModel, "OKButton", "OK", (nWidth - 40) / 2, nHeight - 20, 40, 14, 1)

    ' створюємо peer (реалізацію вікна) і показуємо діалог
    oDlg.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)
    oDlg.execute()      ' чекає на взаємодію користувача
    oDlg.dispose()      ' закриває й звільняє ресурси
End Sub

