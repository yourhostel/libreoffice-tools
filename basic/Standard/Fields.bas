REM  *****  BASIC  *****

' Fields.bas

' =====================================================
' === Процедура AddLogo ===============================
' =====================================================
' → Додає до діалогу картинку (логотип) за вказаними координатами та розмірами.
' → Підтягує файл з диска й вставляє у форму як ImageControl.
' → Масштабує зображення, щоб воно вмістилося у заданих розмірах.
Sub AddLogo(oDialogModel As Object, _
            sName        As String, _
            PositionX    As Integer, _
            PositionY    As Integer, _
            Width        As Integer, _
            Height       As Integer)

    Dim oImage As Object
    oImage            = oDialogModel.createInstance("com.sun.star.awt.UnoControlImageControlModel")
    oImage.Name       = sName
    oImage.PositionX  = PositionX
    oImage.PositionY  = PositionY
    oImage.Width      = Width
    oImage.Height     = Height
    oImage.ScaleImage = True
    oImage.ImageURL   = ConvertToURL(PATH_TO_LOGO)
    oDialogModel.insertByName(sName, oImage)
End Sub

' =====================================================
' === Процедура FieldTemplate =========================
' =====================================================
' → Додає на діалог мітку та поле введення (Edit) під нею.
' → Параметри:
'   — NamePrefix: префікс для імені (Label та Field)
'   — LabelText: текст мітки
'   — PositionX, PositionY: координати верхнього лівого кута поля
'   — vText: значення за замовчуванням
'   — WidthLabel: ширина мітки
'   — WidthField: ширина поля
'   — ReadOnly: необов’язково. True — тільки для читання.
Sub FieldTemplate(oDialogModel As Object, _
                   NamePrefix  As String, _
                   LabelText   As String, _
                   PositionX   As Integer, _
                   PositionY   As Integer, _
                   vText       As String, _
                   WidthLabel  As Integer, _
                   WidthField  As Integer, _
          Optional ReadOnly    As Variant)
                  
    Dim bReadOnly As Boolean
    If IsMissing(ReadOnly) Then
        bReadOnly = False
    Else
        bReadOnly = ReadOnly
    End If

    ' ==== Мітка ====
    Dim oLabel As Object
    oLabel           = oDialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    oLabel.Name      = NamePrefix & "Label"
    oLabel.Label     = LabelText
    oLabel.PositionX = PositionX
    oLabel.PositionY = PositionY - 10 ' Мітка вище за поле
    oLabel.Width     = WidthLabel
    oLabel.Height    = 10
    oLabel.TextColor = pRGB(TEXT_COLOR)
    oDialogModel.insertByName(oLabel.Name, oLabel)
    
    ' ==== Поле ====
    Dim oField As Object
    oField           = oDialogModel.createInstance("com.sun.star.awt.UnoControlEditModel")
    oField.Name      = NamePrefix & "Field"
    oField.PositionX = PositionX
    oField.PositionY = PositionY
    oField.Width     = WidthField
    oField.Height    = 15
    oField.Text      = vText
    oField.ReadOnly  = bReadOnly
    ' MsgBox oField.Name
    oDialogModel.insertByName(oField.Name, oField)
End Sub

' =====================================================
' === Процедура ComboBoxTemplate ======================
' =====================================================
' → Додає на діалог мітку та комбінований список (ComboBox) під нею.
' → Параметри:
'   — NamePrefix: префікс для імені (Label та Combo)
'   — LabelText: текст мітки
'   — PositionX, PositionY: координати верхнього лівого кута ComboBox
'   — vText: значення за замовчуванням
'   — WidthLabel: ширина мітки
'   — WidthCombo: ширина ComboBox
'   — ListOfPlaces: рядок зі значеннями через ;
Sub ComboBoxTemplate(oDialogModel As Object, _
                     NamePrefix   As String, _
                     LabelText    As String, _
                     PositionX    As Integer, _
                     PositionY    As Integer, _
                     vText        As String, _
                     WidthLabel   As Integer, _
                     WidthCombo   As Integer, _
                     ListOfPlaces As String)

    ' ==== Мітка ====
    Dim oLabel As Object
    oLabel = oDialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    oLabel.Name = NamePrefix & "Label"
    oLabel.Label = LabelText
    oLabel.PositionX = PositionX
    oLabel.PositionY = PositionY - 10
    oLabel.Width = WidthLabel
    oLabel.Height = 10
    oLabel.TextColor = pRGB(TEXT_COLOR)
    oDialogModel.insertByName(oLabel.Name, oLabel)

    ' ==== ComboBox ====
    Dim oCombo As Object
    oCombo = oDialogModel.createInstance("com.sun.star.awt.UnoControlComboBoxModel")
    oCombo.Name = NamePrefix & "Combo"
    oCombo.PositionX = PositionX
    oCombo.PositionY = PositionY
    oCombo.Width = WidthCombo
    oCombo.Height = 15
    oCombo.Text = vText
    oCombo.Dropdown = True
    oCombo.StringItemList = Split(ListOfPlaces, ";")
    ' MsgBox oCombo.Name
    oDialogModel.insertByName(oCombo.Name, oCombo)
End Sub

' =====================================================
' === Процедура AddButton =============================
' =====================================================
' → Додає кнопку на діалогову форму.
' → Налаштовує її позицію, розміри, підпис та тип кнопки.
' → За замовчуванням кнопка Standard (не закриває діалог).
Sub AddButton(oDialogModel As Object, _ 
			  sName        As String, _
			  Label        As String, _
              PositionX    As Integer, _ 
              PositionY    As Integer, _
              Width        As Integer, _
              Height       As Integer, _
     Optional PushType     As Variant)
               
	Dim iPushType As Integer
	
    If IsMissing(PushType) Then
        iPushType = 0 ' За замовчуванням — Standard (не закриває вікно)
    Else
        iPushType = PushType
    End If
    
    Dim oButton As Object
    oButton = oDialogModel.createInstance("com.sun.star.awt.UnoControlButtonModel")
    oButton.Name = sName
    oButton.Label = Label
    oButton.PositionX = PositionX
    oButton.PositionY = PositionY
    oButton.Width = Width
    oButton.Height = Height
    oButton.PushButtonType = iPushType
    oDialogModel.insertByName(sName, oButton)
End Sub

' =====================================================
' === Процедура OptionGroupTemplate ===================
' =====================================================
' → Додає на діалог два радіокнопки з мітками над ними.
' → Параметри:
'   — NamePrefix: префікс для імен
'   — LabelLeft, LabelRight: тексти міток над кнопками
'   — PositionX, PositionY: координати лівої кнопки
'   — WidthEach: ширина кожної кнопки
'   — DefaultLeft: True, якщо ліва кнопка за замовчуванням вибрана
' =====================================================
Sub OptionGroupTemplate(oDialogModel As Object, _
                        NamePrefix   As String, _
                        LabelLeft    As String, _
                        LabelRight   As String, _
                        PositionX    As Integer, _
                        PositionY    As Integer, _
                        WidthEach    As Integer, _
               Optional DefaultLeft  As Boolean)

    If IsMissing(DefaultLeft) Then DefaultLeft = True

    ' ==== Ліва мітка ====
    Dim oLabelLeft As Object
    oLabelLeft = oDialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    oLabelLeft.Name = NamePrefix & "LeftLabel"
    oLabelLeft.Label = LabelLeft
    oLabelLeft.PositionX = PositionX
    oLabelLeft.PositionY = PositionY - 10
    oLabelLeft.Width = WidthEach
    oLabelLeft.Height = 10
    oLabelLeft.TextColor = pRGB(TEXT_COLOR)
    oDialogModel.insertByName(oLabelLeft.Name, oLabelLeft)

    ' ==== Права мітка ====
    Dim oLabelRight As Object
    oLabelRight = oDialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    oLabelRight.Name = NamePrefix & "RightLabel"
    oLabelRight.Label = LabelRight
    oLabelRight.PositionX = PositionX + WidthEach + 10
    oLabelRight.PositionY = PositionY - 10
    oLabelRight.Width = WidthEach
    oLabelRight.Height = 10
    oLabelRight.TextColor = pRGB(TEXT_COLOR)
    oDialogModel.insertByName(oLabelRight.Name, oLabelRight)

    ' ==== Ліва радіокнопка ====
    Dim oOptionLeft As Object
    oOptionLeft = oDialogModel.createInstance("com.sun.star.awt.UnoControlRadioButtonModel")
    oOptionLeft.Name = NamePrefix & "Left"
    oOptionLeft.PositionX = PositionX + 8
    oOptionLeft.PositionY = PositionY
    oOptionLeft.Width = WidthEach
    oOptionLeft.Height = 12
    oOptionLeft.State = IIf(DefaultLeft, True, False)
    oDialogModel.insertByName(oOptionLeft.Name, oOptionLeft)

    ' ==== Права радіокнопка ====
    Dim oOptionRight As Object
    oOptionRight = oDialogModel.createInstance("com.sun.star.awt.UnoControlRadioButtonModel")
    oOptionRight.Name = NamePrefix & "Right"
    oOptionRight.PositionX = PositionX + WidthEach + 16
    oOptionRight.PositionY = PositionY
    oOptionRight.Width = WidthEach
    oOptionRight.Height = 12
    oOptionRight.State = IIf(DefaultLeft, False, True)
    oDialogModel.insertByName(oOptionRight.Name, oOptionRight)
End Sub

' =====================================================
' === Процедура AddDateField ==========================
' =====================================================
' → Додає на діалог поле введення дати з календарем (DateField).
' → Параметри:
'   — m: модель діалогу
'   — n: ім’я елемента
'   — x, y: координати верхнього лівого кута
'   — w, h: ширина й висота
' → Вмикає Dropdown, щоб можна було вибрати дату з календаря.
Sub AddDateField(m, n, x, y, w, h, Optional dD As Variant)
    Dim o As Object
    o = m.createInstance("com.sun.star.awt.UnoControlDateFieldModel")
    o.Name       = n & "Date"
    o.PositionX  = x
    o.PositionY  = y
    o.Width      = w
    o.Height     = h
    o.Dropdown   = True
    o.DateFormat = 2
    
    If Not isMissing(dD) Then
        Dim d As Object
        d        = CreateUnoStruct("com.sun.star.util.Date")
        d.Year   = Year(dD)
        d.Month  = Month(dD)
        d.Day    = Day(dD)
        o.Date   = d
    End If
    ' MsgBox o.Name
    m.insertByName(o.Name, o)
End Sub

' =====================================================
' === Процедура AddTimeField ==========================
' =====================================================
' → Додає на діалог поле введення часу з поточним значенням.
' → Параметри:
'   — m: модель діалогу
'   — n: ім’я елемента
'   — x, y: координати верхнього лівого кута
'   — w, h: ширина й висота
' → Встановлює TimeFormat = 1 (HH:MM:SS) та Spin = True.
' → Значення часу встановлюється на поточний системний час.
Sub AddTimeField(m, n, x, y, w, h)
    Dim o As Object
    Dim t As New com.sun.star.util.Time
    
    t.Hours = Hour(Now)
    t.Minutes = Minute(Now)
    t.Seconds = Second(Now)
    t.IsUTC = False
    
    o = m.createInstance("com.sun.star.awt.UnoControlTimeFieldModel")
    o.Name = n & "Time" : o.PositionX = x : o.PositionY = y : o.Width = w : o.Height = h
    o.TimeFormat = 1
    o.Spin = True  
    o.Time = t
    
    m.insertByName(o.Name, o)
End Sub

' =====================================================
' === Процедура AddLabel ==============================
' =====================================================
' → Додає на діалог мітку (Label) з заданим текстом.
' → Параметри:
'   — m: модель діалогу
'   — n: ім’я елемента
'   — txt: текст мітки
'   — x, y: координати верхнього лівого кута
'   — w, h: ширина й висота
' → Використовується для підписів до полів введення та інших елементів.
Sub AddLabel(m, n, txt, x, y, w, h)
    Dim o
    o = m.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    o.Name = n & "Label" : o.PositionX = x : o.PositionY = y : o.Width = w : o.Height = h
    o.Label = txt
    o.TextColor = pRGB(TEXT_COLOR)
    m.insertByName(o.Name, o)
End Sub

Sub AddLabelFont(m, n, x, y, w, h, txt)
    Dim o As Object
    o = m.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    
    With o
        .Name = n & "Label"
        .PositionX  = x
        .PositionY  = y
        .Width      = w
        .Height     = h
        .Label      = txt
        .Multiline  = True
        .FontHeight = 12  ' розмір пт
        .FontWeight = 150 ' 100 normal; 150, 200 — bold
        .FontName   = "Arial"
        .TextColor  = RGB(255, 215, 0)
        
    End With
    
    m.insertByName(o.Name, o)
End Sub

' =====================================================
' === Приклади інших полів ============================
' =====================================================

' =====================================================
' === Процедура DemoControls ==========================
' =====================================================
' → Демонструє діалог із прикладами всіх стандартних елементів керування.
' → Створює діалог, додає різні елементи та виконує його.
Sub DemoControls()
    Dim oDialog As Object, oDialogModel As Object
    Dim sTitle As String
    sTitle = "Діалог з усіма елементами"

    oDialog = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    oDialogModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    oDialog.setModel(oDialogModel)

    ' ==== Параметри діалогу ====
    With oDialogModel
        .PositionX = 100
        .PositionY = 100
        .Width     = 160
        .Height    = 400
        .Title     = sTitle
    End With

    Dim x As Long, y As Long, w As Long, h As Long
    x = 10 : y = 10 : w = 50 : h = 12

    ' ==== Елементи ====
    AddLabel(oDialogModel, "LabelDate", x, y, w, h, "Date")
    y = y + 10
    AddDateField(oDialogModel, "date", x, y, w, h)
    y = y + 20
    AddLabel(oDialogModel, "LabelTime", x, y, w, h, "Time")
    y = y + 10
    AddTimeField(oDialogModel, "time", x, y, 70, h)
    y = y + 20
    AddLabel(oDialogModel, "LabelEdit", x, y, w, h, "Edit")    
    y = y + 10 
    AddEdit(oDialogModel, "Edit", x, y, 70, h)
    oDialog.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)
    
    ' ==== Групова рамка з трьома елементами ====
    Dim gX As Long, gY As Long, gW As Long, gH As Long
    gX = x 
    gY = y + 15 
    gW = 140 
    gH = 65

    AddGroupBox(oDialogModel, "GroupBoxDemo", gX, gY, gW, gH, "Група")
    
    ' Елементи всередині групи
    Dim innerX As Long, innerY As Long
    innerX = gX + 10 
    innerY = gY + 15
    
    AddPatternField(oDialogModel, "PatternInGroup", innerX, innerY, 60, h)
    innerY = innerY + 15
    AddCheckBox(oDialogModel, "CheckInGroup", innerX, innerY, 60, h, "Чек")
    innerY = innerY + 15
    AddEdit(oDialogModel, "EditInGroup", innerX, innerY, 60, h)
    
    oDialog.execute()
    oDialog.dispose()
End Sub

' =====================================================
' === Процедура AddEditTemplate =======================
' =====================================================
' → Додає на діалог текстове поле для введення (Edit).
' → Використовується для введення рядків тексту.
Sub AddEditTemplate(m, n, x, y, w, h, t, vs)
    Dim o
    o = m.createInstance("com.sun.star.awt.UnoControlEditModel")
    
    With o
        .Name      = n & "Edit"    ' ім'я елемента
        .MultiLine = True          ' багаторядковий режим
        .ReadOnly  = True          ' тільки для читання
        .VScroll   = vs            ' вертикальний скрол
        .HScroll   = False         ' горизонтальний скрол
        .Text      = t             ' текст повідомлення
        .Width     = w             ' ширина текстового поля
        .Height    = h             ' висота текстового поля
        .PositionX = x             ' відступ зліва
        .PositionY = y             ' відступ зверху
        .TextColor = RGB(255, 255, 255)
        .BackgroundColor = RGB(22, 11, 172)
    End With
    m.insertByName(o.Name, o)
End Sub

' =====================================================
' === Процедура AddCheckBox ===========================
' =====================================================
' → Додає на діалог прапорець (CheckBox).
' → Використовується для вибору так/ні.
Sub AddCheckBox(m, n, x, y, w, h, txt)
    Dim o
    o = m.createInstance("com.sun.star.awt.UnoControlCheckBoxModel")
    o.Name = n : o.PositionX = x : o.PositionY = y : o.Width = w : o.Height = h
    o.Label = txt
    m.insertByName(n, o)
End Sub

' =====================================================
' === Процедура AddRadioButton ========================
' =====================================================
' → Додає на діалог перемикач (RadioButton).
' → Використовується для вибору одного варіанту з групи.
Sub AddRadioButton(m, n, x, y, w, h, txt)
    Dim o
    o = m.createInstance("com.sun.star.awt.UnoControlRadioButtonModel")
    o.Name = n : o.PositionX = x : o.PositionY = y : o.Width = w : o.Height = h
    o.Label = txt
    m.insertByName(n, o)
End Sub

' =====================================================
' === Процедура AddListBox ============================
' =====================================================
' → Додає на діалог список (ListBox) зі значеннями.
' → Використовується для вибору зі списку.
Sub AddListBox(m, n, x, y, w, h, arr)
    Dim o
    o = m.createInstance("com.sun.star.awt.UnoControlListBoxModel")
    o.Name = n : o.PositionX = x : o.PositionY = y : o.Width = w : o.Height = h
    o.StringItemList = arr
    m.insertByName(n, o)
End Sub

' =====================================================
' === Процедура AddComboBox ===========================
' =====================================================
' → Додає на діалог комбінований список (ComboBox) зі значеннями.
' → Дозволяє вибір або введення власного значення.
Sub AddComboBox(m, n, x, y, w, h, arr)
    Dim o
    o = m.createInstance("com.sun.star.awt.UnoControlComboBoxModel")
    o.Name = n : o.PositionX = x : o.PositionY = y : o.Width = w : o.Height = h
    o.StringItemList = arr
    m.insertByName(n, o)
End Sub

' =====================================================
' === Процедура AddButtonDemo =========================
' =====================================================
' → Додає на діалог кнопку з підписом.
' → Використовується для запуску дії при натисканні.
Sub AddButtonDemo(m, n, x, y, w, h, txt)
    Dim o
    o = m.createInstance("com.sun.star.awt.UnoControlButtonModel")
    o.Name = n : o.PositionX = x : o.PositionY = y : o.Width = w : o.Height = h
    o.Label = txt
    m.insertByName(n, o)
End Sub

' =====================================================
' === Процедура AddProgressBar ========================
' =====================================================
' → Додає на діалог індикатор виконання (ProgressBar).
' → Використовується для відображення прогресу.
Sub AddProgressBar(m, n, x, y, w, h)
    Dim o
    o = m.createInstance("com.sun.star.awt.UnoControlProgressBarModel")
    o.Name = n : o.PositionX = x : o.PositionY = y : o.Width = w : o.Height = h
    m.insertByName(n, o)
End Sub

' =====================================================
' === Процедура AddNumericField =======================
' =====================================================
' → Додає на діалог поле для введення чисел (NumericField).
' → Дозволяє вводити тільки числа.
Sub AddNumericField(m, n, x, y, w, h)
    Dim o
    o = m.createInstance("com.sun.star.awt.UnoControlNumericFieldModel")
    o.Name = n : o.PositionX = x : o.PositionY = y : o.Width = w : o.Height = h
    m.insertByName(n, o)
End Sub

' =====================================================
' === Процедура AddCurrencyField ======================
' =====================================================
' → Додає на діалог поле для введення сум (CurrencyField).
' → Відображає й дозволяє вводити валютні значення.
Sub AddCurrencyField(m, n, x, y, w, h)
    Dim o
    o = m.createInstance("com.sun.star.awt.UnoControlCurrencyFieldModel")
    o.Name = n : o.PositionX = x : o.PositionY = y : o.Width = w : o.Height = h
    m.insertByName(n, o)
End Sub

' =====================================================
' === Процедура AddPatternField =======================
' =====================================================
' → Додає на діалог поле з маскою (PatternField).
' → Використовується для вводу значень за заданим шаблоном.
Sub AddPatternField(m, n, x, y, w, h)
    Dim o
    o = m.createInstance("com.sun.star.awt.UnoControlPatternFieldModel")
    o.Name = n : o.PositionX = x : o.PositionY = y : o.Width = w : o.Height = h
    o.EditMask = "+38 (NNN) NNN NN NN"
    o.LiteralMask = "+38 (___) ___ __ __"
    m.insertByName(n, o)
End Sub

' =====================================================
' === Процедура AddFormattedField =====================
' =====================================================
' → Додає на діалог форматоване поле (FormattedField).
' → Дозволяє вводити значення з форматуванням.
Sub AddFormattedField(m, n, x, y, w, h)
    Dim o
    o = m.createInstance("com.sun.star.awt.UnoControlFormattedFieldModel")
    o.Name = n : o.PositionX = x : o.PositionY = y : o.Width = w : o.Height = h
    m.insertByName(n, o)
End Sub

' =====================================================
' === Процедура AddFileControl ========================
' =====================================================
' → Додає на діалог елемент вибору файлу (FileControl).
' → Використовується для вибору файлів із файлової системи.
Sub AddFileControl(m, n, x, y, w, h)
    Dim o
    o = m.createInstance("com.sun.star.awt.UnoControlFileControlModel")
    o.Name = n : o.PositionX = x : o.PositionY = y : o.Width = w : o.Height = h
    m.insertByName(n, o)
End Sub

' =====================================================
' === Процедура AddGroupBox ===========================
' =====================================================
' → Додає на діалог групову рамку (GroupBox) з підписом.
' → Використовується для візуального об’єднання елементів.
Sub AddGroupBox(m, n, x, y, w, h, txt)
    Dim o
    o = m.createInstance("com.sun.star.awt.UnoControlGroupBoxModel")
    o.Name = n : o.PositionX = x : o.PositionY = y : o.Width = w : o.Height = h
    o.Label = txt
    o.TextColor = pRGB(TEXT_COLOR)
    m.insertByName(n, o)
End Sub

' =====================================================
' === Процедура AddImage ==============================
' =====================================================
' → Додає на діалог елемент відображення зображення (ImageControl).
' → Використовується для показу картинок або іконок.
Sub AddImage(m, n, x, y, w, h)
    Dim o
    o = m.createInstance("com.sun.star.awt.UnoControlImageControlModel")
    o.Name = n : o.PositionX = x : o.PositionY = y : o.Width = w : o.Height = h
    m.insertByName(n, o)
End Sub

